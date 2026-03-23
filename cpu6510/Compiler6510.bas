Option Explicit
'================================================================================
' Module:       Compiler6510
' Purpose:      Assembles 6510 source from the CPU sheet program grid into
'               machine bytes, writes them to gMemory, and updates the
'               MemoryTable display - exactly as Assemble8080_ToMachine does
'               for the 8080.
'
' Entry point:  Assemble6510_ToMachine  (called by AssembleForCPUMode in
'               modGlobals when CPUMode = "6510")
'
' SHARED with Compiler8080 (no changes needed there):
'   BuildOpcodeMap, EmitBytesToMemoryTable, RefreshMemoryTable,
'   cmpResolveValue16, cmpRaiseAsmError, StripComment, cmpStringToBytesVariant,
'   usrMemReset, ResetAddressList, gMemory
'
' PSEUDO-OPS supported:
'   ORG  nnnn   set program counter
'   EQU         constant already resolved by clsLabels; skip
'   END         program terminator; no bytes emitted
'   HLT / STP   halt; treated as terminator, no bytes emitted
'   DB   ...    define bytes: string literal OR comma-separated hex/decimal
'   DS   n      define n bytes of $00  (n = hex or decimal)
'
' ADDRESSING MODE SYNTAX in the OP1 column:
'   #nn          Immediate          LDA #4A
'   nn           Zero page          LDA 42
'   nn,X         Zero page,X        LDA 42,X
'   nn,Y         Zero page,Y        LDX 42,Y
'   nnnn         Absolute           LDA C000
'   nnnn,X       Absolute,X         LDA C000,X
'   nnnn,Y       Absolute,Y         LDA C000,Y
'   (nnnn)       Indirect (JMP)     JMP (C000)
'   (nn,X)       Indexed indirect   LDA (42,X)
'   (nn),Y       Indirect indexed   LDA (42),Y
'   A or ""      Accumulator/Implied ASL A  /  NOP
'   LABELNAME    resolved to address for any mode above
'
' ERROR HANDLING:
'   The Error cell is cleared at entry so stale errors from previous runs
'   do not interfere. On Error GoTo Asm6510Error catches runtime errors.
'   All deliberate error paths call cmpRaiseAsmError then GoTo Asm6510Done
'   so ScreenUpdating is always restored before exit.
'================================================================================

Private Const EMPTY_RUN_LIMIT  As Long   = 8
Private Const BRANCH_MNEMS     As String = "|BCC|BCS|BEQ|BMI|BNE|BPL|BVC|BVS|"

'================================================================================
' Assemble6510_ToMachine
'================================================================================
Public Sub Assemble6510_ToMachine()

    Dim wsCPU As Worksheet
    Dim wsOp  As Worksheet
    Set wsCPU = ThisWorkbook.Worksheets("CPU")

    '--- Locate opcode sheet before touching ScreenUpdating -------------------
    On Error Resume Next
    Set wsOp = ThisWorkbook.Worksheets("6510 Op to Hex")
    On Error GoTo 0
    If wsOp Is Nothing Then
        cmpRaiseAsmError wsCPU, "Sheet '6510 Op to Hex' not found - import it first", _
                         wsCPU.Range("Line0")
        Exit Sub
    End If

    '--- All errors from here route through Asm6510Error ----------------------
    On Error GoTo Asm6510Error

    '--- Clear stale error state from any previous run ------------------------
    wsCPU.Range("Error").value = 0
    wsCPU.Range("errMessage").value = ""

    '--- Initialise memory and labels -----------------------------------------
    usrMemReset
    ResetAddressList
    Dim myMemory As clsMemory
    Set myMemory = gMemory

    '--- Read grid offsets (same named ranges as 8080) ------------------------
    Dim ofsLabel   As Long: ofsLabel  = CLng(wsCPU.Range("ofs_label").value)
    Dim ofsOpcode  As Long: ofsOpcode = CLng(wsCPU.Range("ofs_opcode").value)
    Dim ofsOp1     As Long: ofsOp1    = CLng(wsCPU.Range("ofs_op1").value)
    Dim ofsOp2     As Long: ofsOp2    = CLng(wsCPU.Range("ofs_op2").value)

    ' gMemTableCol is Private to Compiler8080 - resolve locally
    Dim memTableCol As Long
    memTableCol = CLng(wsCPU.Range("MemoryTableAddress").Column)

    Dim memStartDec As Long: memStartDec = usrHexToDec(CStr(wsCPU.Range("MemStart").value))
    Dim memSize     As Long: memSize     = usrHexToDec(CStr(wsCPU.Range("MemSize").value))

    '--- Build opcode dictionary (shared helper from Compiler8080) ------------
    Dim opMap As Object
    Set opMap = BuildOpcodeMap(wsOp)

    '--- Build label map from the program grid --------------------------------
    Dim lbls As clsLabels
    Set lbls = New clsLabels

    '--- MemoryTable range anchor ---------------------------------------------
    Dim rngMem As Range
    Set rngMem = wsCPU.Range("MemoryTable")

    '--- Recalculate so Dec Line / Bytes formulas are current ----------------
    Application.Calculate
    Application.ScreenUpdating = False

    '--- Clear prior compiled output -----------------------------------------
    rngMem.ClearContents
    wsCPU.Range(wsCPU.Cells(rngMem.Row, memTableCol), _
                wsCPU.Cells(rngMem.Row + rngMem.Rows.Count - 1, memTableCol)).ClearContents

    '--- Walk the program grid -----------------------------------------------
    Dim base As Range
    Set base = wsCPU.Range("Line0")

    Dim pc       As Long: pc       = memStartDec
    Dim emptyRun As Long: emptyRun = 0
    Dim r        As Long

    For r = 0 To memSize

        Dim cur As Range
        Set cur = base.Offset(r, 0)

        Dim labelName As String
        Dim opcode    As String
        Dim op1       As String
        Dim op2       As String

        labelName = UCase$(Trim$(CStr(cur.Offset(0, ofsLabel).value)))
        opcode    = UCase$(Trim$(CStr(cur.Offset(0, ofsOpcode).value)))
        op1       = StripComment(Trim$(CStr(cur.Offset(0, ofsOp1).value)))
        op2       = StripComment(Trim$(CStr(cur.Offset(0, ofsOp2).value)))

        ' Skip comment-only rows
        If labelName = ";" Then GoTo NextRow6510

        ' Empty-row stop condition
        If labelName = "" And opcode = "" And op1 = "" And op2 = "" Then
            emptyRun = emptyRun + 1
            If emptyRun >= EMPTY_RUN_LIMIT Then Exit For
            GoTo NextRow6510
        Else
            emptyRun = 0
        End If

        Dim bytesOut As Variant
        bytesOut = Empty

        Select Case opcode

            Case ""
                GoTo NextRow6510

            '--- Assembler directives ----------------------------------------

            Case "ORG"
                If op1 = "" Then
                    cmpRaiseAsmError wsCPU, "ORG: missing operand", cur
                    GoTo Asm6510Done
                End If
                pc = cmpResolveValue16(lbls, UCase$(op1))
                GoTo NextRow6510

            Case "EQU"
                GoTo NextRow6510

            Case "END", "HLT", "STP"
                GoTo NextRow6510

            Case "DB"
                Dim dbVal As String
                dbVal = CStr(cur.Offset(0, ofsOp1).value)
                If Len(Trim$(dbVal)) = 0 Then
                    cmpRaiseAsmError wsCPU, "DB: missing operand", cur
                    GoTo Asm6510Done
                End If
                bytesOut = Hlp6510DBToBytes(dbVal)
                If IsEmpty(bytesOut) Then
                    cmpRaiseAsmError wsCPU, "DB: invalid operand: " & dbVal, cur
                    GoTo Asm6510Done
                End If

            Case "DS"
                Dim dsVal As String
                dsVal = Trim$(CStr(cur.Offset(0, ofsOp1).value))
                If Not usrValIsHexString(dsVal) And Not IsNumeric(dsVal) Then
                    cmpRaiseAsmError wsCPU, "DS: invalid length: " & dsVal, cur
                    GoTo Asm6510Done
                End If
                Dim dsLen As Long
                dsLen = IIf(usrValIsHexString(dsVal), usrHexToDec(dsVal), CLng(dsVal))
                If dsLen < 1 Then GoTo NextRow6510
                Dim dsArr() As Variant
                ReDim dsArr(0 To dsLen - 1)
                Dim di As Long
                For di = 0 To dsLen - 1
                    dsArr(di) = 0
                Next di
                bytesOut = dsArr

            '--- Normal 6510 instructions ------------------------------------
            Case Else

                Dim op1Resolved As String
                op1Resolved = Trim$(UCase$(op1))
                Dim lbl As clsLabelRecord
                Set lbl = lbls.GetLabel(op1Resolved)
                If Not lbl Is Nothing Then
                    op1Resolved = lbl.clsAddressHex
                End If

                Dim info As Variant
                info = Hlp6510FindOpcode(opMap, opcode, op1Resolved, UCase$(Trim$(op2)))
                If IsEmpty(info) Then
                    cmpRaiseAsmError wsCPU, _
                        "6510: Unknown instruction: " & opcode & " " & op1 & " " & op2, cur
                    GoTo Asm6510Done
                End If

                bytesOut = Hlp6510EncodeInstruction( _
                               lbls, CByte(info(0) And &HFF&), CLng(info(1)), _
                               CStr(info(2)), op1Resolved, opcode, pc, cur, wsCPU)
                If IsEmpty(bytesOut) Then GoTo Asm6510Done

        End Select

        ' Emit bytes to MemoryTable and gMemory
        If Not IsEmpty(bytesOut) Then
            EmitBytesToMemoryTable wsCPU, rngMem, memStartDec, pc, bytesOut, myMemory, memTableCol
            pc = pc + (UBound(bytesOut) - LBound(bytesOut) + 1)
        End If

NextRow6510:
    Next r

    '--- Success --------------------------------------------------------------
    Application.ScreenUpdating = True
    wsCPU.Range("errMessage").value = "Assemble complete (6510)"
    RefreshMemoryTable
    Exit Sub

    '--- Expected error paths land here (cmpRaiseAsmError already called) -----
Asm6510Done:
    Application.ScreenUpdating = True
    RefreshMemoryTable
    Exit Sub

    '--- Unhandled runtime errors ---------------------------------------------
Asm6510Error:
    Application.ScreenUpdating = True
    cmpRaiseAsmError wsCPU, _
        "6510 runtime error " & Err.Number & ": " & Err.Description, _
        wsCPU.Range("Line0")

End Sub

'================================================================================
' Hlp6510FindOpcode
'================================================================================
Private Function Hlp6510FindOpcode(ByVal opMap  As Object, _
                                   ByVal mnem   As String, _
                                   ByVal op1    As String, _
                                   ByVal op2    As String) As Variant
    Dim spec As String
    spec = Hlp6510ClassifyOp1(op1)

    Dim key As String

    key = mnem & "|" & spec & "|"
    If opMap.Exists(key) Then Hlp6510FindOpcode = opMap(key): Exit Function

    ' ZP fallback to ADDRESS (some instructions have no zero-page form)
    If spec = "ZP" Then
        key = mnem & "|ADDRESS|"
        If opMap.Exists(key) Then Hlp6510FindOpcode = opMap(key): Exit Function
    End If

    ' IMP/ACC cross-fallback
    If spec = "IMP" Or spec = "ACC" Then
        key = mnem & "|IMP|"
        If opMap.Exists(key) Then Hlp6510FindOpcode = opMap(key): Exit Function
        key = mnem & "|ACC|"
        If opMap.Exists(key) Then Hlp6510FindOpcode = opMap(key): Exit Function
    End If

    Hlp6510FindOpcode = Empty
End Function

'================================================================================
' Hlp6510ClassifyOp1
'================================================================================
Private Function Hlp6510ClassifyOp1(ByVal op1 As String) As String
    op1 = Trim$(UCase$(op1))

    If op1 = "" Or op1 = "A" Then
        Hlp6510ClassifyOp1 = "IMP": Exit Function
    End If

    If Left$(op1, 1) = "#" Then
        Hlp6510ClassifyOp1 = "BYTE": Exit Function
    End If

    If Left$(op1, 1) = "(" And InStr(op1, ",X)") > 0 Then
        Hlp6510ClassifyOp1 = "IND_X": Exit Function
    End If

    If Left$(op1, 1) = "(" And Right$(op1, 2) = "),Y" Then
        Hlp6510ClassifyOp1 = "IND_Y": Exit Function
    End If

    If Left$(op1, 1) = "(" And Right$(op1, 1) = ")" Then
        Hlp6510ClassifyOp1 = "IND": Exit Function
    End If

    Dim commaPos As Long
    commaPos = InStr(op1, ",")
    If commaPos > 0 Then
        Dim baseToken As String: baseToken = Left$(op1, commaPos - 1)
        Dim indexReg  As String: indexReg  = Mid$(op1, commaPos + 1)
        Dim zpBase    As Boolean: zpBase   = Hlp6510IsZeroPage(baseToken)
        Select Case indexReg
            Case "X": Hlp6510ClassifyOp1 = IIf(zpBase, "ZP_X", "ABS_X")
            Case "Y": Hlp6510ClassifyOp1 = IIf(zpBase, "ZP_Y", "ABS_Y")
            Case Else: Hlp6510ClassifyOp1 = "ADDRESS"
        End Select
        Exit Function
    End If

    If Hlp6510IsZeroPage(op1) Then
        Hlp6510ClassifyOp1 = "ZP"
    Else
        Hlp6510ClassifyOp1 = "ADDRESS"
    End If
End Function

'================================================================================
' Hlp6510IsZeroPage
'================================================================================
Private Function Hlp6510IsZeroPage(ByVal token As String) As Boolean
    token = Trim$(UCase$(token))
    If Right$(token, 1) = "H" Then token = Left$(token, Len(token) - 1)
    If usrValIsHexString(token) Then
        Hlp6510IsZeroPage = (usrHexToDec(token) <= &HFF&)
    Else
        Hlp6510IsZeroPage = False
    End If
End Function

'================================================================================
' Hlp6510EncodeInstruction
' Returns Empty only if a genuine assembly error occurred (already reported).
' NOTE: does NOT check wsCPU.Range("Error") - that cell may hold stale values
' from previous runs. Error signalling is via cmpRaiseAsmError + return Empty.
'================================================================================
Private Function Hlp6510EncodeInstruction(ByVal lbls    As clsLabels, _
                                          ByVal opByte  As Byte, _
                                          ByVal nBytes  As Long, _
                                          ByVal spec    As String, _
                                          ByVal op1     As String, _
                                          ByVal mnem    As String, _
                                          ByVal pc      As Long, _
                                          ByVal curRow  As Range, _
                                          ByVal wsCPU   As Worksheet) As Variant
    Dim out() As Variant

    Select Case nBytes

        Case 1
            ReDim out(0 To 0)
            out(0) = opByte

        Case 2
            ReDim out(0 To 1)
            out(0) = opByte

            If Hlp6510IsBranch(mnem) Then
                ' Relative branch: signed offset from byte after the instruction
                Dim target As Long
                Dim targetOk As Boolean
                target = Hlp6510ResolveAddr(lbls, op1, mnem, wsCPU, curRow, targetOk)
                If Not targetOk Then
                    Hlp6510EncodeInstruction = Empty: Exit Function
                End If
                Dim relOffset As Long
                relOffset = target - (pc + 2)
                If relOffset < -128 Or relOffset > 127 Then
                    cmpRaiseAsmError wsCPU, _
                        mnem & ": branch out of range ($" & _
                        Right$("0000" & Hex$(target), 4) & " is " & relOffset & _
                        " bytes from $" & Right$("0000" & Hex$(pc), 4) & ")", curRow
                    Hlp6510EncodeInstruction = Empty: Exit Function
                End If
                out(1) = CByte(relOffset And &HFF&)
            Else
                ' 2-byte immediate
                Dim immStr As String
                immStr = Hlp6510StripOperandSyntax(op1, spec)
                Dim imm8 As Long
                Dim imm8Ok As Boolean
                imm8 = Hlp6510ResolveImm8(lbls, immStr, mnem, wsCPU, curRow, imm8Ok)
                If Not imm8Ok Then
                    Hlp6510EncodeInstruction = Empty: Exit Function
                End If
                out(1) = CByte(imm8 And &HFF&)
            End If

        Case 3
            ReDim out(0 To 2)
            out(0) = opByte
            Dim addrStr As String
            addrStr = Hlp6510StripOperandSyntax(op1, spec)
            Dim addr16 As Long
            Dim addr16Ok As Boolean
            addr16 = Hlp6510ResolveAddr(lbls, addrStr, mnem, wsCPU, curRow, addr16Ok)
            If Not addr16Ok Then
                Hlp6510EncodeInstruction = Empty: Exit Function
            End If
            out(1) = CByte(addr16 And &HFF&)
            out(2) = CByte((addr16 \ 256&) And &HFF&)

        Case Else
            cmpRaiseAsmError wsCPU, _
                "6510: unsupported byte count " & nBytes & " for " & mnem, curRow
            Hlp6510EncodeInstruction = Empty: Exit Function

    End Select

    Hlp6510EncodeInstruction = out
End Function

'================================================================================
' Hlp6510StripOperandSyntax
'================================================================================
Private Function Hlp6510StripOperandSyntax(ByVal op1  As String, _
                                           ByVal spec As String) As String
    op1 = Trim$(UCase$(op1))

    Select Case spec
        Case "BYTE"
            Hlp6510StripOperandSyntax = IIf(Left$(op1, 1) = "#", Mid$(op1, 2), op1)

        Case "IND_X"
            Dim p1 As Long: p1 = InStr(op1, ",X)")
            Hlp6510StripOperandSyntax = Mid$(op1, 2, p1 - 2)

        Case "IND_Y"
            Dim p2 As Long: p2 = InStr(op1, "),Y")
            Hlp6510StripOperandSyntax = Mid$(op1, 2, p2 - 2)

        Case "IND"
            Hlp6510StripOperandSyntax = Mid$(op1, 2, Len(op1) - 2)

        Case "ZP_X", "ZP_Y", "ABS_X", "ABS_Y"
            Dim cp As Long: cp = InStr(op1, ",")
            Hlp6510StripOperandSyntax = IIf(cp > 0, Left$(op1, cp - 1), op1)

        Case Else
            Hlp6510StripOperandSyntax = op1
    End Select
End Function

'================================================================================
' Hlp6510ResolveAddr  -  label or hex -> 16-bit Long
' Returns 0 and sets ok=False if unresolvable (error already raised).
'================================================================================
Private Function Hlp6510ResolveAddr(ByVal lbls   As clsLabels, _
                                    ByVal token  As String, _
                                    ByVal opname As String, _
                                    ByVal wsCPU  As Worksheet, _
                                    ByVal curRow As Range, _
                                    ByRef  ok    As Boolean) As Long
    ok = True
    token = Trim$(UCase$(token))

    Dim rec As clsLabelRecord
    Set rec = lbls.GetLabel(token)
    If Not rec Is Nothing Then
        Hlp6510ResolveAddr = rec.clsAddress And &HFFFF&: Exit Function
    End If

    If usrValIsHexString(token) Then
        Hlp6510ResolveAddr = usrHexToDec(token) And &HFFFF&
    ElseIf IsNumeric(token) Then
        Hlp6510ResolveAddr = CLng(token) And &HFFFF&
    Else
        cmpRaiseAsmError wsCPU, opname & ": unresolved address: " & token, curRow
        ok = False
        Hlp6510ResolveAddr = 0
    End If
End Function

'================================================================================
' Hlp6510ResolveImm8  -  label or hex -> 8-bit Long
' Returns 0 and sets ok=False if unresolvable (error already raised).
'================================================================================
Private Function Hlp6510ResolveImm8(ByVal lbls   As clsLabels, _
                                    ByVal token  As String, _
                                    ByVal opname As String, _
                                    ByVal wsCPU  As Worksheet, _
                                    ByVal curRow As Range, _
                                    ByRef  ok    As Boolean) As Long
    ok = True
    token = Trim$(UCase$(token))

    Dim rec As clsLabelRecord
    Set rec = lbls.GetLabel(token)
    If Not rec Is Nothing Then
        Hlp6510ResolveImm8 = rec.clsAddress And &HFF&: Exit Function
    End If

    If usrValIsHexString(token) Then
        Hlp6510ResolveImm8 = usrHexToDec(token) And &HFF&
    ElseIf IsNumeric(token) Then
        Hlp6510ResolveImm8 = CLng(token) And &HFF&
    Else
        cmpRaiseAsmError wsCPU, opname & ": invalid 8-bit immediate: " & token, curRow
        ok = False
        Hlp6510ResolveImm8 = 0
    End If
End Function

'================================================================================
' Hlp6510IsBranch
'================================================================================
Private Function Hlp6510IsBranch(ByVal mnem As String) As Boolean
    Hlp6510IsBranch = (InStr(BRANCH_MNEMS, "|" & UCase$(mnem) & "|") > 0)
End Function

'================================================================================
' Hlp6510DBToBytes
'================================================================================
Private Function Hlp6510DBToBytes(ByVal s As String) As Variant
    s = Trim$(s)

    If Left$(s, 1) = Chr$(34) Or Left$(s, 1) = "'" Then
        Dim q     As String: q = Left$(s, 1)
        Dim inner As String
        inner = IIf(Right$(s, 1) = q And Len(s) >= 2, Mid$(s, 2, Len(s) - 2), Mid$(s, 2))
        Hlp6510DBToBytes = cmpStringToBytesVariant(inner)
        Exit Function
    End If

    Dim parts() As String
    parts = Split(s, ",")
    Dim out() As Variant
    ReDim out(0 To UBound(parts))
    Dim i As Long
    For i = 0 To UBound(parts)
        Dim tok As String: tok = Trim$(UCase$(parts(i)))
        If usrValIsHexString(tok) Then
            out(i) = usrHexToDec(tok) And &HFF&
        ElseIf IsNumeric(tok) Then
            out(i) = CLng(tok) And &HFF&
        Else
            Hlp6510DBToBytes = Empty: Exit Function
        End If
    Next i
    Hlp6510DBToBytes = out
End Function

'================================================================================
' ForceCompile6510
'================================================================================
Public Sub ForceCompile6510(Optional ByVal compile As Boolean = False)
    Dim wsCPU As Worksheet
    Set wsCPU = ThisWorkbook.Worksheets("CPU")

    If compile Then
        Assemble6510_ToMachine
        Exit Sub
    End If

    Dim ofs_opcode As Long: ofs_opcode = CLng(wsCPU.Range("ofs_opcode").value)
    Dim memSize    As Long: memSize    = usrHexToDec(CStr(wsCPU.Range("MemSize").value))
    Dim arrOpcode  As Variant
    arrOpcode = wsCPU.Range("Line0").Offset(0, ofs_opcode).Resize(memSize, 1).value

    Dim emptyRuns As Long: emptyRuns = 0
    Dim i         As Long
    For i = 1 To memSize
        Dim op As String: op = CStr(arrOpcode(i, 1))
        If op = "DB" Or op = "DS" Then compile = True
        If op = "" Then emptyRuns = emptyRuns + 1 Else emptyRuns = 0
        If compile Or emptyRuns >= EMPTY_RUN_LIMIT Then Exit For
    Next i

    If compile Then Assemble6510_ToMachine
End Sub
