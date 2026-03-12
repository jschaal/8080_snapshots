Attribute VB_Name = "Compiler"
Option Explicit

' --- Compile gating / caching ---
Private gCompileEnabled As Boolean
Private gCompiledValid As Boolean
Private gCompiledForResetStamp As Double   ' optional: timestamp of last reset compile
Private gLastCodeSignature As String       ' optional: detect edits without reset
' MemoryTable is 8 bytes wide: AB..AI
Private Const BYTES_PER_MEMROW As Long = 8
Private Const EMPTY_RUN_LIMIT As Long = 8
Private gMemTableCol As Long
Private pLabels As New clsLabels
' =============================================================================
' Assemble8080_ToMachine
' Compiles CPU listing into 8080 machine bytes and writes to:
'   - Column Z: base address for each 8-byte row
'   - MemoryTable (AB..AI): hex bytes
' Optionally writes "AA BB CC" into the listing Mem column and marks RowStat="C".
' =============================================================================
Public Sub Assemble8080_ToMachine()

    Dim wsCPU As Worksheet, wsOp As Worksheet
    Set wsCPU = ThisWorkbook.Worksheets("CPU")
    Set wsOp = ThisWorkbook.Worksheets("8080 Op to Hex") ' opcode mapping table
    Dim myMemory As clsMemory
    Dim lbl As clsLabelRecord
    
    usrMemReset
    ResetAddressList
    Set myMemory = gMemory
    ' Offsets into the CPU program grid (named ranges)
    Dim ofsLabel As Long, ofsOpcode As Long, ofsOp1 As Long, ofsOp2 As Long

    ofsLabel = CLng(wsCPU.Range("ofs_label").value)
    ofsOpcode = CLng(wsCPU.Range("ofs_opcode").value)
    ofsOp1 = CLng(wsCPU.Range("ofs_op1").value)
    ofsOp2 = CLng(wsCPU.Range("ofs_op2").value)
    gMemTableCol = CLng(wsCPU.Range("MemoryTableAddress").Column)
    

    ' Memory window
    Dim memStartDec As Long, memSize As Long
    memStartDec = usrHexToDec(CStr(wsCPU.Range("MemStart").value))
    memSize = usrHexToDec(CStr(wsCPU.Range("MemSize").value))

    ' Build opcode lookup dictionary from "8080 Op to Hex"
    Dim opMap As Object
    Set opMap = BuildOpcodeMap(wsOp)

    ' Labels/EQU/DB metadata (your existing class)
    Dim lbls As clsLabels
    Set lbls = New clsLabels

    ' MemoryTable byte cells (AB..AI) named range
    Dim rngMem As Range
    Set rngMem = wsCPU.Range("MemoryTable")

    Application.Calculate
    Application.ScreenUpdating = False

    ' Clear prior compiled bytes (only AB..AI)
    rngMem.ClearContents

    ' Also clear address column Z for the memory rows we might touch (overwrite is OK per you)
    wsCPU.Range(wsCPU.Cells(rngMem.row, gMemTableCol), _
                wsCPU.Cells(rngMem.row + rngMem.Rows.Count - 1, gMemTableCol)).ClearContents

    Dim base As Range
    Set base = wsCPU.Range("Line0") ' start of listing grid

    Dim pc As Long
    pc = memStartDec

    Dim r As Long, emptyRun As Long
    emptyRun = 0

    ' Walk listing rows (cap at memSize rows for safety)
    For r = 0 To memSize

        Dim cur As Range
        Set cur = base.offset(r, 0)

        Dim labelName As String, opcode As String, op1 As String, op2 As String
        labelName = UCase$(Trim$(CStr(cur.offset(0, ofsLabel).value)))
        opcode = UCase$(Trim$(CStr(cur.offset(0, ofsOpcode).value)))
        op1 = StripComment(Trim$(CStr(cur.offset(0, ofsOp1).value)))
        op2 = StripComment(Trim$(CStr(cur.offset(0, ofsOp2).value)))

        ' stop condition: consecutive empty program rows
        If (labelName = ";") Then
            GoTo NextRow
        End If
        
        If (labelName = "") And (opcode = "") And (op1 = "") And (op2 = "") Then
            emptyRun = emptyRun + 1
            If emptyRun >= EMPTY_RUN_LIMIT Then Exit For
            GoTo NextRow
        Else
            emptyRun = 0
        End If

        Set lbl = pLabels.GetLabel(op1)
        If Not lbl Is Nothing Then
            op1 = lbl.clsAddressHex ' use the label's hex value/address
        End If

        Dim bytesOut As Variant
        bytesOut = Empty

        Select Case opcode
            Case ""
                GoTo NextRow

            Case "ORG"
                If op1 = "" Then
                    RaiseAsmError wsCPU, "ORG missing operand", cur
                    Exit Sub
                End If
                pc = ResolveValue16(lbls, UCase$(op1))
                GoTo NextRow

            Case "EQU"
                ' In your model, EQU constants are captured during label parsing already.
                GoTo NextRow

            Case "DB"
                ' String literal only for now (raw OP1)
                If Len(CStr(cur.offset(0, ofsOp1).value)) = 0 Then
                    RaiseAsmError wsCPU, "DB missing string", cur
                    Exit Sub
                End If
                bytesOut = StringToBytesVariant(CStr(cur.offset(0, ofsOp1).value))

            Case "DS"
                Dim ds As Long
                Dim i As Long
                If Not usrValIsHexString(cur.offset(0, ofsOp1).value) Then
                     RaiseAsmError wsCPU, "DS Invalid DS Length", cur
                    Exit Sub
                End If
                ds = usrHexToDec(cur.offset(0, ofsOp1).value)
                Dim out() As Variant
                ReDim out(0 To ds - 1)
                For i = 1 To ds
                    out(i - 1) = 255
                Next i
                bytesOut = out
                    
            Case Else
                ' Normal instruction: lookup + encode (1/2/3 bytes)
                Dim info As Variant
                info = HlpFindOpcode(opMap, opcode, UCase$(op1), UCase$(op2))
                If IsEmpty(info) Then
                    RaiseAsmError wsCPU, "Unknown instruction form: " & opcode & " " & op1 & " " & op2, cur
                    Exit Sub
                End If

                bytesOut = HlpEncodeInstruction(lbls, CByte(info(0)), CLng(info(1)), _
                                             CStr(info(2)), CStr(info(3)), _
                                             UCase$(op1), UCase$(op2), opcode, cur, wsCPU)
        End Select

        ' Emit compiled bytes to MemoryTable using PC-based mapping (Z + AB..AI)
        If Not IsEmpty(bytesOut) Then
            EmitBytesToMemoryTable wsCPU, rngMem, memStartDec, pc, bytesOut, myMemory


            pc = pc + (UBound(bytesOut) - LBound(bytesOut) + 1)
        End If


NextRow:
    Next r
    Application.ScreenUpdating = True

    wsCPU.Range("errMessage").value = "Assemble complete"

End Sub

' =============================================================================
' Memory emit: PC -> row/col within AB..AI and base address into column Z
' =============================================================================
Private Sub EmitBytesToMemoryTable(ByVal wsCPU As Worksheet, _
 ByVal rngMem As Range, _
 ByVal memStartDec As Long, _
 ByVal startAddrDec As Long, _
 ByVal bytesOut As Variant, _
 ByRef myMem As clsMemory)

    Dim i As Long
    Dim baseAddr As Long
    Dim lastTargetRow As Long

    lastTargetRow = 0
    baseAddr = memStartDec

    For i = LBound(bytesOut) To UBound(bytesOut)

        Dim addr As Long
        addr = startAddrDec + (i - LBound(bytesOut))

        Dim offset As Long
        offset = addr - memStartDec
        If offset < 0 Then Exit Sub

        Dim rowIndex As Long, colIndex As Long
        rowIndex = offset \ BYTES_PER_MEMROW
        colIndex = offset Mod BYTES_PER_MEMROW

        Dim targetRow As Long, targetCol As Long
        targetRow = 1 + rowIndex
        targetCol = 1 + colIndex

        If targetRow > rngMem.Rows.Count Then Exit Sub
        If targetCol > rngMem.Columns.Count Then Exit Sub

        rngMem.Cells(targetRow, targetCol).value = Right$("00" & hex$(CLng(bytesOut(i)) And &HFF&), 2)
        myMem.addr(addr) = CLng(bytesOut(i)) And &HFF&

        baseAddr = memStartDec + (rowIndex * BYTES_PER_MEMROW)
        lastTargetRow = targetRow
    Next i

    If lastTargetRow > 0 Then
        wsCPU.Cells(rngMem.row + (lastTargetRow - 1), gMemTableCol).value = baseAddr
    End If

End Sub

' =============================================================================
' Opcode map from "8080 Op to Hex" table: key = MNEMONIC|OP1SPEC|OP2SPEC
' =============================================================================
Public Function BuildOpcodeMap(ByVal wsOp As Worksheet) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim lastRow As Long
    lastRow = wsOp.Cells(wsOp.Rows.Count, 1).End(xlUp).row

    Dim r As Long, emptyRun As Long
    emptyRun = 0

    ' Columns in your sheet:
    ' A Opcode, B Hex, D OP1, E OP2, F Bytes
    For r = 2 To lastRow
        Dim mnem As String
        mnem = UCase$(Trim$(CStr(wsOp.Cells(r, 1).value)))
        If mnem = "" Then
            emptyRun = emptyRun + 1
            If emptyRun >= 10 Then Exit For
            GoTo ContinueRow
        Else
            emptyRun = 0
        End If

        Dim hexTxt As String: hexTxt = UCase$(Trim$(CStr(wsOp.Cells(r, 2).value)))
        Dim spec1 As String: spec1 = UCase$(Trim$(CStr(wsOp.Cells(r, 4).value)))
        Dim spec2 As String: spec2 = UCase$(Trim$(CStr(wsOp.Cells(r, 5).value)))
        Dim bytes As Long: bytes = CLng(wsOp.Cells(r, 6).value)

        If hexTxt = "" Or bytes < 1 Then GoTo ContinueRow

        Dim opByte As Long
        opByte = usrHexToDec(hexTxt) And &HFF&

        Dim key As String
        key = mnem & "|" & spec1 & "|" & spec2

        If Not dict.Exists(key) Then
            dict.ADD key, Array(opByte, bytes, spec1, spec2)
        End If

ContinueRow:
    Next r

    Set BuildOpcodeMap = dict
End Function

Private Function HlpFindOpcode(ByVal opMap As Object, ByVal mnem As String, ByVal op1 As String, ByVal op2 As String) As Variant
    Dim key As String

    key = mnem & "|" & op1 & "|" & op2
    If opMap.Exists(key) Then HlpFindOpcode = opMap(key): Exit Function

    key = mnem & "|" & op1 & "|"
    If opMap.Exists(key) Then HlpFindOpcode = opMap(key): Exit Function

    Dim p As Variant
    For Each p In Array("BYTE", "ADDRESS", "PORT", "DATA")
        key = mnem & "|" & op1 & "|" & CStr(p)
        If opMap.Exists(key) Then HlpFindOpcode = opMap(key): Exit Function
    Next p

    For Each p In Array("BYTE", "ADDRESS", "PORT", "DATA")
        key = mnem & "|" & CStr(p) & "|" & op2
        If opMap.Exists(key) Then HlpFindOpcode = opMap(key): Exit Function
        key = mnem & "|" & CStr(p) & "|"
        If opMap.Exists(key) Then HlpFindOpcode = opMap(key): Exit Function
        key = mnem & "|" & CStr(p) & "|" & CStr(p)
        If opMap.Exists(key) Then HlpFindOpcode = opMap(key): Exit Function
    Next p

    HlpFindOpcode = Empty
End Function

' =============================================================================
' Encode instruction bytes (1/2/3 bytes), using label/EQU resolution via clsLabels
' =============================================================================
Private Function HlpEncodeInstruction(ByVal lbls As clsLabels, ByVal opByte As Byte, ByVal nBytes As Long, _
                                   ByVal spec1 As String, ByVal spec2 As String, _
                                   ByVal op1 As String, ByVal op2 As String, _
                                   ByVal opname As String, ByVal curRow As Range, ByVal wsCPU As Worksheet) As Variant

    Dim out() As Variant

    Select Case nBytes
        Case 1
            ReDim out(0 To 0)
            out(0) = opByte

        Case 2
            ReDim out(0 To 1)
            out(0) = opByte

            Dim imm8 As Long
            imm8 = ResolveValue8(lbls, HlpPickImmediateToken(spec1, spec2, op1, op2), opname, wsCPU, curRow)
            out(1) = CByte(imm8 And &HFF&)

        Case 3
            ReDim out(0 To 2)
            out(0) = opByte

            Dim imm16 As Long
            imm16 = ResolveValue16(lbls, HlpPickImmediateToken(spec1, spec2, op1, op2))

            ' 8080 immediates are little-endian: low byte then high byte
            out(1) = CByte(imm16 And &HFF&)
            out(2) = CByte((imm16 \ 256) And &HFF&)

        Case Else
            RaiseAsmError wsCPU, "Unsupported instruction length: " & nBytes & " for " & opname, curRow
            HlpEncodeInstruction = Empty
            Exit Function
    End Select

    HlpEncodeInstruction = out
End Function

Private Function HlpPickImmediateToken(ByVal spec1 As String, ByVal spec2 As String, ByVal op1 As String, ByVal op2 As String) As String
    If HlpIsPlaceholder(spec1) Then
        HlpPickImmediateToken = op1
    ElseIf HlpIsPlaceholder(spec2) Then
        HlpPickImmediateToken = op2
    Else
        If Len(op2) > 0 Then HlpPickImmediateToken = op2 Else HlpPickImmediateToken = op1
    End If
End Function

Private Function HlpIsPlaceholder(ByVal s As String) As Boolean
    s = UCase$(Trim$(s))
    HlpIsPlaceholder = (s = "BYTE" Or s = "ADDRESS" Or s = "PORT" Or s = "DATA")
End Function

' =============================================================================
' Operand resolution
' =============================================================================
Private Function ResolveValue8(ByVal lbls As clsLabels, ByVal token As String, _
                               ByVal opname As String, ByVal wsCPU As Worksheet, ByVal curRow As Range) As Long
    token = UCase$(Trim$(token))
    If token = "" Then
        RaiseAsmError wsCPU, opname & " missing 8-bit immediate", curRow
        ResolveValue8 = 0
        Exit Function
    End If

    Dim rec As clsLabelRecord
    Set rec = lbls.GetLabel(token)
    If Not rec Is Nothing Then
        ResolveValue8 = (rec.clsAddress And &HFF&)
        Exit Function
    End If

    If usrValIsHexString(token) Then
        ResolveValue8 = usrHexToDec(token) And &HFF&
    ElseIf IsNumeric(token) Then
        ResolveValue8 = CLng(token) And &HFF&
    Else
        RaiseAsmError wsCPU, opname & " invalid imm8: " & token, curRow
        ResolveValue8 = 0
    End If
End Function

Private Function ResolveValue16(ByVal lbls As clsLabels, ByVal token As String) As Long
    token = UCase$(Trim$(token))
    If token = "" Then ResolveValue16 = 0: Exit Function

    Dim rec As clsLabelRecord
    Set rec = lbls.GetLabel(token)
    If Not rec Is Nothing Then
        ResolveValue16 = (rec.clsAddress And &HFFFF&)
        Exit Function
    End If

    If usrValIsHexString(token) Then
        ResolveValue16 = usrHexToDec(token) And &HFFFF&
    ElseIf IsNumeric(token) Then
        ResolveValue16 = CLng(token) And &HFFFF&
    Else
        ResolveValue16 = 0
    End If
End Function

' =============================================================================
' DB string literal only (for now)
' =============================================================================
Private Function StringToBytesVariant(ByVal s As String) As Variant
    Dim out() As Variant, i As Long
    If Len(s) = 0 Then
        StringToBytesVariant = Empty
        Exit Function
    End If
    ReDim out(0 To Len(s))
    For i = 1 To Len(s)
        out(i - 1) = Asc(Mid$(s, i, 1)) And &HFF&
    Next i
    out(Len(s)) = 0
    StringToBytesVariant = out
End Function

Private Function BytesToHexString(ByVal bytesOut As Variant) As String
    Dim i As Long, s As String
    s = ""
    For i = LBound(bytesOut) To UBound(bytesOut)
        s = s & IIf(Len(s) > 0, " ", "") & Right$("00" & hex$(CLng(bytesOut(i)) And &HFF&), 2)
    Next i
    BytesToHexString = s
End Function

' =============================================================================
' Helpers / safety wrappers
' =============================================================================
Private Function StripComment(ByVal s As String) As String
    Dim p As Long
    p = InStr(1, s, ";", vbBinaryCompare)
    If p > 0 Then StripComment = Trim$(Left$(s, p - 1)) Else StripComment = Trim$(s)
End Function

Private Sub RaiseAsmError(ByVal wsCPU As Worksheet, ByVal msg As String, ByVal row As Range)
    wsCPU.Range("Error").value = 1
    wsCPU.Range("errMessage").value = "ASM ERROR: " & msg & " @ " & row.address(0, 0)
End Sub

Public Sub Compile_SetEnabled(ByVal enabled As Boolean)
    gCompileEnabled = enabled
End Sub

Public Function Compile_IsEnabled() As Boolean
    ' default enabled if never set
    If gCompileEnabled = False And gCompiledForResetStamp = 0 Then
        gCompileEnabled = True
    End If
    Compile_IsEnabled = gCompileEnabled
End Function

Public Sub Compile_Invalidate()
    gCompiledValid = False
    gLastCodeSignature = vbNullString
End Sub
Public Sub CompileIfNeeded(ByVal wsCPU As Worksheet)
    ' If disabled, do nothing
    If Not Compile_IsEnabled() Then Exit Sub

    Dim doReset As Boolean
    doReset = (CLng(wsCPU.Range("Reset").value) = 1)

    ' Only compile on Reset, unless cache invalidated
    If doReset Then
        gCompiledValid = False
    End If

    If gCompiledValid Then Exit Sub

    ' Optional: cheap signature check to avoid compiling if nothing changed even on Reset
    ' (you can comment this out initially)
    'Dim sig As String
    'sig = ComputeCodeSignature(wsCPU)
    'If sig = gLastCodeSignature And doReset = False Then
    '    gCompiledValid = True
    '    Exit Sub
    'End If

    ' Call your actual compiler
    Assemble8080_ToMachine ' or whatever defaults you want

    ' Mark cache valid
    gCompiledValid = True
    gCompiledForResetStamp = Timer
    'gLastCodeSignature = sig
End Sub


