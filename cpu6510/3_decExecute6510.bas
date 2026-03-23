Option Explicit
'================================================================================
' FILE 3 OF 4  —  decExecute6510  (complete new module)
'
' HOW TO APPLY:
'   In the VBA editor: Insert > Module, rename it "decExecute6510",
'   then paste the entire contents of this file into it.
'
' DEPENDENCIES (must exist before this compiles):
'   modGlobals  — gCacheValid, gCacheLine0Dec, gCacheCountRows, gCacheMemStart,
'                 gCacheMemEnd, gCacheOfsOpcode, gCacheOfsOp1, gCacheOfsOp2,
'                 gCacheOfsRowStat, gCacheOfsLabel, gArrOpcode, gArrRowStat,
'                 gArrOp1, gArrOp2, gArrLabel, gBreak, gCurrentIter,
'                 InvalidateExecCache, g6510, ResetDecCPU6510, mConsoleOpen,
'                 gMemory, ResetAddressList, usrHexToDec, usrValIsHexString,
'                 ERR_EXEC_END, ERR_BAD_OPCODE, ERR_MAX_ITERS, ERR_STOP
'   cls6510CPU  — the 6510 CPU class
'   Trace module — TraceEnabled, TraceLive, TraceClear, TraceFlush,
'                  TraceAppend6510 (paste that sub into the Trace module;
'                  template at the bottom of this file)
'================================================================================

'------------------------------------------------------------------------------
' Execute6510
' Main 6510 execution loop.  Identical structure to Execute8080.
' Uses the shared cache from modGlobals directly — no module-private cache.
'------------------------------------------------------------------------------
Public Sub Execute6510()

    Dim wsCPU As Worksheet
    Set wsCPU = ThisWorkbook.Worksheets("CPU")

    '--- Named-range reads ---
    Dim stepMode   As Boolean: stepMode   = (CLng(wsCPU.Range("Step").value) = 1)
    Dim doReset    As Boolean: doReset    = (CLng(wsCPU.Range("Reset").value) = 1)
    Dim skipBreaks As Boolean: skipBreaks = (CLng(wsCPU.Range("SkipBreaks").value) = 1)
    Dim maxIters   As Long:    maxIters   = CLng(wsCPU.Range("Max_Iters").value)

    Dim ofs_opcode  As Long: ofs_opcode  = CLng(wsCPU.Range("ofs_opcode").value)
    Dim ofs_op1     As Long: ofs_op1     = CLng(wsCPU.Range("ofs_op1").value)
    Dim ofs_op2     As Long: ofs_op2     = CLng(wsCPU.Range("ofs_op2").value)
    Dim ofs_rowstat As Long: ofs_rowstat = CLng(wsCPU.Range("ofs_rowstat").value)
    Dim ofs_label   As Long: ofs_label   = CLng(wsCPU.Range("ofs_label").value)

    Dim memStart As Long: memStart = usrHexToDec(CStr(wsCPU.Range("MemStart").value))
    Dim memEnd   As Long: memEnd   = memStart + usrHexToDec(CStr(wsCPU.Range("MemSize").value))

    '--- Run state ---
    Dim myCpu            As cls6510CPU
    Dim errcode          As Long:    errcode          = 0
    Dim pc_dec           As Long
    Dim headless         As Boolean
    Dim prevCalc         As XlCalculation
    Dim stepTaken        As Boolean: stepTaken        = False
    Dim executedThisIter As Boolean
    Dim endedByEnd       As Boolean: endedByEnd       = False

    wsCPU.Range("Stop").value       = 0
    wsCPU.Range("Error").value      = 0
    wsCPU.Range("errMessage").value = ""

    '--- ORG on first line (same as Execute8080) ---
    If UCase$(CStr(wsCPU.Range("Line0").offset(0, ofs_opcode).value)) = "ORG" And doReset Then
        Dim orgVal As String: orgVal = CStr(wsCPU.Range("Line0").offset(0, ofs_op1).value)
        If usrValIsHexString(orgVal) Then
            wsCPU.Range("MemStart").value = orgVal
            Application.Calculate
            memStart = usrHexToDec(orgVal)
            memEnd   = memStart + usrHexToDec(CStr(wsCPU.Range("MemSize").value))
        End If
    End If

    '--- Reset vs continue ---
    Set myCpu = g6510

    If doReset Then
        If TraceEnabled() Then TraceClear
        InvalidateExecCache                  ' shared cache in modGlobals
        ResetDecCPU6510
        Set myCpu = g6510
        ResetAddressList
        gBreak       = False                 ' shared in modGlobals
        gCurrentIter = 0                     ' shared in modGlobals
        mConsoleOpen = False

        pc_dec = CLng(wsCPU.Range("Line0_dec").value)
        myCpu.SetPC pc_dec
        myCpu.SetReg "SP", &HFF              ' 6510 SP = $FF at reset

        wsCPU.Range("StackStart").value = "01FF"
        wsCPU.Range("StackDetails").ClearContents

        If stepMode Then
            RefreshFlags6510 myCpu
            RefreshRegisters6510 myCpu
            myCpu.RefreshStack True
            Application.Calculate
        End If
    Else
        pc_dec = CLng(myCpu.Reg("PC"))
    End If

    '--- Headless toggle ---
    headless = Not stepMode
    If headless Then
        Application.ScreenUpdating = False
        Application.EnableEvents   = False
        prevCalc = Application.Calculation
        Application.Calculation = xlCalculationManual
    End If

    '--- Build / refresh shared cache ---
    Dim base      As Range: Set base  = wsCPU.Range("Line0_dec")
    Dim line0_dec As Long:  line0_dec = CLng(base.value)
    Dim countRows As Long:  countRows = (memEnd - memStart) + 1
    If countRows < 1 Then countRows = 1

    If (Not gCacheValid)                      Or _
       (gCacheCountRows  <> countRows)        Or _
       (gCacheLine0Dec   <> line0_dec)        Or _
       (gCacheMemStart   <> memStart)         Or _
       (gCacheMemEnd     <> memEnd)           Or _
       (gCacheOfsOpcode  <> ofs_opcode)       Or _
       (gCacheOfsOp1     <> ofs_op1)          Or _
       (gCacheOfsOp2     <> ofs_op2)          Or _
       (gCacheOfsRowStat <> ofs_rowstat)      Or _
       (gCacheOfsLabel   <> ofs_label) Then

        gCacheLine0Dec   = line0_dec
        gCacheCountRows  = countRows
        gCacheMemStart   = memStart
        gCacheMemEnd     = memEnd
        gCacheOfsOpcode  = ofs_opcode
        gCacheOfsOp1     = ofs_op1
        gCacheOfsOp2     = ofs_op2
        gCacheOfsRowStat = ofs_rowstat
        gCacheOfsLabel   = ofs_label

        gArrOpcode  = base.offset(0, ofs_opcode).Resize(countRows, 1).value
        gArrRowStat = base.offset(0, ofs_rowstat).Resize(countRows, 1).value
        gArrOp1     = base.offset(0, ofs_op1).Resize(countRows, 1).value
        gArrOp2     = base.offset(0, ofs_op2).Resize(countRows, 1).value
        gArrLabel   = base.offset(0, ofs_label).Resize(countRows, 1).value

        Dim idxN As Long
        For idxN = 1 To countRows
            If Not IsEmpty(gArrOpcode(idxN, 1))  Then gArrOpcode(idxN, 1)  = Trim$(UCase$(CStr(gArrOpcode(idxN, 1))))
            If Not IsEmpty(gArrRowStat(idxN, 1)) Then gArrRowStat(idxN, 1) = Trim$(UCase$(CStr(gArrRowStat(idxN, 1))))
            If Not IsEmpty(gArrOp1(idxN, 1))     Then gArrOp1(idxN, 1)     = Trim$(CStr(gArrOp1(idxN, 1)))
            If Not IsEmpty(gArrOp2(idxN, 1))     Then gArrOp2(idxN, 1)     = Trim$(CStr(gArrOp2(idxN, 1)))
            If Not IsEmpty(gArrLabel(idxN, 1))   Then gArrLabel(idxN, 1)   = Trim$(UCase$(CStr(gArrLabel(idxN, 1))))
        Next idxN

        gCacheValid = True
    End If

    '--- Copy to locals for the hot loop ---
    Dim arrOpcode  As Variant: arrOpcode  = gArrOpcode
    Dim arrRowStat As Variant: arrRowStat = gArrRowStat
    Dim arrOp1     As Variant: arrOp1     = gArrOp1
    Dim arrOp2     As Variant: arrOp2     = gArrOp2
    Dim arrLabel   As Variant: arrLabel   = gArrLabel
    Dim ubRows     As Long:    ubRows     = UBound(arrOpcode, 1)

    '--- Establish starting row ---
    Dim rowIdx As Long: rowIdx = myCpu.GetCurrentIdx
    If rowIdx < 0 Or rowIdx >= ubRows Then GoTo ExecFinalize6510

    Dim opcode As String: opcode = CStr(arrOpcode(rowIdx + 1, 1))

    '==========================================================================
    ' MAIN LOOP
    '==========================================================================
    Do While errcode = 0 And pc_dec <= memEnd _
         And (Not stepMode Or (stepMode And Not stepTaken))

        gCurrentIter = gCurrentIter + 1

        Dim op1     As String: op1     = CStr(arrOp1(rowIdx + 1, 1))
        Dim op2     As String: op2     = CStr(arrOp2(rowIdx + 1, 1))
        Dim label   As String: label   = CStr(arrLabel(rowIdx + 1, 1))
        Dim rowStat As String: rowStat = CStr(arrRowStat(rowIdx + 1, 1))

        op1 = Trim$(UCase$(op1))
        op2 = Trim$(UCase$(op2))

        '--- Break handling (identical to Execute8080) ---
        If rowStat = "B" And Not stepMode And Not skipBreaks Then
            If gBreak Then
                gBreak = False
            Else
                gBreak = True
                Exit Do
            End If
        End If

        executedThisIter = False

        If rowStat <> "C" And opcode <> "" Then

            '--- Trace: pre-execution memory snapshot ---
            Dim memAddrHex  As String: memAddrHex  = ""
            Dim memBefore   As String: memBefore   = ""
            Dim memAfter    As String: memAfter    = ""
            Dim memNote     As String: memNote     = ""
            Dim memAddrDec  As Long:   memAddrDec  = 0
            Dim memTrackLen As Long:   memTrackLen = 0

            If TraceEnabled() Then
                Dim preSP As Long: preSP = myCpu.Reg("SP") And 255
                Select Case opcode
                    Case "STA", "STX", "STY"
                        memNote = opcode & " -> " & op1
                    Case "PHA", "PHP"
                        memAddrDec  = &H100& + preSP
                        memTrackLen = 1
                        memNote = opcode & " push $01" & Right$("00" & Hex$(preSP), 2)
                    Case "JSR"
                        memAddrDec  = &H100& + ((preSP - 1) And 255)
                        memTrackLen = 2
                        memNote = "JSR push ret addr"
                End Select
                If memTrackLen = 1 Then
                    memBefore  = Right$("00" & Hex$(CLng(gMemory.addr(memAddrDec)) And 255), 2)
                    memAddrHex = Right$("0000" & Hex$(memAddrDec), 4)
                ElseIf memTrackLen = 2 Then
                    Dim b0 As Long: b0 = CLng(gMemory.addr(memAddrDec)) And 255
                    Dim b1 As Long: b1 = CLng(gMemory.addr((memAddrDec + 1) And &HFFFF&)) And 255
                    memBefore  = Right$("00" & Hex$(b0), 2) & " " & Right$("00" & Hex$(b1), 2)
                    memAddrHex = Right$("0000" & Hex$(memAddrDec), 4)
                End If
            End If

            '--- Execute ---
            errcode = myCpu.RunOpcode(opcode, op1, op2, label)
            executedThisIter = True

            '--- Trace: post-execution snapshot ---
            If TraceEnabled() Then
                If memTrackLen >= 1 Then
                    Dim a0 As Long: a0 = CLng(gMemory.addr(memAddrDec)) And 255
                    memAfter = Right$("00" & Hex$(a0), 2)
                    If memTrackLen = 2 Then
                        Dim a1 As Long: a1 = CLng(gMemory.addr((memAddrDec + 1) And &HFFFF&)) And 255
                        memAfter = memAfter & " " & Right$("00" & Hex$(a1), 2)
                    End If
                    If memBefore = memAfter Then
                        memAddrHex = "": memBefore = "": memAfter = "": memNote = ""
                    End If
                End If
                TraceAppend6510 myCpu, pc_dec, opcode, op1, op2, errcode, _
                                memAddrHex, memBefore, memAfter, memNote
                If stepMode And TraceLive() Then TraceFlush
            End If

            If errcode = ERR_EXEC_END Then endedByEnd = True: Exit Do

            If errcode = ERR_BAD_OPCODE Then
                wsCPU.Range("Error").value = errcode
                wsCPU.Range("errMessage").value = _
                    "6510 Error: Unknown opcode [" & opcode & "] at $" & _
                    Right$("0000" & Hex$(pc_dec), 4)
                Exit Do
            End If

            If gCurrentIter >= maxIters Then
                errcode = ERR_MAX_ITERS
                wsCPU.Range("Error").value = errcode
                wsCPU.Range("errMessage").value = _
                    "6510 Error: Max iterations (" & maxIters & ") at $" & _
                    Right$("0000" & Hex$(pc_dec), 4)
                Exit Do
            End If
        End If

        '--- Advance PC ---
        If errcode = 0 Then pc_dec = myCpu.IncPC()

        rowIdx = myCpu.GetCurrentIdx
        If rowIdx < 0 Or rowIdx >= ubRows Then Exit Do

        opcode = CStr(arrOpcode(rowIdx + 1, 1))

        If stepMode And executedThisIter Then stepTaken = True

        '--- Step-mode UI paint ---
        If stepMode And errcode = 0 Then
            Do While opcode = "" And pc_dec <= memEnd
                pc_dec = myCpu.IncPC()
                rowIdx = pc_dec - line0_dec
                If rowIdx < 0 Or rowIdx >= ubRows Then Exit Do
                opcode = CStr(arrOpcode(rowIdx + 1, 1))
            Loop
            If myCpu.RegistersDirty Then RefreshRegisters6510 myCpu: myCpu.RegistersDirty = False
            If myCpu.FlagsDirty     Then RefreshFlags6510 myCpu:     myCpu.FlagsDirty     = False
            If myCpu.StackDirty     Then myCpu.RefreshStack True:    myCpu.StackDirty     = False
        End If

        If wsCPU.Range("Stop").value = 1 Then
            errcode = ERR_STOP
            wsCPU.Range("Error").value = errcode
            wsCPU.Range("errMessage").value = _
                "6510 Stopped at $" & Right$("0000" & Hex$(pc_dec), 4)
            Exit Do
        End If
    Loop

    '==========================================================================
ExecFinalize6510:
    If headless Or stepMode Or gBreak Then
        Application.Calculation    = prevCalc
        Application.EnableEvents   = True
        Application.ScreenUpdating = True
        If Not stepMode Then
            RefreshRegisters6510 myCpu
            RefreshFlags6510 myCpu
            myCpu.RefreshStack True
        Else
            wsCPU.Range("PC").value = Right$("0000" & Hex$(myCpu.Reg("PC")), 4)
        End If
        Application.StatusBar = False
        wsCPU.Range("consoleText") = usrForm.tbConsole.value
    End If

    If TraceEnabled() Then TraceFlush

    If gBreak Then
        wsCPU.Range("errMessage") = "Break at $" & Right$("0000" & Hex$(pc_dec), 4)
        wsCPU.Range("Reset").value = 0
    ElseIf endedByEnd Then
        wsCPU.Range("errMessage") = "Execution Complete"
    End If

End Sub

'------------------------------------------------------------------------------
' RefreshRegisters6510
' Named ranges required on CPU sheet: R_A, R_X, R_Y, PC, SP
'------------------------------------------------------------------------------
Public Sub RefreshRegisters6510(ByRef myCpu As cls6510CPU)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("CPU")
    ws.Range("R_A").value = Right$("00" & Hex$(myCpu.Reg("A") And 255), 2)
    ws.Range("R_X").value = Right$("00" & Hex$(myCpu.Reg("X") And 255), 2)
    ws.Range("R_Y").value = Right$("00" & Hex$(myCpu.Reg("Y") And 255), 2)
    ws.Range("PC").value  = Right$("0000" & Hex$(myCpu.Reg("PC") And &HFFFF&), 4)
    ws.Range("SP").value  = Right$("00" & Hex$(myCpu.Reg("SP") And 255), 2)
    On Error GoTo 0
    myCpu.RegistersDirty = False
End Sub

'------------------------------------------------------------------------------
' RefreshFlags6510
' Named ranges required on CPU sheet: N, V, B, D, I, Z, C, SR
'------------------------------------------------------------------------------
Public Sub RefreshFlags6510(ByRef myCpu As cls6510CPU)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("CPU")
    ws.Range("N").value  = myCpu.Flag("N")
    ws.Range("V").value  = myCpu.Flag("V")
    ws.Range("B").value  = myCpu.Flag("B")
    ws.Range("D").value  = myCpu.Flag("D")
    ws.Range("I").value  = myCpu.Flag("I")
    ws.Range("Z").value  = myCpu.Flag("Z")
    ws.Range("C").value  = myCpu.Flag("C")
    ws.Range("SR").value = Right$("00" & Hex$(myCpu.GetSR()), 2)
    On Error GoTo 0
    myCpu.FlagsDirty = False
End Sub

'================================================================================
' PASTE THE FOLLOWING SUB INTO THE Trace MODULE
' (alongside the existing TraceAppend sub — it needs access to that module's
'  Private gTraceArr, gTraceCount, TRACE_MAX_ROWS, TraceEnsureHeaders,
'  and TraceEnsureCapacity, which is why it must live there not here)
'
' Register columns:  A->col6, X->col7, Y->col8, cols 9-12 blank, SP->col13
' Flag columns:      N->col14(S), Z->col15, V->col16(P), C->col17(CY), I->col18(AC)
'================================================================================
'
'Public Sub TraceAppend6510(ByRef cpu As cls6510CPU, _
'                           ByVal pc_dec   As Long, _
'                           ByVal opcode   As String, _
'                           ByVal op1      As String, _
'                           ByVal op2      As String, _
'                           ByVal errcode  As Long, _
'                           Optional ByVal memAddr   As String = "", _
'                           Optional ByVal memBefore As String = "", _
'                           Optional ByVal memAfter  As String = "", _
'                           Optional ByVal memNote   As String = "")
'    If Not TraceEnabled() Then Exit Sub
'    If gTraceCount >= TRACE_MAX_ROWS Then Exit Sub
'    TraceEnsureHeaders
'    gTraceCount = gTraceCount + 1
'    TraceEnsureCapacity gTraceCount
'    Dim r As Long: r = gTraceCount
'    gTraceArr(r,  1) = r
'    gTraceArr(r,  2) = Right$("0000" & Hex$(pc_dec And &HFFFF&), 4)
'    gTraceArr(r,  3) = opcode
'    gTraceArr(r,  4) = op1
'    gTraceArr(r,  5) = op2
'    gTraceArr(r,  6) = Right$("00" & Hex$(cpu.Reg("A")  And 255), 2)
'    gTraceArr(r,  7) = Right$("00" & Hex$(cpu.Reg("X")  And 255), 2)
'    gTraceArr(r,  8) = Right$("00" & Hex$(cpu.Reg("Y")  And 255), 2)
'    gTraceArr(r,  9) = ""   ' no D
'    gTraceArr(r, 10) = ""   ' no E
'    gTraceArr(r, 11) = ""   ' no H
'    gTraceArr(r, 12) = ""   ' no L
'    gTraceArr(r, 13) = Right$("00" & Hex$(cpu.Reg("SP") And 255), 2)
'    gTraceArr(r, 14) = cpu.Flag("N")   ' S  col = Negative
'    gTraceArr(r, 15) = cpu.Flag("Z")   ' Z  col = Zero
'    gTraceArr(r, 16) = cpu.Flag("V")   ' P  col = oVerflow
'    gTraceArr(r, 17) = cpu.Flag("C")   ' CY col = Carry
'    gTraceArr(r, 18) = cpu.Flag("I")   ' AC col = Interrupt (repurposed)
'    gTraceArr(r, 19) = errcode
'    gTraceArr(r, 20) = memAddr
'    gTraceArr(r, 21) = memBefore
'    gTraceArr(r, 22) = memAfter
'    gTraceArr(r, 23) = memNote
'End Sub
