Option Explicit
'================================================================================
' TestRunner  —  changes to support per-test CPU selection
'
' WHAT THIS ADDS:
'   A "CPU" column in the control table (column T, after Compile).
'   Each test row can specify "8080", "Z80", or "6510".
'   Blank means "use whatever CPUMode is currently set" (backwards-compatible).
'   The runner saves and restores CPUMode just like it already does for
'   Step, Reset, Trace, and SkipBreaks.
'
' HOW TO APPLY:
'   1. On the "Unit Tests" sheet, add a header cell "CPU" in T5
'      (one column to the right of "Compile" in S5).
'   2. Name that cell "CPUTest" using the Name Box.
'   3. For each test row in the control table, put the CPU model in column T:
'        8080  (or leave blank to inherit current setting)
'        Z80
'        6510
'   4. In the TestRunner module, make the FOUR edits described below.
'
' EDITS TO RunDynamicTests — make these four changes:
'================================================================================

' -----------------------------------------------------------------------
' EDIT 1: Add cpuCol declaration alongside the existing runCol / compileCol
' Find this block:
'
'     Dim resultsCol As Long, runCol As Long, compileCol As Long
'
' Replace with:
'
'     Dim resultsCol As Long, runCol As Long, compileCol As Long, cpuCol As Long
' -----------------------------------------------------------------------

' -----------------------------------------------------------------------
' EDIT 2: Add overrideCPUMode variables alongside the existing override vars.
' Find this block:
'
'     Dim overrideStepMode As Boolean
'     Dim overrideReset As Boolean
'     Dim overrideTrace As Boolean
'     Dim overrideSkipBreaks As Boolean
'
' Replace with:
'
'     Dim overrideStepMode As Boolean
'     Dim overrideReset As Boolean
'     Dim overrideTrace As Boolean
'     Dim overrideSkipBreaks As Boolean
'     Dim savedCPUMode As String
'     Dim testCPUMode As String
' -----------------------------------------------------------------------

' -----------------------------------------------------------------------
' EDIT 3: Read cpuCol after the existing runCol / compileCol lines.
' Find this block:
'
'     runCol     = wsTest.Range("RunTest").Column
'     compileCol = wsTest.Range("CompileTest").Column
'     resultsCol = wsTest.Range("TestRunner").Column
'
' Replace with:
'
'     runCol     = wsTest.Range("RunTest").Column
'     compileCol = wsTest.Range("CompileTest").Column
'     resultsCol = wsTest.Range("TestRunner").Column
'     On Error Resume Next
'     cpuCol = wsTest.Range("CPUTest").Column
'     If Err.Number <> 0 Then cpuCol = 0
'     On Error GoTo 0
' -----------------------------------------------------------------------

' -----------------------------------------------------------------------
' EDIT 4: Inside the For Each t loop, just before "If runThis Then",
' add the CPU mode override/restore around SelectEngine.
'
' Find this block (the execute section, inside "If Not testRange Is Nothing"):
'
'                     ' C) Execute
'                     wsEMU.Range("Reset").value = 1
'                     ConsolePrint ">" & testName & ":", False, False
'                     ResetAddressList
'                     Application.Calculate
'                     SelectEngine
'
' Replace with:
'
'                     ' C) Execute
'                     wsEMU.Range("Reset").value = 1
'                     ConsolePrint ">" & testName & ":", False, False
'                     ResetAddressList
'                     Application.Calculate
'                     ' --- Per-test CPU override ---
'                     savedCPUMode = CPUMode()
'                     testCPUMode = ""
'                     If cpuCol > 0 Then
'                         testCPUMode = UCase$(Trim$(CStr(wsTest.Cells(nameRow, cpuCol).value)))
'                     End If
'                     If testCPUMode <> "" Then
'                         wsEMU.Range("CPUMode").value = testCPUMode
'                     End If
'                     SelectEngine
'                     If testCPUMode <> "" Then
'                         wsEMU.Range("CPUMode").value = savedCPUMode
'                     End If
' -----------------------------------------------------------------------

'================================================================================
' FULL UPDATED RunDynamicTests for reference
' (easier to read as a complete sub than as patch instructions)
' You can replace your existing RunDynamicTests entirely with this.
'================================================================================

Public Sub RunDynamicTests()
    Dim wsTest As Worksheet, wsEMU As Worksheet
    Dim memCapacity As Long
    Dim startCol As Integer, ofs_label As Integer
    Dim endCol As Integer
    Dim tests As Collection
    Dim t As Variant
    Dim resultsCol As Long, runCol As Long, compileCol As Long, cpuCol As Long
    Dim overrideStepMode As Boolean
    Dim overrideReset As Boolean
    Dim overrideTrace As Boolean
    Dim overrideSkipBreaks As Boolean
    Dim savedCPUMode As String
    Dim testCPUMode As String
    Dim compileThis As Boolean

    Set wsTest = Sheets("Unit Tests")
    Set wsEMU  = Sheets("CPU")

    memCapacity = usrHexToDec(wsEMU.Range("MemSize").value)
    ofs_label   = wsEMU.Range("ofs_label").value
    startCol    = ofs_label
    endCol      = wsEMU.Range("ofs_op2").value

    ' --- Save and override run-control flags (unchanged from original) ---
    overrideStepMode = Range("Step") = 1
    If overrideStepMode Then Range("Step") = 0

    overrideReset = Range("Reset") = 0
    If overrideReset Then Range("Reset") = 1

    overrideTrace = Range("Trace") = 1
    If overrideTrace Then Range("Trace") = 0

    overrideSkipBreaks = Range("SkipBreaks") = 0
    If overrideSkipBreaks Then Range("SkipBreaks") = 1

    ' --- Column lookups ---
    runCol     = wsTest.Range("RunTest").Column
    compileCol = wsTest.Range("CompileTest").Column
    resultsCol = wsTest.Range("TestRunner").Column

    ' CPUTest column is optional — gracefully absent on older sheets
    On Error Resume Next
    cpuCol = wsTest.Range("CPUTest").Column
    If Err.Number <> 0 Then cpuCol = 0
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set tests = DiscoverAllTests(wsTest)

    Dim batchPassed As Boolean: batchPassed = True
    Dim passCount   As Long:    passCount   = 0
    Dim totalTests  As Long:    totalTests  = 0

    For Each t In tests
        Dim testName  As String:   testName  = CStr(t(0))
        Dim testRange As Range:    Set testRange = t(1)
        Dim nameRow   As Long:     nameRow   = CLng(t(2))

        Dim runThis As Boolean
        runThis = (val(wsTest.Cells(nameRow, runCol).value) = 1)

        If runThis Then
            If testName = "Unit_VBA" Then
                Dim passed As Boolean
                VBARunAllTests
                passed = VBAPassed
                totalTests = totalTests + VBATestCount
                passCount  = passCount + VBATestCount - VBAFailureCount
                If passed Then
                    wsTest.Cells(nameRow, resultsCol).value          = "PASS"
                    wsTest.Cells(nameRow, resultsCol).Interior.Color = vbGreen
                    Call SetRunFlagInControlTable(testName, 0)
                Else
                    wsTest.Cells(nameRow, resultsCol).value          = "FAIL"
                    wsTest.Cells(nameRow, resultsCol).Interior.Color = vbRed
                    batchPassed = False
                End If
            Else
                compileThis = (val(wsTest.Cells(nameRow, compileCol).value) = 1)
                If compileThis Then Assemble8080_ToMachine

                totalTests = totalTests + 1
                Application.StatusBar = "Currently Running: " & testName

                If Not testRange Is Nothing Then
                    ' A) Clear emulator area
                    wsEMU.Range("Line0").offset(0, ofs_label).Resize(memCapacity + 1, endCol).ClearContents
                    wsEMU.Range("Line0").offset(0, startCol).Resize(memCapacity, 1).ClearContents

                    ' B) Copy program rows
                    wsEMU.Range("Line0").offset(0, ofs_label).Resize(testRange.Rows.Count, 4).value = _
                        testRange.Columns(3).Resize(, 5).value

                    If TestHasConsoleAssertions(testRange) Then ClearConsole

                    ' C) Execute — with per-test CPU override
                    wsEMU.Range("Reset").value = 1
                    ConsolePrint ">" & testName & ":", False, False
                    ResetAddressList
                    Application.Calculate

                    savedCPUMode = CPUMode()
                    testCPUMode  = ""
                    If cpuCol > 0 Then
                        testCPUMode = UCase$(Trim$(CStr(wsTest.Cells(nameRow, cpuCol).value)))
                    End If
                    If testCPUMode <> "" Then wsEMU.Range("CPUMode").value = testCPUMode
                    SelectEngine
                    If testCPUMode <> "" Then wsEMU.Range("CPUMode").value = savedCPUMode

                    ' D) Validate
                    If TestValidateMultipleCriteria(testRange) Then
                        wsTest.Cells(nameRow, resultsCol).value          = "PASS"
                        wsTest.Cells(nameRow, resultsCol).Interior.Color = vbGreen
                        passCount = passCount + 1
                        Call SetRunFlagInControlTable(testName, 0)
                    Else
                        wsTest.Cells(nameRow, resultsCol).value          = "FAIL"
                        wsTest.Cells(nameRow, resultsCol).Interior.Color = vbRed
                        batchPassed = False
                    End If

                    ConsolePrint wsTest.Cells(nameRow, resultsCol).value, True
                Else
                    wsTest.Cells(nameRow, resultsCol).value          = "SKIPPED"
                    wsTest.Cells(nameRow, resultsCol).Interior.Color = RGB(200, 200, 200)
                End If
            End If
        Else
            wsTest.Cells(nameRow, resultsCol).value          = "SKIPPED"
            wsTest.Cells(nameRow, resultsCol).Interior.Color = RGB(200, 200, 200)
        End If
    Next t

    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.ScreenUpdating = True
    Application.StatusBar = False
    ConsolePrint "", True
    Range("consoleText") = usrForm.tbConsole.value

    ' --- Restore run-control flags ---
    If overrideStepMode  Then Range("Step")       = 1
    If overrideReset     Then Range("Reset")      = 0
    If overrideTrace     Then Range("Trace")      = 1
    If overrideSkipBreaks Then Range("SkipBreaks") = 0

    Dim finalMsg As String
    If batchPassed Then
        finalMsg = "OVERALL RESULT: PASS" & vbCrLf & _
                   "Passed " & passCount & " of " & totalTests & " tests."
        MsgBox finalMsg, vbInformation, "Test Runner Success"
    Else
        finalMsg = "OVERALL RESULT: FAIL" & vbCrLf & _
                   "Only " & passCount & " of " & totalTests & " tests passed."
        MsgBox finalMsg, vbCritical, "Test Runner Failure"
    End If
End Sub
