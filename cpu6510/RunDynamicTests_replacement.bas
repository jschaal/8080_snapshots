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
    Set wsEMU = Sheets("CPU")

    memCapacity = usrHexToDec(wsEMU.Range("MemSize").value)
    ofs_label = wsEMU.Range("ofs_label").value
    startCol = ofs_label
    endCol = wsEMU.Range("ofs_op2").value

    ' --- Save and override run-control flags ---
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

    ' CPUTest column is optional - gracefully absent on older sheets
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
        Dim testName  As String: testName  = CStr(t(0))
        Dim testRange As Range:  Set testRange = t(1)
        Dim nameRow   As Long:   nameRow   = CLng(t(2))

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
                totalTests = totalTests + 1
                Application.StatusBar = "Currently Running: " & testName

                If Not testRange Is Nothing Then

                    ' A) Clear emulator area
                    wsEMU.Range("Line0").Offset(0, ofs_label).Resize(memCapacity + 1, endCol).ClearContents
                    wsEMU.Range("Line0").Offset(0, startCol).Resize(memCapacity, 1).ClearContents

                    ' B) Copy program rows
                    wsEMU.Range("Line0").Offset(0, ofs_label).Resize(testRange.Rows.Count, 4).value = _
                        testRange.Columns(3).Resize(, 5).value

                    If TestHasConsoleAssertions(testRange) Then ClearConsole

                    ' C) Set CPU mode for this test FIRST - before compile and execute
                    savedCPUMode = CPUMode()
                    testCPUMode  = ""
                    If cpuCol > 0 Then
                        testCPUMode = UCase$(Trim$(CStr(wsTest.Cells(nameRow, cpuCol).value)))
                    End If
                    ' Blank defaults to 8080 so existing tests without a CPU tag
                    ' always run correctly regardless of the sheet's CPUMode cell
                    If testCPUMode = "" Then testCPUMode = "8080"
                    wsEMU.Range("CPUMode").value = testCPUMode

                    ' D) Compile if requested - uses the now-correct CPUMode
                    compileThis = (val(wsTest.Cells(nameRow, compileCol).value) = 1)
                    If compileThis Then AssembleForCPUMode

                    ' E) Execute
                    wsEMU.Range("Reset").value = 1
                    ConsolePrint ">" & testName & ":", False, False
                    ResetAddressList
                    Application.Calculate
                    SelectEngine

                    ' F) Restore CPUMode to whatever the user had selected
                    wsEMU.Range("CPUMode").value = savedCPUMode

                    ' G) Validate
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
