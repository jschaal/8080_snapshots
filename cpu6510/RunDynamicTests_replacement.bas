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
    ' runCol and compileCol live in the program area header (col M and N),
    ' read via nameRow (the program area row returned by DiscoverAllTests).
    runCol     = wsTest.Range("RunTest").Column      ' col M
    compileCol = wsTest.Range("CompileTest").Column  ' col N
    resultsCol = wsTest.Range("TestRunner").Column   ' col K

    ' cpuCol lives in the control table (col T), read via ctrlRow.
    On Error Resume Next
    cpuCol = wsTest.Range("CPUTest").Column
    If Err.Number <> 0 Then cpuCol = 0
    On Error GoTo 0

    ' --- Control table anchor for CPU column lookup only ---
    ' nameRow from DiscoverAllTests = program area row (col A).
    ' The CPU column is only in the control table (cols P-T).
    ' We match testName against col P to find the control table row.
    Dim ctrlAnchor  As Range:  Set ctrlAnchor = wsTest.Range("TestTable")
    Dim ctrlLastRow As Long
    ctrlLastRow = wsTest.Cells(wsTest.Rows.Count, ctrlAnchor.Column).End(xlUp).Row
    Dim ctrlNames   As Range
    Set ctrlNames = wsTest.Range(ctrlAnchor.Offset(1, 0), _
                                 wsTest.Cells(ctrlLastRow, ctrlAnchor.Column))

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set tests = DiscoverAllTests(wsTest)

    Dim batchPassed As Boolean: batchPassed = True
    Dim passCount   As Long:    passCount   = 0
    Dim totalTests  As Long:    totalTests  = 0

    For Each t In tests
        Dim testName  As String: testName  = CStr(t(0))
        Dim testRange As Range:  Set testRange = t(1)
        Dim nameRow   As Long:   nameRow   = CLng(t(2))  ' program area row

        ' Run and Compile flags come from the program area header row (nameRow)
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

                    ' C) Set CPU mode BEFORE compile and execute.
                    '    CPU column is only in the control table - look it up by name.
                    savedCPUMode = CPUMode()
                    testCPUMode  = "8080"   ' safe default
                    If cpuCol > 0 Then
                        Dim ctrlIdx As Variant
                        ctrlIdx = Application.Match(testName, ctrlNames, 0)
                        If Not IsError(ctrlIdx) Then
                            Dim ctrlRow As Long
                            ctrlRow = ctrlNames.Row + CLng(ctrlIdx) - 1
                            Dim rawCPU As String
                            rawCPU = UCase$(Trim$(CStr(wsTest.Cells(ctrlRow, cpuCol).value)))
                            If rawCPU <> "" Then testCPUMode = rawCPU
                        End If
                    End If
                    wsEMU.Range("CPUMode").value = testCPUMode

                    ' D) Compile if requested (uses now-correct CPUMode)
                    compileThis = (val(wsTest.Cells(nameRow, compileCol).value) = 1)
                    If compileThis Then AssembleForCPUMode

                    ' E) Execute
                    wsEMU.Range("Reset").value = 1
                    ConsolePrint ">" & testName & ":", False, False
                    ResetAddressList
                    Application.Calculate
                    SelectEngine

                    ' F) Restore CPUMode to what user had selected
                    wsEMU.Range("CPUMode").value = savedCPUMode

                    ' G) Validate and record result
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
    If overrideStepMode   Then Range("Step")       = 1
    If overrideReset      Then Range("Reset")      = 0
    If overrideTrace      Then Range("Trace")      = 1
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
