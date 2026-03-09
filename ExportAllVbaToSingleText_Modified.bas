Public Sub ExportAllVbaToSingleText()
    '==============================================================================
    ' Modified ExportAllVbaToSingleText
    ' Enhancements:
    '   - Logs exports to "Export Log" worksheet
    '   - Automatically increments VersionNumber by 0.01
    '   - All files exported in this run share the same version number
    '   - Adds version number to exported filenames
    '==============================================================================
    
    Dim vbProj As Object ' late bound VBIDE.VBProject
    Dim vbComp As Object
    Dim codeMod As Object
    Dim f, g As Integer
    Dim outputDir As Variant
    Dim baseName As String
    Dim timeStamp As String
    Dim finalFileName As String
    Dim outputPath As String
    Dim outputPathStndard As String
    Dim fullPathXlsm As String
    Dim lastLine As Long
    Dim fso As Object
    
    ' Version tracking variables
    Dim currentVersion As Double
    Dim newVersion As Double
    Dim versionStr As String
    Dim exportLogWs As Worksheet
    Dim nextLogRow As Long
    Dim lastLogRow As Long
    
    ' Initialization
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set exportLogWs = ThisWorkbook.Sheets("Export Log")
    
    ' Find the last row with data in the Export Log (starting from row 3)
    lastLogRow = 2
    Do While exportLogWs.Cells(lastLogRow + 1, 1).value <> Empty
        lastLogRow = lastLogRow + 1
    Loop
    
    ' Get the version from the last log entry, or use 1.00 if table is empty
    If lastLogRow > 2 Then
        currentVersion = CDbl(exportLogWs.Cells(lastLogRow, 1).value)
    Else
        currentVersion = 1.0
    End If
    
    ' Increment version for this run
    newVersion = currentVersion + 0.01
    newVersion = Application.Round(newVersion, 2) ' Ensure proper rounding
    versionStr = Format(newVersion, "0.00")
    
    ' 1) Timestamp
    timeStamp = Format(Now, "yyyy-mm-dd_HHmm")
    
    ' 2) Workbook name without extension
    baseName = ThisWorkbook.name
    If InStrRev(baseName, ".") > 0 Then
        baseName = Left(baseName, InStrRev(baseName, ".") - 1)
    End If
    
    ' Declare variables for export loop
    Dim base As Range, saveRow As Range
    Set base = Range("Exports")
    
    ' 3) Export to all selected output directories
    For Each saveRow In base.Rows
        If saveRow.Cells(1, 1) = 1 Then
            outputDir = saveRow.Cells(1, 2)
            
            ' Ensure trailing backslash
            If Right(outputDir, 1) <> "\" Then outputDir = outputDir & "\"
            
            ' Specific file paths - WITH VERSION NUMBER
            finalFileName = baseName & "_" & timeStamp & "_v" & versionStr
            outputPath = outputDir & finalFileName & ".txt"
            fullPathXlsm = outputDir & finalFileName & ".xlsm"
            outputPathStndard = outputDir & baseName & ".txt"
            
            ' Ensure folder exists
            On Error Resume Next
            If Not fso.FolderExists(outputDir) Then fso.CreateFolder outputDir
            On Error GoTo 0
            
            ' Open the text files
            f = FreeFile
            g = FreeFile
            Open outputPath For Output As #f
            
            g = FreeFile
            Open outputPathStndard For Output As #g
            
            Set vbProj = ThisWorkbook.VBProject
            Print #f, "=== VBA EXPORT: " & ThisWorkbook.name & " ==="
            Print #f, "Timestamp: " & Now
            Print #f, "Version: " & versionStr
            Print #f, String(80, "=")
            
            Print #g, "=== VBA EXPORT: " & ThisWorkbook.name & " ==="
            Print #g, "Timestamp: " & Now
            Print #g, "Version: " & versionStr
            Print #g, String(80, "=")
    
            For Each vbComp In vbProj.VBComponents
                Set codeMod = vbComp.CodeModule
                lastLine = codeMod.CountOfLines
    
                Print #f, vbCrLf & "''' Component: " & vbComp.name
                Print #g, vbCrLf & "''' Component: " & vbComp.name
                If lastLine > 0 Then
                    Print #f, codeMod.lines(1, lastLine)
                    Print #g, codeMod.lines(1, lastLine)
                Else
                    Print #f, "''' [No code]"
                    Print #g, "''' [No code]"
                End If
            Next vbComp
    
            Close #f
            Close #g
    
            ' Binary copy of workbook (with version in filename)
            ThisWorkbook.SaveCopyAs fullPathXlsm
            
            ' Find next empty row in Export Log and write one entry per file
            nextLogRow = 3
            Do While exportLogWs.Cells(nextLogRow, 1).value <> Empty Or _
                     exportLogWs.Cells(nextLogRow, 2).value <> Empty Or _
                     exportLogWs.Cells(nextLogRow, 3).value <> Empty
                nextLogRow = nextLogRow + 1
            Loop
            
            ' Write entry for .txt file (full path)
            With exportLogWs
                .Cells(nextLogRow, 1).value = newVersion
                .Cells(nextLogRow, 2).value = Now
                .Cells(nextLogRow, 3).value = outputPath
            End With
            
            ' Write entry for .xlsm file (full path, next row)
            With exportLogWs
                .Cells(nextLogRow + 1, 1).value = newVersion
                .Cells(nextLogRow + 1, 2).value = Now
                .Cells(nextLogRow + 1, 3).value = fullPathXlsm
            End With
        End If
    Next saveRow
    
    ' Notify user
    MsgBox "Export completed!" & vbCrLf & _
           "Version " & versionStr & " - " & Format(Now, "yyyy-mm-dd hh:mm:ss") & vbCrLf & _
           "Log entries written to Export Log sheet", _
           vbInformation, "VBA Export Complete"
    
End Sub

