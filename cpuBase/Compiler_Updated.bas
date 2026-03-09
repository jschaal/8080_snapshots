'==============================================================================
' Module: Compiler (Updated for Class-Based Architecture)
' Purpose: Assemble CPU code to machine bytes using extensible encoder classes
' Supports: 8080, Z80 (and easily extensible to 6502, etc)
'==============================================================================

Option Explicit

Private Const BYTES_PER_MEMROW As Long = 8
Private Const EMPTY_RUN_LIMIT As Long = 8
Private gMemTableCol As Long
Private pLabels As New clsLabels

'==============================================================================
' AssembleToMachine - Main Entry Point
' Detects CPU type and uses appropriate encoder
'==============================================================================
Public Sub AssembleToMachine()
    
    Dim wsCPU As Worksheet
    Set wsCPU = ThisWorkbook.Worksheets("CPU")
    
    ' Detect CPU type - you can customize this detection
    Dim cpuType As String
    cpuType = GetSelectedCPU()
    
    ' Create appropriate encoder
    Dim encoder As Object
    Select Case cpuType
        Case "8080"
            Set encoder = New cls8080Encoder
        Case "Z80"
            Set encoder = New clsZ80Encoder
        Case Else
            MsgBox "Unknown CPU type: " & cpuType, vbExclamation
            Exit Sub
    End Select
    
    ' Run compilation with the encoder
    CompileWithEncoder wsCPU, encoder

End Sub

'==============================================================================
' GetSelectedCPU
' Returns the CPU type to use
' Customize this to read from your UI or named range
'==============================================================================
Private Function GetSelectedCPU() As String
    
    ' Option 1: Read from a named range
    ' Set up a named range "SelectedCPU" with value "8080" or "Z80"
    On Error Resume Next
    GetSelectedCPU = CStr(ThisWorkbook.Range("SelectedCPU").value)
    On Error GoTo 0
    
    ' Option 2: Default to 8080 if not found
    If GetSelectedCPU = "" Then
        GetSelectedCPU = "8080"
    End If
    
End Function

'==============================================================================
' CompileWithEncoder
' Core compilation logic using an encoder instance
'==============================================================================
Private Sub CompileWithEncoder(ByVal wsCPU As Worksheet, ByVal encoder As Object)
    
    Dim wsOp As Worksheet
    ' Get opcode sheet from encoder
    Set wsOp = ThisWorkbook.Worksheets(encoder.OpcodeSheetName)
    
    Dim myMemory As clsMemory
    Dim lbl As clsLabelRecord
    
    usrMemReset
    ResetAddressList
    Set myMemory = gMemory
    
    ' Offsets into the CPU program grid
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
    
    ' Build opcode lookup dictionary using encoder
    Dim opMap As Object
    Set opMap = encoder.BuildOpcodeMap(wsOp)
    
    ' Labels/EQU/DB metadata
    Dim lbls As clsLabels
    Set lbls = New clsLabels
    
    ' MemoryTable byte cells
    Dim rngMem As Range
    Set rngMem = wsCPU.Range("MemoryTable")
    
    Application.Calculate
    Application.ScreenUpdating = False
    
    ' Clear prior compiled bytes
    rngMem.ClearContents
    wsCPU.Range(wsCPU.Cells(rngMem.row, gMemTableCol), _
                wsCPU.Cells(rngMem.row + rngMem.Rows.Count - 1, gMemTableCol)).ClearContents
    
    Dim base As Range
    Set base = wsCPU.Range("Line0")
    
    Dim pc As Long
    pc = memStartDec
    
    Dim r As Long, emptyRun As Long
    emptyRun = 0
    
    ' Walk listing rows
    For r = 0 To memSize
        
        Dim cur As Range
        Set cur = base.offset(r, 0)
        
        Dim labelName As String, opcode As String, op1 As String, op2 As String
        labelName = UCase$(Trim$(CStr(cur.offset(0, ofsLabel).value)))
        opcode = UCase$(Trim$(CStr(cur.offset(0, ofsOpcode).value)))
        op1 = StripComment(Trim$(CStr(cur.offset(0, ofsOp1).value)))
        op2 = StripComment(Trim$(CStr(cur.offset(0, ofsOp2).value)))
        
        ' Stop condition: skip lines starting with ;
        If (labelName = ";") Then
            GoTo NextRow
        End If
        
        ' Stop condition: consecutive empty rows
        If (labelName = "") And (opcode = "") And (op1 = "") And (op2 = "") Then
            emptyRun = emptyRun + 1
            If emptyRun >= EMPTY_RUN_LIMIT Then Exit For
            GoTo NextRow
        Else
            emptyRun = 0
        End If
        
        ' Resolve labels
        Set lbl = pLabels.GetLabel(op1)
        If Not lbl Is Nothing Then
            op1 = lbl.clsAddressHex
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
                GoTo NextRow
                
            Case "DB"
                If Len(CStr(cur.offset(0, ofsOp1).value)) = 0 Then
                    RaiseAsmError wsCPU, "DB missing string", cur
                    Exit Sub
                End If
                bytesOut = StringToBytesVariant(CStr(cur.offset(0, ofsOp1).value))
                
            Case "DS"
                Dim ds As Long, i As Long
                If Not usrValIsHexString(cur.offset(0, ofsOp1).value) Then
                    RaiseAsmError wsCPU, "DS Invalid length", cur
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
                ' Use encoder to look up and encode instruction
                Dim info As Variant
                info = encoder.FindOpcode(opMap, opcode, UCase$(op1), UCase$(op2))
                If IsEmpty(info) Then
                    RaiseAsmError wsCPU, "Unknown instruction: " & opcode & " " & op1 & " " & op2, cur
                    Exit Sub
                End If
                
                bytesOut = encoder.EncodeInstruction(lbls, CByte(info(0)), CLng(info(1)), _
                                                    CStr(info(2)), CStr(info(3)), _
                                                    UCase$(op1), UCase$(op2), opcode, cur, wsCPU)
        End Select
        
        ' Emit compiled bytes to MemoryTable
        If Not IsEmpty(bytesOut) Then
            EmitBytesToMemoryTable wsCPU, rngMem, memStartDec, pc, bytesOut, myMemory
            pc = pc + (UBound(bytesOut) - LBound(bytesOut) + 1)
        End If
        
NextRow:
    Next r
    
    Application.ScreenUpdating = True
    wsCPU.Range("errMessage").value = "Assemble complete (" & encoder.CPUName & ")"

End Sub

'==============================================================================
' EmitBytesToMemoryTable
' Writes compiled bytes to MemoryTable based on PC
'==============================================================================
Private Sub EmitBytesToMemoryTable(ByVal wsCPU As Worksheet, _
 ByVal rngMem As Range, _
 ByVal memStartDec As Long, _
 ByVal startAddrDec As Long, _
 ByVal bytesOut As Variant, _
 ByRef myMem As clsMemory)
    
    Dim i As Long, baseAddr As Long, lastTargetRow As Long
    lastTargetRow = 0
    baseAddr = memStartDec
    
    For i = LBound(bytesOut) To UBound(bytesOut)
        
        Dim addr As Long, offset As Long, rowIndex As Long, colIndex As Long
        Dim targetRow As Long, targetCol As Long
        
        addr = startAddrDec + (i - LBound(bytesOut))
        offset = addr - memStartDec
        If offset < 0 Then Exit Sub
        
        rowIndex = offset \ BYTES_PER_MEMROW
        colIndex = offset Mod BYTES_PER_MEMROW
        
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

