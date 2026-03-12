'==============================================================================
' Class: cls8080Encoder
' Purpose: Concrete implementation of CPUEncoder for Intel 8080 processor
' Inherits from clsCPUEncoder and implements 8080-specific instruction encoding
'==============================================================================

Option Explicit

'==============================================================================
' CPUName Property
' Returns the name of this CPU
'==============================================================================
Public Property Get CPUName() As String
    CPUName = "8080"
End Property

'==============================================================================
' OpcodeSheetName Property
' Returns the worksheet name containing 8080 opcode mappings
'==============================================================================
Public Property Get OpcodeSheetName() As String
    OpcodeSheetName = "8080 Op to Hex"
End Property

'==============================================================================
' BuildOpcodeMap - 8080 Implementation
' Reads 8080 opcode table and builds lookup dictionary
' Uses the standard table format from "8080 Op to Hex" sheet
'==============================================================================
Public Function BuildOpcodeMap(ByVal wsOp As Worksheet) As Object
    ' Call the base class implementation which works for standard 8080 format
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    Dim lastRow As Long
    lastRow = wsOp.Cells(wsOp.Rows.Count, 1).End(xlUp).Row
    
    Dim r As Long, emptyRun As Long
    emptyRun = 0
    
    ' Columns: A=Opcode, B=Hex, D=OP1, E=OP2, F=Bytes
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

'==============================================================================
' FindOpcode - 8080 Implementation
' Looks up an instruction in the opcode map with fallback matching
'==============================================================================
Public Function FindOpcode(ByVal opMap As Object, ByVal mnem As String, _
                          ByVal op1 As String, ByVal op2 As String) As Variant
    Dim key As String
    
    ' Try exact match
    key = mnem & "|" & op1 & "|" & op2
    If opMap.Exists(key) Then FindOpcode = opMap(key): Exit Function
    
    ' Try with only OP1
    key = mnem & "|" & op1 & "|"
    If opMap.Exists(key) Then FindOpcode = opMap(key): Exit Function
    
    ' Try with template specs
    Dim p As Variant
    For Each p In Array("BYTE", "ADDRESS", "PORT", "DATA")
        key = mnem & "|" & op1 & "|" & CStr(p)
        If opMap.Exists(key) Then FindOpcode = opMap(key): Exit Function
    Next p
    
    For Each p In Array("BYTE", "ADDRESS", "PORT", "DATA")
        key = mnem & "|" & CStr(p) & "|" & op2
        If opMap.Exists(key) Then FindOpcode = opMap(key): Exit Function
        key = mnem & "|" & CStr(p) & "|"
        If opMap.Exists(key) Then FindOpcode = opMap(key): Exit Function
    Next p
    
    ' No match found
    FindOpcode = Empty
End Function

'==============================================================================
' EncodeInstruction - 8080 Specific Implementation
' Converts 8080 instruction mnemonic and operands to machine bytes
'
' Returns a Variant array of bytes, or Empty if encoding failed
'==============================================================================
Public Function EncodeInstruction(ByRef lbls As clsLabels, _
                                 ByVal opByte As Long, ByVal byteCount As Long, _
                                 ByVal spec1 As String, ByVal spec2 As String, _
                                 ByVal op1 As String, ByVal op2 As String, _
                                 ByVal mnem As String, ByVal cur As Range, _
                                 ByVal wsCPU As Worksheet) As Variant

    Dim out() As Byte
    ReDim out(0 To byteCount - 1)
    
    ' First byte is always the opcode
    out(0) = opByte And &HFF&
    
    ' Handle operands based on specification
    If byteCount = 1 Then
        ' Single-byte instruction, no operands
        EncodeInstruction = out
        Exit Function
    End If
    
    ' 2-byte or 3-byte instruction with operands
    If byteCount = 2 Then
        ' 8-bit immediate or register-based
        If spec1 = "BYTE" Or spec1 = "DATA" Then
            out(1) = usrHexToDec(op1) And &HFF&
        End If
        EncodeInstruction = out
        Exit Function
    End If
    
    If byteCount = 3 Then
        ' 16-bit address (little-endian: low byte, high byte)
        If spec1 = "ADDRESS" Then
            Dim addr As Long
            addr = ResolveValue16(lbls, op1)
            out(1) = addr And &HFF&
            out(2) = (addr \ 256) And &HFF&
        ElseIf spec1 = "BYTE" And spec2 = "BYTE" Then
            ' Two 8-bit values (rare in 8080)
            out(1) = usrHexToDec(op1) And &HFF&
            out(2) = usrHexToDec(op2) And &HFF&
        End If
        EncodeInstruction = out
        Exit Function
    End If
    
    ' Fallback
    EncodeInstruction = out

End Function

'==============================================================================
' ValidateOperand - 8080 Specific Validation
'==============================================================================
Public Function ValidateOperand(ByVal spec As String, ByVal value As String) As Boolean
    
    Select Case spec
        Case "BYTE"
            ' 8-bit value (0-255)
            Dim byteVal As Long
            byteVal = usrHexToDec(value)
            ValidateOperand = (byteVal >= 0 And byteVal <= 255)
            
        Case "ADDRESS"
            ' 16-bit address (0-65535)
            Dim addr As Long
            addr = usrHexToDec(value)
            ValidateOperand = (addr >= 0 And addr <= 65535)
            
        Case "PORT"
            ' Port number (0-255)
            Dim port As Long
            port = usrHexToDec(value)
            ValidateOperand = (port >= 0 And port <= 255)
            
        Case Else
            ' Default: assume valid
            ValidateOperand = True
    End Select
    
End Function

