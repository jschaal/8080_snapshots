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

