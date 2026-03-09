'==============================================================================
' Class: clsZ80Encoder
' Purpose: Concrete implementation of CPUEncoder for Zilog Z80 processor
' Supports: Shadow registers, offset-based operations, block operations
' Inherits from clsCPUEncoder and implements Z80-specific instruction encoding
'==============================================================================

Option Explicit

'==============================================================================
' CPUName Property
' Returns the name of this CPU
'==============================================================================
Public Property Get CPUName() As String
    CPUName = "Z80"
End Property

'==============================================================================
' OpcodeSheetName Property
' Returns the worksheet name containing Z80 opcode mappings
'==============================================================================
Public Property Get OpcodeSheetName() As String
    OpcodeSheetName = "Z80 Op to Hex"
End Property

'==============================================================================
' EncodeInstruction - Z80 Specific Implementation
' Converts Z80 instruction mnemonic and operands to machine bytes
' Handles: Prefix opcodes (CB, ED, DD, FD), offsets, shadow registers
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
    
    ' Check for prefix opcodes (Z80 specific)
    ' Note: opByte may include prefix already from opcode table
    ' Format examples: CB 47, DD 7E, ED B0
    ' The hex column in the table should contain the full encoding
    
    Dim i As Long
    
    ' For multi-byte instructions (prefix + base)
    If byteCount > 1 Then
        ' First check if this is a prefix instruction
        If mnem = "BIT" Or mnem = "RES" Or mnem = "SET" Then
            ' CB prefix instructions (bit operations)
            out(0) = &HCB&
            out(1) = opByte And &HFF&
            EncodeInstruction = out
            Exit Function
        End If
        
        If mnem = "LDIR" Or mnem = "LDDR" Or mnem = "CPIR" Or mnem = "CPDR" Then
            ' ED prefix block operations
            out(0) = &HED&
            out(1) = opByte And &HFF&
            EncodeInstruction = out
            Exit Function
        End If
        
        If Left$(op1, 2) = "IX" Or Left$(op1, 2) = "IY" Then
            ' DD/FD prefix for index register operations
            Dim prefix As Byte
            prefix = IIf(Left$(op1, 2) = "IX", &HDD&, &HFD&)
            
            out(0) = prefix
            
            If byteCount = 3 Then
                ' Offset-based operation: prefix, base opcode, offset
                out(1) = opByte And &HFF&
                ' Offset is in op2, parse as signed byte
                Dim offset As Integer
                offset = usrHexToDec(op2)
                If offset > 127 Then offset = offset - 256 ' Convert to signed
                out(2) = offset And &HFF&
            Else
                ' No offset, just base opcode
                out(1) = opByte And &HFF&
            End If
            
            EncodeInstruction = out
            Exit Function
        End If
        
        If mnem = "EX" And op1 = "AF" And op2 = "AF'" Then
            ' Shadow register exchange: EX AF, AF'
            out(0) = &H08&
            EncodeInstruction = out
            Exit Function
        End If
        
        If mnem = "EXX" Then
            ' Shadow register exchange all: EXX
            out(0) = &HD9&
            EncodeInstruction = out
            Exit Function
        End If
        
        ' Relative jump (JR, DJNZ)
        If mnem = "JR" Or mnem = "DJNZ" Then
            out(0) = opByte And &HFF&
            ' op1 or op2 contains the relative offset
            Dim relOffset As Integer
            If mnem = "DJNZ" Then
                relOffset = usrHexToDec(op1)
            Else
                relOffset = usrHexToDec(op2)
            End If
            ' Convert to signed offset
            If relOffset > 127 Then relOffset = relOffset - 256
            out(1) = relOffset And &HFF&
            EncodeInstruction = out
            Exit Function
        End If
    End If
    
    ' Fallback to standard encoding (same as 8080)
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
            ' Two 8-bit values
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
' ValidateOperand - Z80 Specific Validation
' Includes validation for index registers, bit positions, offsets
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
            
        Case "OFFSET"
            ' Signed offset (-128 to +127)
            Dim offset As Integer
            offset = usrHexToDec(value)
            If offset > 127 Then offset = offset - 256
            ValidateOperand = (offset >= -128 And offset <= 127)
            
        Case "BIT"
            ' Bit position (0-7)
            Dim bitPos As Long
            bitPos = usrHexToDec(value)
            ValidateOperand = (bitPos >= 0 And bitPos <= 7)
            
        Case "REGISTER_INDEX"
            ' Index register: IX or IY
            ValidateOperand = (UCase$(value) = "IX" Or UCase$(value) = "IY")
            
        Case Else
            ' Default: assume valid
            ValidateOperand = True
    End Select
    
End Function

'==============================================================================
' Helper: ParseIndexRegisterWithOffset
' Extracts index register and offset from expressions like "(IX+5)" or "(IY-10)"
' Returns: Array(register, offset) or Empty if invalid
'==============================================================================
Public Function ParseIndexRegisterWithOffset(ByVal expr As String) As Variant
    
    expr = Trim$(expr)
    
    ' Check for parentheses
    If Left$(expr, 1) <> "(" Or Right$(expr, 1) <> ")" Then
        ParseIndexRegisterWithOffset = Empty
        Exit Function
    End If
    
    ' Remove parentheses
    expr = Mid$(expr, 2, Len(expr) - 2)
    
    ' Look for + or - operator
    Dim plusPos As Long, minusPos As Long
    plusPos = InStr(expr, "+")
    minusPos = InStr(expr, "-", 1)
    
    Dim regPart As String, offsetPart As String
    Dim opPos As Long
    
    If plusPos > 0 Then
        opPos = plusPos
    ElseIf minusPos > 0 Then
        opPos = minusPos
    Else
        ParseIndexRegisterWithOffset = Empty
        Exit Function
    End If
    
    regPart = Trim$(Left$(expr, opPos - 1))
    offsetPart = Trim$(Mid$(expr, opPos))
    
    ' Validate register
    regPart = UCase$(regPart)
    If regPart <> "IX" And regPart <> "IY" Then
        ParseIndexRegisterWithOffset = Empty
        Exit Function
    End If
    
    ' Parse offset
    Dim offset As Long
    On Error GoTo InvalidOffset
    offset = usrHexToDec(offsetPart)
    On Error GoTo 0
    
    ' Return as array
    ParseIndexRegisterWithOffset = Array(regPart, offset)
    Exit Function
    
InvalidOffset:
    ParseIndexRegisterWithOffset = Empty
    
End Function

