'==============================================================================
' Class: clsCPUEncoder
' Purpose: Abstract base class for CPU instruction encoders
' Allows for extensible architecture supporting multiple CPUs (8080, Z80, 6502, etc)
'==============================================================================

Option Explicit

' Properties that subclasses must implement
Public Property Get CPUName() As String
    ' Override in subclass: "8080", "Z80", "6502", etc
    Err.Raise 11 ' NotImplemented
End Property

Public Property Get OpcodeSheetName() As String
    ' Override in subclass: sheet name containing opcode mappings
    Err.Raise 11
End Property

'==============================================================================
' BuildOpcodeMap
' Reads opcode table from the appropriate CPU sheet and builds lookup dictionary
' Must be overridden by subclasses if they have different table structures
'==============================================================================
Public Function BuildOpcodeMap(ByVal wsOp As Worksheet) As Object
    ' Default implementation - override if your CPU has different table format
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
' FindOpcode
' Looks up an instruction in the opcode map with fallback matching
' Can be overridden by subclasses for specialized lookup logic
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
    For Each p In Array("BYTE", "ADDRESS", "PORT", "DATA", "OFFSET")
        key = mnem & "|" & op1 & "|" & CStr(p)
        If opMap.Exists(key) Then FindOpcode = opMap(key): Exit Function
    Next p
    
    For Each p In Array("BYTE", "ADDRESS", "PORT", "DATA", "OFFSET")
        key = mnem & "|" & CStr(p) & "|" & op2
        If opMap.Exists(key) Then FindOpcode = opMap(key): Exit Function
        key = mnem & "|" & CStr(p) & "|"
        If opMap.Exists(key) Then FindOpcode = opMap(key): Exit Function
    Next p
    
    ' No match found
    FindOpcode = Empty
End Function

'==============================================================================
' EncodeInstruction
' Converts an instruction mnemonic and operands to machine bytes
' MUST be overridden by subclasses
'==============================================================================
Public Function EncodeInstruction(ByRef lbls As Object, _
                                 ByVal opByte As Long, ByVal byteCount As Long, _
                                 ByVal spec1 As String, ByVal spec2 As String, _
                                 ByVal op1 As String, ByVal op2 As String, _
                                 ByVal mnem As String, ByVal cur As Range, _
                                 ByVal wsCPU As Worksheet) As Variant
    ' Must be overridden in subclass
    Err.Raise 11 ' NotImplemented
End Function

'==============================================================================
' ValidateOperand
' Validates an operand value is within acceptable range for this CPU
' Can be overridden by subclasses for CPU-specific validation
'==============================================================================
Public Function ValidateOperand(ByVal spec As String, ByVal value As String) As Boolean
    ' Default: assume value is valid
    ' Override for specific validation rules
    ValidateOperand = True
End Function

