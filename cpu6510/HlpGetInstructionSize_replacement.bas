'================================================================================
' HlpGetInstructionSize  (drop-in replacement for the existing function)
'
' HOW TO APPLY:
'   In modGlobals, select the entire existing HlpGetInstructionSize function
'   from "Public Function HlpGetInstructionSize" to its "End Function" and
'   replace it with this block.
'
' WHAT CHANGED:
'   Added an early CPUMode()="6510" branch at the top of the Select Case.
'   All 6510 mnemonics are handled there and Exit Function immediately,
'   so they never fall through to the 8080/Z80 cases below.
'   The 8080/Z80 section is unchanged.
'================================================================================
Public Function HlpGetInstructionSize( _
    ByVal mnem As String, _
    Optional ByVal op1 As String = "", _
    Optional ByVal op2 As String = "" _
) As Long

    mnem = UCase$(Trim$(mnem))
    op1  = UCase$(Trim$(op1))
    op2  = UCase$(Trim$(op2))

    Dim size As Long
    size = 1    ' default: most instructions are 1 byte

    '===========================================================================
    ' 6510 PATH
    ' Handled first so none of these mnemonics fall through to the 8080/Z80
    ' cases below (Select Case takes the first match, so order matters).
    '===========================================================================
    If CPUMode() = "6510" Then

        Select Case mnem

            ' Pseudo-ops - no bytes emitted
            Case "", "EQU", "ORG"
                size = 0

            Case "END", "HLT", "STP"
                size = 0

            ' DB: count chars in the string literal + 1 null terminator,
            '     or count comma-separated byte tokens
            Case "DB"
                If Left$(op1, 1) = Chr$(34) Or Left$(op1, 1) = "'" Then
                    size = Len(op1) - 1  ' strip 2 quotes, add 1 null = Len-1
                Else
                    size = UBound(Split(op1, ",")) - LBound(Split(op1, ",")) + 1
                End If

            ' DS: reserve n bytes (op1 is hex or decimal count)
            Case "DS"
                If usrValIsHexString(op1) Then
                    size = usrHexToDec(op1)
                ElseIf IsNumeric(op1) Then
                    size = CLng(op1)
                Else
                    size = 0
                End If

            ' All 6510 instructions: size determined by addressing mode in op1
            Case "ADC", "AND", "ASL", "BIT", "CMP", "CPX", "CPY", _
                 "DEC", "EOR", "INC", "LDA", "LDX", "LDY", _
                 "LSR", "ORA", "ROL", "ROR", "SBC", "STA", "STX", "STY"
                size = Hlp6510InstrSize(op1)

            ' Branches: always 2 bytes (opcode + signed offset)
            Case "BCC", "BCS", "BEQ", "BMI", "BNE", "BPL", "BVC", "BVS"
                size = 2

            ' JSR: always 3 bytes (opcode + 16-bit address)
            Case "JSR"
                size = 3

            ' JMP: always 3 bytes (absolute or indirect, both 3 on 6510)
            Case "JMP"
                size = 3

            ' All implied / accumulator 6510 instructions: 1 byte
            Case "BRK", "CLC", "CLD", "CLI", "CLV", "DEX", "DEY", _
                 "INX", "INY", "NOP", "PHA", "PHP", "PLA", "PLP", _
                 "RTI", "RTS", "SEC", "SED", "SEI", _
                 "TAX", "TAY", "TSX", "TXA", "TXS", "TYA"
                size = 1

            Case Else
                size = 1    ' unknown 6510 mnemonic - safe default

        End Select

        HlpGetInstructionSize = size
        Exit Function

    End If

    '===========================================================================
    ' 8080 / Z80 PATH  (unchanged from original)
    '===========================================================================
    Select Case mnem

        ' Pseudo-ops / no code emitted
        Case "", "EQU"
            size = 0

        Case "DB"
            size = Len(op1)   ' string payload already validated elsewhere

        ' 16-bit immediate
        Case "LXI", "JMP", "JC", "JNC", "JZ", "JNZ", _
             "JM", "JPE", "JPO", "CALL", "LDA", "STA", _
             "LHLD", "SHLD"
            size = 3

        ' 16-bit Calls to Address (8080)
        Case "CC", "CM", "CNC", "CNZ", "CPE", "CPO", "CZ"
            size = 3

        ' 8-bit immediate
        Case "MVI", "ADI", "ACI", "ANI", "ORI", "XRI", _
             "CPI", "SUI", "SBI", "IN", "OUT", "PRI"
            size = 2

        ' Relative jumps (Z80)
        Case "JR", "DJNZ"
            size = 2

        ' BIT / SET / RES (Z80 CB-prefixed)
        Case "BIT", "SET", "RES"
            size = 2

        ' LD family (needs operand inspection)
        Case "LD"
            size = HlpSizeLD(op1, op2)

        ' New Z80 operations
        Case "AND"
            size = HlpSizeZ80Enhanced(mnem, op1, op2)

        ' Block operations
        Case "CPD", "CPDR", "CPIR", "IND", "INDR", "INI", "INIR", _
             "LDD", "LDDR", "LDI", "LDIR", "OTDR", "OTIR", "OUTD", "OUTI"
            size = 2

        ' Z80 Shift / Negate / Bit
        Case "NEG", "RETI", "RETN", "RL", "RLD", "RR", "RRD", _
             "SLA", "SLR", "SRA", "SRL"
            size = 2

        ' Interrupt mode
        Case "IM"
            size = 2

        ' OR unification
        Case "OR"
            If usrIsRegister08Bit(op1) And (op2 = "(HL)" Or usrIsRegister08Bit(op2)) Then
                size = 1
            ElseIf op2 = "BYTE" Then
                size = 2
            Else
                size = 4
            End If

        ' Enhanced Z80 versions of 8080 opcodes
        Case "CPI"
            size = 2

        Case "ADC", "ADD", "DEC", "INC", "SBC", "SUB", "XOR"
            size = HlpSizeZ80Enhanced(mnem, op1, op2)

        Case "CP"
            If op1 <> "A" And op2 = "" Then
                size = 3
            Else
                size = HlpSizeZ80Enhanced(mnem, op1, op2)
            End If

        Case "JP"
            If op1 = "(HL)" Then
                size = 1
            ElseIf op1 = "(IX)" Or op1 = "(IY)" Then
                size = 2
            Else
                size = 3
            End If

        Case "POP", "PUSH"
            If op1 = "IX" Or op1 = "IY" Then
                size = 2
            Else
                size = 1
            End If

        Case "RLC", "RRC"
            If op1 = "" Then
                size = 1
            Else
                size = 2
            End If

        ' All others (register-register, implied)
        Case Else
            size = 1

    End Select

    HlpGetInstructionSize = size

End Function
