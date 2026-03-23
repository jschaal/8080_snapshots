Option Explicit
'================================================================================
' Module:  VBATests6510
' Purpose: Direct VBA unit tests for cls6510CPU instruction implementations.
'          Mirrors the VBATests pattern - call VBA6510RunAllTests from
'          VBARunAllTests in the VBATests module.
'
' HOW TO WIRE IN:
'   1. Add this module to the project (Insert > Module, name it VBATests6510).
'   2. In VBATests.VBARunAllTests, add:
'          VBA6510RunAllTests
'   3. The shared pTestCount / pTestFailureCount / pTestError counters in
'      VBATests are Private, so this module uses its own counters and rolls
'      them up via VBA6510TestCount / VBA6510FailureCount at the end.
'      Alternatively - see note at bottom for making them Friend/Public.
'================================================================================

Private p6510TestCount    As Long
Private p6510FailureCount As Long
Private p6510Error        As Boolean
Private p6510Header       As String

'================================================================================
' Public interface (called by VBARunAllTests in VBATests)
'================================================================================
Public Sub VBA6510RunAllTests()
    p6510TestCount    = 0
    p6510FailureCount = 0
    p6510Error        = False

    Debug.Print ""
    Debug.Print "+--------------------------------------------------------------------------------+"
    Debug.Print "|              6510 CPU EMULATOR - VBA UNIT TEST SUITE                          |"
    Debug.Print "+--------------------------------------------------------------------------------+"

    ' Load / Store
    Test6510_LDA
    Test6510_LDX
    Test6510_LDY
    Test6510_STA
    Test6510_STX
    Test6510_STY

    ' Transfers
    Test6510_Transfers

    ' Arithmetic
    Test6510_ADC
    Test6510_SBC
    Test6510_INC_DEC

    ' Logic
    Test6510_AND
    Test6510_ORA
    Test6510_EOR

    ' Shifts and rotates
    Test6510_ShiftsRotates

    ' Flags
    Test6510_FlagOps

    ' Stack
    Test6510_Stack

    ' Compare
    Test6510_Compare

    Debug.Print ""
    Debug.Print "6510 VBA Tests: " & p6510TestCount & " assertions, " & _
                p6510FailureCount & " failures."
    Debug.Print "+--------------------------------------------------------------------------------+"
End Sub

Public Function VBA6510TestCount() As Long
    VBA6510TestCount = p6510TestCount
End Function

Public Function VBA6510FailureCount() As Long
    VBA6510FailureCount = p6510FailureCount
End Function

Public Function VBA6510Passed() As Boolean
    VBA6510Passed = Not p6510Error
End Function

'================================================================================
' Helpers
'================================================================================
Private Function NewCPU() As cls6510cpu
    Set NewCPU = New cls6510cpu
End Function

Private Sub Chk(ByVal expected As Long, ByVal actual As Long, ByVal name As String)
    p6510TestCount = p6510TestCount + 1
    If expected <> actual Then
        Debug.Print "  FAIL: " & name & " expected=" & expected & " got=" & actual
        p6510FailureCount = p6510FailureCount + 1
        p6510Error = True
    End If
End Sub

Private Sub ChkFlag(ByVal expected As Long, ByVal actual As Long, ByVal flagName As String, ByVal ctx As String)
    Chk expected, actual, ctx & " [" & flagName & "]"
End Sub

'================================================================================
' LDA
'================================================================================
Private Sub Test6510_LDA()
    p6510Header = "LDA"
    Dim cpu As cls6510cpu

    ' Immediate: LDA #$AA
    Set cpu = NewCPU
    cpu.LDA "#AA"
    Chk &HAA, cpu.Reg("A"),  "LDA #AA -> A=$AA"
    ChkFlag 1, cpu.Flag("N"), "N", "LDA #AA"  ' bit 7 set -> negative
    ChkFlag 0, cpu.Flag("Z"), "Z", "LDA #AA"

    ' Immediate: LDA #$00 sets Zero flag
    Set cpu = NewCPU
    cpu.LDA "#00"
    Chk 0, cpu.Reg("A"), "LDA #00 -> A=0"
    ChkFlag 1, cpu.Flag("Z"), "Z", "LDA #00"
    ChkFlag 0, cpu.Flag("N"), "N", "LDA #00"

    ' Immediate: LDA #$7F -> positive, no N
    Set cpu = NewCPU
    cpu.LDA "#7F"
    Chk &H7F, cpu.Reg("A"), "LDA #7F -> A=$7F"
    ChkFlag 0, cpu.Flag("N"), "N", "LDA #7F"
    ChkFlag 0, cpu.Flag("Z"), "Z", "LDA #7F"
End Sub

'================================================================================
' LDX
'================================================================================
Private Sub Test6510_LDX()
    p6510Header = "LDX"
    Dim cpu As cls6510cpu

    Set cpu = NewCPU
    cpu.LDX "#01"
    Chk 1, cpu.Reg("X"), "LDX #01 -> X=1"
    ChkFlag 0, cpu.Flag("N"), "N", "LDX #01"
    ChkFlag 0, cpu.Flag("Z"), "Z", "LDX #01"

    Set cpu = NewCPU
    cpu.LDX "#00"
    Chk 0, cpu.Reg("X"), "LDX #00 -> X=0"
    ChkFlag 1, cpu.Flag("Z"), "Z", "LDX #00"

    Set cpu = NewCPU
    cpu.LDX "#80"
    Chk &H80, cpu.Reg("X"), "LDX #80 -> X=$80"
    ChkFlag 1, cpu.Flag("N"), "N", "LDX #80"
End Sub

'================================================================================
' LDY
'================================================================================
Private Sub Test6510_LDY()
    p6510Header = "LDY"
    Dim cpu As cls6510cpu

    Set cpu = NewCPU
    cpu.LDY "#FF"
    Chk &HFF, cpu.Reg("Y"), "LDY #FF -> Y=$FF"
    ChkFlag 1, cpu.Flag("N"), "N", "LDY #FF"

    Set cpu = NewCPU
    cpu.LDY "#00"
    ChkFlag 1, cpu.Flag("Z"), "Z", "LDY #00"
End Sub

'================================================================================
' STA / STX / STY
' Note: gMemory window starts at MemStart (default $0100).
' Zero-page addresses ($00-$FF) are below the window, so we use
' absolute addresses within the window for these VBA tests.
' Sheet-based integration tests can use ORG $0000 to test zero-page.
'================================================================================
Private Sub Test6510_STA()
    p6510Header = "STA"
    Dim cpu As cls6510cpu

    Set cpu = NewCPU
    cpu.LDA "#AB"
    cpu.STA "0150"    ' absolute address within MemStart window
    Chk &HAB, CLng(gMemory.addr(&H150)), "STA $0150 -> mem[$150]=$AB"
End Sub

Private Sub Test6510_STX()
    p6510Header = "STX"
    Dim cpu As cls6510cpu

    Set cpu = NewCPU
    cpu.LDX "#CD"
    cpu.STX "0151"
    Chk &HCD, CLng(gMemory.addr(&H151)), "STX $0151 -> mem[$151]=$CD"
End Sub

Private Sub Test6510_STY()
    p6510Header = "STY"
    Dim cpu As cls6510cpu

    Set cpu = NewCPU
    cpu.LDY "#EF"
    cpu.STY "0152"
    Chk &HEF, CLng(gMemory.addr(&H152)), "STY $0152 -> mem[$152]=$EF"
End Sub

'================================================================================
' Transfers: TAX, TAY, TXA, TYA, TSX, TXS
'================================================================================
Private Sub Test6510_Transfers()
    p6510Header = "Transfers"
    Dim cpu As cls6510cpu

    ' TAX
    Set cpu = NewCPU
    cpu.LDA "#42"
    cpu.TAX
    Chk &H42, cpu.Reg("X"), "TAX: X=A=$42"
    ChkFlag 0, cpu.Flag("Z"), "Z", "TAX $42"
    ChkFlag 0, cpu.Flag("N"), "N", "TAX $42"

    ' TAX with zero
    Set cpu = NewCPU
    cpu.LDA "#00"
    cpu.TAX
    ChkFlag 1, cpu.Flag("Z"), "Z", "TAX $00"

    ' TAX with negative
    Set cpu = NewCPU
    cpu.LDA "#FF"
    cpu.TAX
    ChkFlag 1, cpu.Flag("N"), "N", "TAX $FF"

    ' TAY
    Set cpu = NewCPU
    cpu.LDA "#55"
    cpu.TAY
    Chk &H55, cpu.Reg("Y"), "TAY: Y=A=$55"

    ' TXA
    Set cpu = NewCPU
    cpu.LDX "#33"
    cpu.TXA
    Chk &H33, cpu.Reg("A"), "TXA: A=X=$33"

    ' TYA
    Set cpu = NewCPU
    cpu.LDY "#77"
    cpu.TYA
    Chk &H77, cpu.Reg("A"), "TYA: A=Y=$77"

    ' TXS / TSX (SP lives on page 1, TXS does not affect flags)
    Set cpu = NewCPU
    cpu.LDX "#80"
    cpu.TXS
    Chk &H80, cpu.Reg("SP"), "TXS: SP=X=$80"

    Set cpu = NewCPU
    cpu.LDX "#80"
    cpu.TXS
    cpu.LDX "#00"   ' clobber X
    cpu.TSX
    Chk &H80, cpu.Reg("X"), "TSX: X=SP=$80"
    ChkFlag 1, cpu.Flag("N"), "N", "TSX $80"
End Sub

'================================================================================
' ADC
'================================================================================
Private Sub Test6510_ADC()
    p6510Header = "ADC"
    Dim cpu As cls6510cpu

    ' Basic: no carry in, no carry out
    Set cpu = NewCPU
    cpu.LDA "#10"
    cpu.CLC
    cpu.ADC "#20"
    Chk &H30, cpu.Reg("A"), "ADC #10+#20=$30"
    ChkFlag 0, cpu.Flag("C"), "C", "ADC no carry out"
    ChkFlag 0, cpu.Flag("V"), "V", "ADC no overflow"

    ' Carry out
    Set cpu = NewCPU
    cpu.LDA "#FF"
    cpu.CLC
    cpu.ADC "#01"
    Chk 0, cpu.Reg("A"), "ADC $FF+$01=0 (carry)"
    ChkFlag 1, cpu.Flag("C"), "C", "ADC carry out"
    ChkFlag 1, cpu.Flag("Z"), "Z", "ADC zero result"

    ' Carry in
    Set cpu = NewCPU
    cpu.LDA "#10"
    cpu.SEC             ' set carry
    cpu.ADC "#10"
    Chk &H21, cpu.Reg("A"), "ADC with carry-in: $10+$10+1=$21"

    ' Overflow: positive + positive = negative
    Set cpu = NewCPU
    cpu.LDA "#50"
    cpu.CLC
    cpu.ADC "#50"
    Chk &HA0, cpu.Reg("A"), "ADC overflow: $50+$50=$A0"
    ChkFlag 1, cpu.Flag("V"), "V", "ADC overflow set"
    ChkFlag 1, cpu.Flag("N"), "N", "ADC result negative"

    ' Overflow: negative + negative = positive
    Set cpu = NewCPU
    cpu.LDA "#D0"   ' -48 signed
    cpu.CLC
    cpu.ADC "#90"   ' -112 signed
    ChkFlag 1, cpu.Flag("V"), "V", "ADC neg+neg overflow"
    ChkFlag 1, cpu.Flag("C"), "C", "ADC neg+neg carry"
End Sub

'================================================================================
' SBC  (6510: C=1 means no borrow - always SEC before SBC)
'================================================================================
Private Sub Test6510_SBC()
    p6510Header = "SBC"
    Dim cpu As cls6510cpu

    ' Basic: $50 - $10 = $40, no borrow
    Set cpu = NewCPU
    cpu.LDA "#50"
    cpu.SEC
    cpu.SBC "#10"
    Chk &H40, cpu.Reg("A"), "SBC $50-$10=$40"
    ChkFlag 1, cpu.Flag("C"), "C", "SBC no borrow (C=1)"
    ChkFlag 0, cpu.Flag("V"), "V", "SBC no overflow"

    ' Borrow: $10 - $20 = $F0
    Set cpu = NewCPU
    cpu.LDA "#10"
    cpu.SEC
    cpu.SBC "#20"
    Chk &HF0, cpu.Reg("A"), "SBC $10-$20=$F0 (borrow)"
    ChkFlag 0, cpu.Flag("C"), "C", "SBC borrow (C=0)"

    ' Borrow-in (C=0 means extra borrow)
    Set cpu = NewCPU
    cpu.LDA "#50"
    cpu.CLC             ' C=0 -> extra borrow
    cpu.SBC "#10"
    Chk &H3F, cpu.Reg("A"), "SBC with borrow-in: $50-$10-1=$3F"

    ' Overflow: positive - negative = negative
    Set cpu = NewCPU
    cpu.LDA "#50"
    cpu.SEC
    cpu.SBC "#B0"   ' $50 - (-$50) should overflow
    ChkFlag 1, cpu.Flag("V"), "V", "SBC overflow pos-neg"
End Sub

'================================================================================
' INC / DEC
'================================================================================
Private Sub Test6510_INC_DEC()
    p6510Header = "INC/DEC"
    Dim cpu As cls6510cpu

    ' INX
    Set cpu = NewCPU
    cpu.LDX "#41"
    cpu.INX
    Chk &H42, cpu.Reg("X"), "INX $41->$42"
    ChkFlag 0, cpu.Flag("Z"), "Z", "INX not zero"

    ' INX wrap
    Set cpu = NewCPU
    cpu.LDX "#FF"
    cpu.INX
    Chk 0, cpu.Reg("X"), "INX $FF->$00 wrap"
    ChkFlag 1, cpu.Flag("Z"), "Z", "INX wrap zero"

    ' INY
    Set cpu = NewCPU
    cpu.LDY "#7F"
    cpu.INY
    Chk &H80, cpu.Reg("Y"), "INY $7F->$80"
    ChkFlag 1, cpu.Flag("N"), "N", "INY into negative"

    ' DEX
    Set cpu = NewCPU
    cpu.LDX "#01"
    cpu.DEX
    Chk 0, cpu.Reg("X"), "DEX $01->$00"
    ChkFlag 1, cpu.Flag("Z"), "Z", "DEX to zero"

    ' DEX wrap
    Set cpu = NewCPU
    cpu.LDX "#00"
    cpu.DEX
    Chk &HFF, cpu.Reg("X"), "DEX $00->$FF wrap"
    ChkFlag 1, cpu.Flag("N"), "N", "DEX wrap negative"

    ' DEY
    Set cpu = NewCPU
    cpu.LDY "#80"
    cpu.DEY
    Chk &H7F, cpu.Reg("Y"), "DEY $80->$7F"
    ChkFlag 0, cpu.Flag("N"), "N", "DEY into positive"
End Sub

'================================================================================
' AND / ORA / EOR
'================================================================================
Private Sub Test6510_AND()
    p6510Header = "AND"
    Dim cpu As cls6510cpu

    Set cpu = NewCPU
    cpu.LDA "#FF"
    cpu.AND_ "#0F"
    Chk &H0F, cpu.Reg("A"), "AND $FF&#0F=$0F"
    ChkFlag 0, cpu.Flag("N"), "N", "AND positive"

    Set cpu = NewCPU
    cpu.LDA "#AA"
    cpu.AND_ "#55"
    Chk 0, cpu.Reg("A"), "AND $AA&$55=0"
    ChkFlag 1, cpu.Flag("Z"), "Z", "AND zero"
    ChkFlag 0, cpu.Flag("C"), "C", "AND C unaffected (stays 0)"
End Sub

Private Sub Test6510_ORA()
    p6510Header = "ORA"
    Dim cpu As cls6510cpu

    Set cpu = NewCPU
    cpu.LDA "#F0"
    cpu.ORA "#0F"
    Chk &HFF, cpu.Reg("A"), "ORA $F0|$0F=$FF"
    ChkFlag 1, cpu.Flag("N"), "N", "ORA negative"

    Set cpu = NewCPU
    cpu.LDA "#00"
    cpu.ORA "#00"
    ChkFlag 1, cpu.Flag("Z"), "Z", "ORA zero"
End Sub

Private Sub Test6510_EOR()
    p6510Header = "EOR"
    Dim cpu As cls6510cpu

    Set cpu = NewCPU
    cpu.LDA "#FF"
    cpu.EOR "#FF"
    Chk 0, cpu.Reg("A"), "EOR $FF^$FF=0"
    ChkFlag 1, cpu.Flag("Z"), "Z", "EOR zero"

    Set cpu = NewCPU
    cpu.LDA "#AA"
    cpu.EOR "#55"
    Chk &HFF, cpu.Reg("A"), "EOR $AA^$55=$FF"
    ChkFlag 1, cpu.Flag("N"), "N", "EOR negative"
End Sub

'================================================================================
' Shifts and Rotates: ASL, LSR, ROL, ROR
'================================================================================
Private Sub Test6510_ShiftsRotates()
    p6510Header = "Shifts/Rotates"
    Dim cpu As cls6510cpu

    ' ASL A: shift left, bit 7 -> C, bit 0 = 0
    Set cpu = NewCPU
    cpu.LDA "#81"   ' 1000 0001
    cpu.ASL "A"
    Chk &H02, cpu.Reg("A"), "ASL $81->$02"
    ChkFlag 1, cpu.Flag("C"), "C", "ASL carry from bit7"

    Set cpu = NewCPU
    cpu.LDA "#40"
    cpu.ASL "A"
    Chk &H80, cpu.Reg("A"), "ASL $40->$80"
    ChkFlag 1, cpu.Flag("N"), "N", "ASL into negative"
    ChkFlag 0, cpu.Flag("C"), "C", "ASL no carry"

    ' LSR A: shift right, bit 0 -> C, bit 7 = 0
    Set cpu = NewCPU
    cpu.LDA "#81"   ' 1000 0001
    cpu.LSR "A"
    Chk &H40, cpu.Reg("A"), "LSR $81->$40"
    ChkFlag 1, cpu.Flag("C"), "C", "LSR carry from bit0"
    ChkFlag 0, cpu.Flag("N"), "N", "LSR always clears N"

    ' ROL A: rotate left through carry
    Set cpu = NewCPU
    cpu.LDA "#40"
    cpu.CLC
    cpu.ROL "A"
    Chk &H80, cpu.Reg("A"), "ROL $40 C=0 -> $80"
    ChkFlag 0, cpu.Flag("C"), "C", "ROL no carry out"

    Set cpu = NewCPU
    cpu.LDA "#80"
    cpu.CLC
    cpu.ROL "A"
    Chk 0, cpu.Reg("A"), "ROL $80 C=0 -> $00"
    ChkFlag 1, cpu.Flag("C"), "C", "ROL carry out from bit7"

    Set cpu = NewCPU
    cpu.LDA "#00"
    cpu.SEC
    cpu.ROL "A"
    Chk 1, cpu.Reg("A"), "ROL $00 C=1 -> $01"

    ' ROR A: rotate right through carry
    Set cpu = NewCPU
    cpu.LDA "#01"
    cpu.CLC
    cpu.ROR "A"
    Chk 0, cpu.Reg("A"), "ROR $01 C=0 -> $00"
    ChkFlag 1, cpu.Flag("C"), "C", "ROR carry out from bit0"

    Set cpu = NewCPU
    cpu.LDA "#00"
    cpu.SEC
    cpu.ROR "A"
    Chk &H80, cpu.Reg("A"), "ROR $00 C=1 -> $80"
    ChkFlag 1, cpu.Flag("N"), "N", "ROR C into bit7"
End Sub

'================================================================================
' Flag operations: CLC, SEC, CLV, CLD, SED, CLI, SEI
'================================================================================
Private Sub Test6510_FlagOps()
    p6510Header = "Flag ops"
    Dim cpu As cls6510cpu

    Set cpu = NewCPU
    cpu.SEC: ChkFlag 1, cpu.Flag("C"), "C", "SEC"
    cpu.CLC: ChkFlag 0, cpu.Flag("C"), "C", "CLC"

    Set cpu = NewCPU
    cpu.LDA "#50": cpu.CLC: cpu.ADC "#50"  ' force V=1
    ChkFlag 1, cpu.Flag("V"), "V", "V set before CLV"
    cpu.CLV
    ChkFlag 0, cpu.Flag("V"), "V", "CLV clears V"

    Set cpu = NewCPU
    cpu.SEI: ChkFlag 1, cpu.Flag("I"), "I", "SEI"
    cpu.CLI: ChkFlag 0, cpu.Flag("I"), "I", "CLI"

    Set cpu = NewCPU
    cpu.SED: ChkFlag 1, cpu.Flag("D"), "D", "SED"
    cpu.CLD: ChkFlag 0, cpu.Flag("D"), "D", "CLD"
End Sub

'================================================================================
' Stack: PHA / PLA / PHP / PLP
'================================================================================
Private Sub Test6510_Stack()
    p6510Header = "Stack"
    Dim cpu As cls6510cpu

    ' PHA / PLA round-trip
    Set cpu = NewCPU
    cpu.LDA "#AB"
    cpu.PHA
    Chk &HFE, cpu.Reg("SP"), "PHA decrements SP to $FE"
    cpu.LDA "#00"           ' clobber A
    cpu.PLA
    Chk &HAB, cpu.Reg("A"), "PLA restores A=$AB"
    ChkFlag 0, cpu.Flag("Z"), "Z", "PLA A=$AB not zero"
    ChkFlag 1, cpu.Flag("N"), "N", "PLA A=$AB negative"
    Chk &HFF, cpu.Reg("SP"), "PLA restores SP to $FF"

    ' PHP / PLP round-trip
    Set cpu = NewCPU
    cpu.SEC     ' C=1
    cpu.SEI     ' I=1
    cpu.PHP
    cpu.CLC     ' clobber flags
    cpu.CLI
    cpu.PLP
    ChkFlag 1, cpu.Flag("C"), "C", "PLP restores C"
    ChkFlag 1, cpu.Flag("I"), "I", "PLP restores I"

    ' Multiple pushes
    Set cpu = NewCPU
    cpu.LDA "#11": cpu.PHA
    cpu.LDA "#22": cpu.PHA
    cpu.LDA "#33": cpu.PHA
    cpu.PLA: Chk &H33, cpu.Reg("A"), "Stack LIFO: 3rd push"
    cpu.PLA: Chk &H22, cpu.Reg("A"), "Stack LIFO: 2nd push"
    cpu.PLA: Chk &H11, cpu.Reg("A"), "Stack LIFO: 1st push"
End Sub

'================================================================================
' Compare: CMP, CPX, CPY
'================================================================================
Private Sub Test6510_Compare()
    p6510Header = "Compare"
    Dim cpu As cls6510cpu

    ' CMP: A = operand -> Z=1, C=1
    Set cpu = NewCPU
    cpu.LDA "#50"
    cpu.CMP "#50"
    Chk &H50, cpu.Reg("A"), "CMP A unchanged"
    ChkFlag 1, cpu.Flag("Z"), "Z", "CMP equal Z=1"
    ChkFlag 1, cpu.Flag("C"), "C", "CMP equal C=1"
    ChkFlag 0, cpu.Flag("N"), "N", "CMP equal N=0"

    ' CMP: A > operand -> Z=0, C=1
    Set cpu = NewCPU
    cpu.LDA "#60"
    cpu.CMP "#50"
    ChkFlag 0, cpu.Flag("Z"), "Z", "CMP greater Z=0"
    ChkFlag 1, cpu.Flag("C"), "C", "CMP greater C=1"

    ' CMP: A < operand -> Z=0, C=0
    Set cpu = NewCPU
    cpu.LDA "#40"
    cpu.CMP "#50"
    ChkFlag 0, cpu.Flag("Z"), "Z", "CMP less Z=0"
    ChkFlag 0, cpu.Flag("C"), "C", "CMP less C=0"
    ChkFlag 1, cpu.Flag("N"), "N", "CMP less N=1"

    ' CPX
    Set cpu = NewCPU
    cpu.LDX "#30"
    cpu.CPX "#30"
    ChkFlag 1, cpu.Flag("Z"), "Z", "CPX equal Z=1"
    ChkFlag 1, cpu.Flag("C"), "C", "CPX equal C=1"

    ' CPY
    Set cpu = NewCPU
    cpu.LDY "#20"
    cpu.CPY "#10"
    ChkFlag 0, cpu.Flag("Z"), "Z", "CPY greater Z=0"
    ChkFlag 1, cpu.Flag("C"), "C", "CPY greater C=1"
End Sub
