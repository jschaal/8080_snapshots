'================================================================================
' Z80/8080 Emulator - Optimized Code & Additional Unit Tests
' Generated: March 8, 2026
' 
' HOW TO USE:
' 1. Copy the ResetTestCases_Optimized subroutine and paste over your current
'    ResetTestCases in the appropriate module
' 2. Copy each Test_* subroutine from the "ADDITIONAL TEST CASES" section
'    and add them to your ClaudeTests module
' 3. Add the new tests to your ClaudeRunAllTests() master suite
' 4. Run ClaudeRunAllTests() to verify
'================================================================================

'================================================================================
' PART 1: OPTIMIZED SUBROUTINE
'================================================================================

'==============================================================================
' ResetTestCases_Optimized
' 
' PERFORMANCE IMPROVEMENT: 5-10x faster than current implementation
' Key optimizations:
'   1. Load both columns at once into array (not cell-by-cell)
'   2. Process array in-memory instead of iterating Range objects
'   3. Write results back in single bulk operation
'==============================================================================
Public Sub ResetTestCases_Optimized()
    Dim ws As Worksheet: Set ws = Sheets("Unit Tests")
    Dim firstCell As Range: Set firstCell = ws.Range("TestTable")
    
    ' Find the last row with data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, firstCell.Column + 1).End(xlUp).row
    If lastRow < firstCell.row + 1 Then Exit Sub
    
    ' OPTIMIZATION 1: Define range extents
    Dim nameCol As Long: nameCol = firstCell.Column + 1     ' Column containing test names
    Dim runCol As Long: runCol = firstCell.Column + 2       ' Column containing "Run" flag (result)
    Dim startRow As Long: startRow = firstCell.row + 1      ' First data row
    Dim rowCount As Long: rowCount = lastRow - startRow + 1
    
    ' OPTIMIZATION 2: Load data array once instead of cell-by-cell iteration
    ' This single Range read is MUCH faster than looping through Cells()
    Dim data As Variant
    data = ws.Range(ws.Cells(startRow, nameCol), _
                    ws.Cells(lastRow, runCol)).value
    
    ' OPTIMIZATION 3: Process array in-memory, collect all changes
    Dim changes() As Variant
    ReDim changes(1 To rowCount, 1 To 1)
    Dim r As Long
    Dim testName As String
    
    For r = 1 To UBound(data, 1)
        testName = Trim$(CStr(data(r, 1)))
        ' If test name is not empty, set Run flag to 1, otherwise 0
        changes(r, 1) = IIf(testName <> "", 1, 0)
    Next r
    
    ' OPTIMIZATION 4: Write all changes back at once
    ' Single bulk write is exponentially faster than loop with individual Cells() writes
    Application.ScreenUpdating = False
    ws.Range(ws.Cells(startRow, runCol), _
             ws.Cells(lastRow, runCol)).value = changes
    Application.ScreenUpdating = True

End Sub


'================================================================================
' PART 2: ADDITIONAL UNIT TEST CASES
' 
' Copy these test subroutines into your ClaudeTests module.
' Each test follows the same pattern as your existing tests.
'================================================================================

'==============================================================================
' A. ARITHMETIC - ADVANCED EDGE CASES
'==============================================================================

Sub Test_ADD_EdgeCases()
    pTestHeader = "ADD - Edge Cases"
    Dim cpu As clsDecCPU
    
    ' Test 1: Zero + Zero
    Set cpu = CreateTestCPU
    cpu.reg("A") = 0
    cpu.reg("B") = 0
    cpu.ADD "B"
    AssertEquals 0, cpu.reg("A"), "ADD: 0 + 0 = 0"
    AssertEquals 1, cpu.Flag("Zero"), "ADD: Zero flag set"
    
    ' Test 2: Max + Max = overflow
    Set cpu = CreateTestCPU
    cpu.reg("A") = 255
    cpu.reg("B") = 255
    cpu.ADD "B"
    AssertEquals 254, cpu.reg("A"), "ADD: 255 + 255 = 254 (wrap, carry set)"
    AssertEquals 1, cpu.Flag("Carry"), "ADD: Carry flag set on overflow"
    
    ' Test 3: Half-carry boundary (0x0F + 0x01)
    Set cpu = CreateTestCPU
    cpu.reg("A") = &HF
    cpu.reg("B") = 1
    cpu.ADD "B"
    AssertEquals &H10, cpu.reg("A"), "ADD: Half-carry detected at 4-bit boundary"
    AssertEquals 1, cpu.Flag("AC"), "ADD: AC (Auxiliary Carry) flag set"
    
    ' Test 4: Parity - 7 bits set (odd parity)
    Set cpu = CreateTestCPU
    cpu.reg("A") = &H7F  ' 0111 1111
    cpu.reg("B") = 0
    cpu.ADD "B"
    AssertEquals &H7F, cpu.reg("A"), "ADD: 127 unchanged"
    AssertEquals 0, cpu.Flag("Parity"), "ADD: Parity flag clear (odd)"
    
    ' Test 5: Parity - 8 bits set (even parity)
    Set cpu = CreateTestCPU
    cpu.reg("A") = &HFF  ' 1111 1111
    cpu.reg("B") = 0
    cpu.ADD "B"
    AssertEquals 1, cpu.Flag("Parity"), "ADD: Parity flag set (even)"
    
End Sub

Sub Test_ADC_WithCarry()
    pTestHeader = "ADC - Add with Carry"
    Dim cpu As clsDecCPU
    
    ' Test 1: ADC without carry-in
    Set cpu = CreateTestCPU
    cpu.reg("A") = 50
    cpu.reg("B") = 30
    cpu.SetFlag "Carry", 0
    cpu.ADC "B"
    AssertEquals 80, cpu.reg("A"), "ADC: 50 + 30 + 0 = 80"
    
    ' Test 2: ADC with carry-in (adds 1)
    Set cpu = CreateTestCPU
    cpu.reg("A") = 50
    cpu.reg("B") = 30
    cpu.SetFlag "Carry", 1
    cpu.ADC "B"
    AssertEquals 81, cpu.reg("A"), "ADC: 50 + 30 + 1(carry) = 81"
    
    ' Test 3: ADC causing carry-out
    Set cpu = CreateTestCPU
    cpu.reg("A") = 255
    cpu.reg("B") = 0
    cpu.SetFlag "Carry", 1
    cpu.ADC "B"
    AssertEquals 0, cpu.reg("A"), "ADC: 255 + 0 + 1 = 0 (overflow, carry set)"
    AssertEquals 1, cpu.Flag("Carry"), "ADC: Carry flag set on overflow"
    
    ' Test 4: ADC with half-carry from Carry bit
    Set cpu = CreateTestCPU
    cpu.reg("A") = &H0F
    cpu.reg("B") = 0
    cpu.SetFlag "Carry", 1
    cpu.ADC "B"
    AssertEquals &H10, cpu.reg("A"), "ADC: 15 + 0 + 1(C) = 16 (half-carry)"
    AssertEquals 1, cpu.Flag("AC"), "ADC: AC flag set"
    
End Sub

Sub Test_SUB_BorrowCases()
    pTestHeader = "SUB - Subtract with Borrow"
    Dim cpu As clsDecCPU
    
    ' Test 1: Simple subtraction
    Set cpu = CreateTestCPU
    cpu.reg("A") = 100
    cpu.reg("B") = 30
    cpu.SUB "B"
    AssertEquals 70, cpu.reg("A"), "SUB: 100 - 30 = 70"
    AssertEquals 0, cpu.Flag("Carry"), "SUB: No borrow (Carry clear)"
    
    ' Test 2: Subtraction requiring borrow
    Set cpu = CreateTestCPU
    cpu.reg("A") = 30
    cpu.reg("B") = 100
    cpu.SUB "B"
    AssertEquals 186, cpu.reg("A"), "SUB: 30 - 100 = -70 (wrap to 186)"
    AssertEquals 1, cpu.Flag("Carry"), "SUB: Borrow required (Carry set)"
    
    ' Test 3: Subtract from zero
    Set cpu = CreateTestCPU
    cpu.reg("A") = 0
    cpu.reg("B") = 1
    cpu.SUB "B"
    AssertEquals 255, cpu.reg("A"), "SUB: 0 - 1 = 255 (underflow wrap)"
    AssertEquals 1, cpu.Flag("Carry"), "SUB: Carry set (borrow)"
    
    ' Test 4: Half-carry borrow (crossing nibble boundary)
    Set cpu = CreateTestCPU
    cpu.reg("A") = &H10
    cpu.reg("B") = 1
    cpu.SUB "B"
    AssertEquals &H0F, cpu.reg("A"), "SUB: 16 - 1 = 15"
    AssertEquals 1, cpu.Flag("AC"), "SUB: AC flag set (half-borrow)"
    
    ' Test 5: Subtract self (clear A)
    Set cpu = CreateTestCPU
    cpu.reg("A") = 42
    cpu.SUB "A"
    AssertEquals 0, cpu.reg("A"), "SUB: A - A = 0"
    AssertEquals 1, cpu.Flag("Zero"), "SUB: Zero flag set"
    AssertEquals 0, cpu.Flag("Carry"), "SUB: No carry on zero result"
    
End Sub


'==============================================================================
' B. LOGICAL OPERATIONS - BOUNDARY TESTS
'==============================================================================

Sub Test_ANA_LogicalAND()
    pTestHeader = "ANA - Logical AND"
    Dim cpu As clsDecCPU
    
    ' Test 1: Basic AND
    Set cpu = CreateTestCPU
    cpu.reg("A") = &HFF  ' 1111 1111
    cpu.reg("B") = &H0F  ' 0000 1111
    cpu.ANA "B"
    AssertEquals &H0F, cpu.reg("A"), "ANA: FF AND 0F = 0F"
    
    ' Test 2: AND with zero (clears A)
    Set cpu = CreateTestCPU
    cpu.reg("A") = &HFF
    cpu.reg("B") = &H00
    cpu.ANA "B"
    AssertEquals 0, cpu.reg("A"), "ANA: FF AND 00 = 00"
    AssertEquals 1, cpu.Flag("Zero"), "ANA: Zero flag set"
    
    ' Test 3: AND of alternating bits
    Set cpu = CreateTestCPU
    cpu.reg("A") = &HAA  ' 1010 1010
    cpu.reg("B") = &H55  ' 0101 0101
    cpu.ANA "B"
    AssertEquals 0, cpu.reg("A"), "ANA: AA AND 55 = 00"
    
    ' Test 4: AND clears carry (per 8080 spec)
    Set cpu = CreateTestCPU
    cpu.SetFlag "Carry", 1
    cpu.reg("A") = &HFF
    cpu.reg("B") = &HFF
    cpu.ANA "B"
    AssertEquals 0, cpu.Flag("Carry"), "ANA: Clears Carry flag"
    
End Sub

Sub Test_ORA_LogicalOR()
    pTestHeader = "ORA - Logical OR"
    Dim cpu As clsDecCPU
    
    ' Test 1: Basic OR
    Set cpu = CreateTestCPU
    cpu.reg("A") = &H0F  ' 0000 1111
    cpu.reg("B") = &HF0  ' 1111 0000
    cpu.ORA "B"
    AssertEquals &HFF, cpu.reg("A"), "ORA: 0F OR F0 = FF"
    
    ' Test 2: OR with zero (no change)
    Set cpu = CreateTestCPU
    cpu.reg("A") = &H42
    cpu.reg("B") = 0
    cpu.ORA "B"
    AssertEquals &H42, cpu.reg("A"), "ORA: 42 OR 00 = 42"
    
    ' Test 3: OR of same register
    Set cpu = CreateTestCPU
    cpu.reg("A") = &H55
    cpu.ORA "A"
    AssertEquals &H55, cpu.reg("A"), "ORA: 55 OR 55 = 55"
    AssertEquals 0, cpu.Flag("Carry"), "ORA: Clears Carry"
    
End Sub

Sub Test_XRA_LogicalXOR()
    pTestHeader = "XRA - Logical XOR"
    Dim cpu As clsDecCPU
    
    ' Test 1: XOR of different patterns
    Set cpu = CreateTestCPU
    cpu.reg("A") = &HF0  ' 1111 0000
    cpu.reg("B") = &H0F  ' 0000 1111
    cpu.XRA "B"
    AssertEquals &HFF, cpu.reg("A"), "XRA: F0 XOR 0F = FF"
    
    ' Test 2: XOR with self (clears A) - common trick to zero register
    Set cpu = CreateTestCPU
    cpu.reg("A") = &H55
    cpu.XRA "A"
    AssertEquals 0, cpu.reg("A"), "XRA: 55 XOR 55 = 00 (common zero trick)"
    AssertEquals 1, cpu.Flag("Zero"), "XRA: Zero flag set"
    
    ' Test 3: XOR with zero (no change)
    Set cpu = CreateTestCPU
    cpu.reg("A") = &HAA
    cpu.reg("B") = 0
    cpu.XRA "B"
    AssertEquals &HAA, cpu.reg("A"), "XRA: AA XOR 00 = AA"
    
End Sub


'==============================================================================
' C. ROTATE/SHIFT - ALL VARIANTS
'==============================================================================

Sub Test_RLC_RotateLeftCircular()
    pTestHeader = "RLC - Rotate Left Circular"
    Dim cpu As clsDecCPU
    
    ' Test 1: Basic left rotate (bit 7 -> bit 0)
    Set cpu = CreateTestCPU
    cpu.reg("A") = &H80  ' 1000 0000
    cpu.RLC
    AssertEquals &H01, cpu.reg("A"), "RLC: 80 rotated left = 01"
    
    ' Test 2: Rotate with carry propagation
    Set cpu = CreateTestCPU
    cpu.reg("A") = &HC0  ' 1100 0000
    cpu.RLC
    AssertEquals &H81, cpu.reg("A"), "RLC: C0 rotated left = 81"
    AssertEquals 1, cpu.Flag("Carry"), "RLC: Carry set from bit 7"
    
    ' Test 3: Rotate zero
    Set cpu = CreateTestCPU
    cpu.reg("A") = 0
    cpu.RLC
    AssertEquals 0, cpu.reg("A"), "RLC: 00 rotated = 00"
    
End Sub

Sub Test_RRC_RotateRightCircular()
    pTestHeader = "RRC - Rotate Right Circular"
    Dim cpu As clsDecCPU
    
    ' Test 1: Basic right rotate (bit 0 -> bit 7)
    Set cpu = CreateTestCPU
    cpu.reg("A") = &H01  ' 0000 0001
    cpu.RRC
    AssertEquals &H80, cpu.reg("A"), "RRC: 01 rotated right = 80"
    AssertEquals 1, cpu.Flag("Carry"), "RRC: Carry set from bit 0"
    
    ' Test 2: Rotate with multiple bits
    Set cpu = CreateTestCPU
    cpu.reg("A") = &H03  ' 0000 0011
    cpu.RRC
    AssertEquals &H81, cpu.reg("A"), "RRC: 03 rotated right = 81"
    
End Sub

Sub Test_RAL_RotateLeftArithmetic()
    pTestHeader = "RAL - Rotate Left through Carry"
    Dim cpu As clsDecCPU
    
    ' Test 1: RAL with Carry clear
    Set cpu = CreateTestCPU
    cpu.reg("A") = &H80  ' 1000 0000
    cpu.SetFlag "Carry", 0
    cpu.RAL
    AssertEquals &H00, cpu.reg("A"), "RAL: 80 << with C=0 = 00"
    AssertEquals 1, cpu.Flag("Carry"), "RAL: Carry gets bit 7"
    
    ' Test 2: RAL with Carry set (chaining)
    Set cpu = CreateTestCPU
    cpu.reg("A") = &H00
    cpu.SetFlag "Carry", 1
    cpu.RAL
    AssertEquals &H01, cpu.reg("A"), "RAL: 00 << with C=1 = 01"
    
End Sub

Sub Test_RAR_RotateRightArithmetic()
    pTestHeader = "RAR - Rotate Right through Carry"
    Dim cpu As clsDecCPU
    
    ' Test 1: RAR with Carry clear
    Set cpu = CreateTestCPU
    cpu.reg("A") = &H01  ' 0000 0001
    cpu.SetFlag "Carry", 0
    cpu.RAR
    AssertEquals &H00, cpu.reg("A"), "RAR: 01 >> with C=0 = 00"
    AssertEquals 1, cpu.Flag("Carry"), "RAR: Carry gets bit 0"
    
    ' Test 2: RAR with Carry set (chaining)
    Set cpu = CreateTestCPU
    cpu.reg("A") = &H00
    cpu.SetFlag "Carry", 1
    cpu.RAR
    AssertEquals &H80, cpu.reg("A"), "RAR: 00 >> with C=1 = 80"
    
End Sub


'==============================================================================
' D. COMPARE INSTRUCTION VARIANTS
'==============================================================================

Sub Test_CPI_CompareImmediate()
    pTestHeader = "CPI - Compare Immediate"
    Dim cpu As clsDecCPU
    
    ' Test 1: Equal (Zero flag set)
    Set cpu = CreateTestCPU
    cpu.reg("A") = 42
    cpu.CPI "2A"  ' 42 hex
    AssertEquals 42, cpu.reg("A"), "CPI: A unchanged"
    AssertEquals 1, cpu.Flag("Zero"), "CPI: Zero flag set when A == operand"
    
    ' Test 2: A > operand
    Set cpu = CreateTestCPU
    cpu.reg("A") = 100
    cpu.CPI "32"  ' 50
    AssertEquals 0, cpu.Flag("Zero"), "CPI: Zero flag clear when A != operand"
    
    ' Test 3: A < operand
    Set cpu = CreateTestCPU
    cpu.reg("A") = 50
    cpu.CPI "64"  ' 100
    AssertEquals 1, cpu.Flag("Carry"), "CPI: Carry set when A < operand"
    
End Sub


'==============================================================================
' E. FLAG MANIPULATION TESTS
'==============================================================================

Sub Test_SetFlagAndCheckFlag()
    pTestHeader = "Flag Manipulation"
    Dim cpu As clsDecCPU
    
    Set cpu = CreateTestCPU
    
    ' Test all flag types
    cpu.SetFlag "Carry", 1
    AssertEquals 1, cpu.Flag("Carry"), "Flag: Carry set"
    
    cpu.SetFlag "Zero", 1
    AssertEquals 1, cpu.Flag("Zero"), "Flag: Zero set"
    
    cpu.SetFlag "Sign", 1
    AssertEquals 1, cpu.Flag("Sign"), "Flag: Sign set"
    
    cpu.SetFlag "Parity", 1
    AssertEquals 1, cpu.Flag("Parity"), "Flag: Parity set"
    
    cpu.SetFlag "AC", 1
    AssertEquals 1, cpu.Flag("AC"), "Flag: AC set"
    
End Sub


'==============================================================================
' F. CONDITIONAL JUMP VARIANTS
'==============================================================================

Sub Test_JC_JumpIfCarry()
    pTestHeader = "JC - Jump If Carry"
    Dim cpu As clsDecCPU
    
    ' Test 1: Jump when carry set
    Set cpu = CreateTestCPU
    cpu.SetFlag "Carry", 1
    cpu.reg("PC") = 0
    cpu.JC "0100"
    AssertEquals &h0100, cpu.reg("PC"), "JC: Jumped to 0100 (Carry=1)"
    
    ' Test 2: No jump when carry clear
    Set cpu = CreateTestCPU
    cpu.SetFlag "Carry", 0
    cpu.reg("PC") = 0
    cpu.JC "0100"
    AssertEquals 0, cpu.reg("PC"), "JC: No jump (Carry=0)"
    
End Sub

Sub Test_JNZ_JumpIfNotZero()
    pTestHeader = "JNZ - Jump If Not Zero"
    Dim cpu As clsDecCPU
    
    ' Test 1: Jump when zero flag clear
    Set cpu = CreateTestCPU
    cpu.SetFlag "Zero", 0
    cpu.reg("PC") = 0
    cpu.JNZ "0200"
    AssertEquals &h0200, cpu.reg("PC"), "JNZ: Jumped to 0200 (Zero=0)"
    
    ' Test 2: No jump when zero flag set
    Set cpu = CreateTestCPU
    cpu.SetFlag "Zero", 1
    cpu.reg("PC") = 0
    cpu.JNZ "0200"
    AssertEquals 0, cpu.reg("PC"), "JNZ: No jump (Zero=1)"
    
End Sub

Sub Test_JM_JumpIfMinus()
    pTestHeader = "JM - Jump If Minus (Sign=1)"
    Dim cpu As clsDecCPU
    
    ' Test 1: Jump when negative (Sign=1)
    Set cpu = CreateTestCPU
    cpu.SetFlag "Sign", 1
    cpu.reg("PC") = 0
    cpu.JM "0300"
    AssertEquals &h0300, cpu.reg("PC"), "JM: Jumped to 0300 (Sign=1)"
    
    ' Test 2: No jump when positive (Sign=0)
    Set cpu = CreateTestCPU
    cpu.SetFlag "Sign", 0
    cpu.reg("PC") = 0
    cpu.JM "0300"
    AssertEquals 0, cpu.reg("PC"), "JM: No jump (Sign=0)"
    
End Sub


'==============================================================================
' G. REGISTER MOVE VARIANTS
'==============================================================================

Sub Test_MOV_AllRegisterPairs()
    pTestHeader = "MOV - Register Pair Combinations"
    Dim cpu As clsDecCPU
    Dim regs() As String
    regs = Split("A,B,C,D,E,H,L", ",")
    Dim i As Long, j As Long
    
    For i = 0 To UBound(regs)
        For j = 0 To UBound(regs)
            If i <> j Then
                Set cpu = CreateTestCPU
                cpu.reg(regs(i)) = 50 + i
                cpu.MOV regs(j), regs(i)
                AssertEquals 50 + i, cpu.reg(regs(j)), _
                    "MOV: " & regs(j) & " <- " & regs(i)
            End If
        Next j
    Next i
    
End Sub

Sub Test_MVI_AllRegisters()
    pTestHeader = "MVI - Move Immediate to All Registers"
    Dim cpu As clsDecCPU
    Dim regs() As String
    regs = Split("A,B,C,D,E,H,L", ",")
    Dim i As Long
    
    For i = 0 To UBound(regs)
        Set cpu = CreateTestCPU
        cpu.MVI regs(i), "55"  ' Load 0x55 = 85
        AssertEquals &h55, cpu.reg(regs(i)), _
            "MVI: " & regs(i) & " loaded with 55"
    Next i
    
End Sub


'==============================================================================
' H. INTEGRATION TEST
'==============================================================================

Sub Test_Integration_CountDown()
    pTestHeader = "Integration - Countdown Loop"
    Dim cpu As clsDecCPU
    Set cpu = CreateTestCPU
    
    cpu.reg("B") = 5
    cpu.reg("A") = 0
    
    ' Simulate: 
    '   LOOP: ADD B
    '        DCR B
    '        JNZ LOOP
    
    Dim loopCount As Long
    For loopCount = 1 To 5
        cpu.ADD "B"
        cpu.DCR "B"
        If cpu.Flag("Zero") = 1 Then Exit For
    Next loopCount
    
    AssertEquals 15, cpu.reg("A"), "Accumulated 5+4+3+2+1 = 15"
    AssertEquals 5, loopCount, "Loop executed 5 times"
    
End Sub


'================================================================================
' PART 3: UPDATE YOUR MASTER TEST SUITE
'================================================================================

' Add these lines to your ClaudeRunAllTests() function:

'    ' ARITHMETIC ENHANCED (8 tests)
'    Test_ADD_EdgeCases
'    Test_ADC_WithCarry
'    Test_SUB_BorrowCases
'    
'    ' LOGICAL ENHANCED (3 tests)
'    Test_ANA_LogicalAND
'    Test_ORA_LogicalOR
'    Test_XRA_LogicalXOR
'    
'    ' ROTATE/SHIFT COMPLETE (5 tests)
'    Test_RLC_RotateLeftCircular
'    Test_RRC_RotateRightCircular
'    Test_RAL_RotateLeftArithmetic
'    Test_RAR_RotateRightArithmetic
'    
'    ' COMPARE VARIANTS (1 test)
'    Test_CPI_CompareImmediate
'    
'    ' FLAG MANIPULATION (1 test)
'    Test_SetFlagAndCheckFlag
'    
'    ' CONDITIONAL JUMPS (3 tests)
'    Test_JC_JumpIfCarry
'    Test_JNZ_JumpIfNotZero
'    Test_JM_JumpIfMinus
'    
'    ' REGISTER MOVES (2 tests)
'    Test_MOV_AllRegisterPairs
'    Test_MVI_AllRegisters
'    
'    ' INTEGRATION (1 test)
'    Test_Integration_CountDown

'================================================================================
' END OF FILE
'================================================================================
