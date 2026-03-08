# 8080/Z80 Opcode Calculation Guide

## Understanding Opcode Encoding

The Intel 8080 uses an elegant binary encoding scheme where opcodes are carefully structured so that instruction type and operands can be decoded directly from the bit pattern. This guide explains how to calculate opcodes from their bit patterns.

---

## Part 1: Binary & Hexadecimal Basics

### Quick Reference

```
Hexadecimal Digit | Binary
------------------|----------
0                  | 0000
1                  | 0001
2                  | 0010
3                  | 0011
4                  | 0100
5                  | 0101
6                  | 0110
7                  | 0111
8                  | 1000
9                  | 1001
A                  | 1010
B                  | 1011
C                  | 1100
D                  | 1101
E                  | 1110
F                  | 1111
```

### Register Encoding (00-7)

The 8080 encodes registers as 3-bit values:

```
Value (Binary/Decimal) | Register | 16-bit Pair
-----------------------|----------|----------
000 / 0                | B        | BC
001 / 1                | C        |
010 / 2                | D        | DE
011 / 3                | E        |
100 / 4                | H        | HL
101 / 5                | L        |
110 / 6                | M        | (HL) - memory
111 / 7                | A        |
```

**Key Point:** Registers come in pairs for 16-bit operations:
- **BC**: B=0, C=1
- **DE**: D=2, E=3
- **HL**: H=4, L=5
- **PSW/AF**: A=7, Flags

---

## Part 2: MOV Instruction - The Classic Example

### MOV Pattern: 01mmmss (or 01 ddd sss in some notations)

The MOV instruction has the pattern:
```
Bit 7 6 5 4 3 2 1 0
    [0][1][m][m][m][s][s][s]
     └──────────┘ │  │  │  │
     Fixed bits  │  │  │  └─ Source register (3 bits)
                 │  └───────────── Destination register (3 bits)
                 └─────────────── Always 01 (binary)
```

**Breaking it down:**
- Bits 7-6: Always `01` (identifies this as a MOV instruction)
- Bits 5-3: Destination register (ddd)
- Bits 2-0: Source register (sss)

### Example 1: MOV A, B (Copy B into A)

**Step 1: Identify the registers**
- Destination: A = 111 (binary)
- Source: B = 000 (binary)

**Step 2: Construct the opcode**
```
Bit Pattern:     0 1 1 1 1 0 0 0
                 └ └ └ └ ┬ ┬ ┬ ┬
                 0 1 A B   (binary: 111 000)
```

**Step 3: Convert binary to hex**
```
Binary: 0111 1000
        │││└ ││└┘
        ││└──┘└──── 1000 = 8 (hex)
        └└───────── 0111 = 7 (hex)

Result: 0x78
```

**Verification:**
- Binary: 01111000
- Hex: 0x78
- **This is the actual 8080 opcode for MOV A,B!**

### Example 2: MOV B, C (Copy C into B)

**Step 1: Identify registers**
- Destination: B = 000
- Source: C = 001

**Step 2: Construct opcode**
```
Bit Pattern:     0 1 0 0 0 0 0 1
                 └ └ └ ─ ─ ─ ─ ┘
                 0 1 B       C (binary: 000 001)
```

**Step 3: Convert to hex**
```
Binary: 0100 0001
Hex:    0x41
```

**Answer: 0x41**

### Example 3: MOV H, M (Copy memory at (HL) into H)

**Step 1: Identify registers**
- Destination: H = 100
- Source: M = 110 (memory at HL)

**Step 2: Construct opcode**
```
Bit Pattern:     0 1 1 0 0 1 1 0
                 └ └ └ ─ ─ ─ ─ ┘
                 0 1 H       M (binary: 100 110)
```

**Step 3: Convert to hex**
```
Binary: 0110 0110
Hex:    0x66
```

**Answer: 0x66**

### Example 4: MOV M, A (Store A into memory at (HL))

**Step 1: Identify registers**
- Destination: M = 110 (memory at HL)
- Source: A = 111

**Step 2: Construct opcode**
```
Bit Pattern:     0 1 1 1 0 1 1 1
                 └ └ └ ─ ─ ─ ─ ┘
                 0 1 M       A (binary: 110 111)
```

**Step 3: Convert to hex**
```
Binary: 0111 0111
Hex:    0x77
```

**Answer: 0x77**

### Complete MOV Opcode Table

Using the formula: **Opcode = 0x40 + (destination × 8) + source**

| Source→ | B (0) | C (1) | D (2) | E (3) | H (4) | L (5) | M (6) | A (7) |
|---------|-------|-------|-------|-------|-------|-------|-------|-------|
| **B (0)** | 0x40 | 0x41 | 0x42 | 0x43 | 0x44 | 0x45 | 0x46 | 0x47 |
| **C (1)** | 0x48 | 0x49 | 0x4A | 0x4B | 0x4C | 0x4D | 0x4E | 0x4F |
| **D (2)** | 0x50 | 0x51 | 0x52 | 0x53 | 0x54 | 0x55 | 0x56 | 0x57 |
| **E (3)** | 0x58 | 0x59 | 0x5A | 0x5B | 0x5C | 0x5D | 0x5E | 0x5F |
| **H (4)** | 0x60 | 0x61 | 0x62 | 0x63 | 0x64 | 0x65 | 0x66 | 0x67 |
| **L (5)** | 0x68 | 0x69 | 0x6A | 0x6B | 0x6C | 0x6D | 0x6E | 0x6F |
| **M (6)** | 0x70 | 0x71 | 0x72 | 0x73 | 0x74 | 0x75 | 0x76 | 0x77 |
| **A (7)** | 0x78 | 0x79 | 0x7A | 0x7B | 0x7C | 0x7D | 0x7E | 0x7F |

**Pattern:** All MOV opcodes fall between 0x40 and 0x7F (64-127 in decimal)

---

## Part 3: Other Instructions Using Similar Patterns

### INR (Increment Register): Pattern 00ddd100

**Pattern:**
```
Bit 7 6 5 4 3 2 1 0
    [0][0][d][d][d][1][0][0]
```

- Bits 7-6: Always 00
- Bits 5-3: Register to increment (ddd)
- Bits 2-0: Always 100 (binary 4)

**Example: INR A (Increment A)**
```
Binary: 0011 1100
Hex:    0x3C
```

**Example: INR B (Increment B)**
```
Binary: 0000 0100
Hex:    0x04
```

**Formula:** Opcode = 0x04 + (register × 8)

| Register | B | C | D | E | H | L | M | A |
|----------|---|---|---|---|---|---|---|---|
| Opcode | 0x04 | 0x0C | 0x14 | 0x1C | 0x24 | 0x2C | 0x34 | 0x3C |

---

### DCR (Decrement Register): Pattern 00ddd101

**Pattern:**
```
Bit 7 6 5 4 3 2 1 0
    [0][0][d][d][d][1][0][1]
```

- Bits 7-6: Always 00
- Bits 5-3: Register to decrement (ddd)
- Bits 2-0: Always 101 (binary 5)

**Example: DCR C (Decrement C)**
```
Binary: 0000 1101
Hex:    0x0D
```

**Formula:** Opcode = 0x05 + (register × 8)

| Register | B | C | D | E | H | L | M | A |
|----------|---|---|---|---|---|---|---|---|
| Opcode | 0x05 | 0x0D | 0x15 | 0x1D | 0x25 | 0x2D | 0x35 | 0x3D |

---

### ADD (Add Register to A): Pattern 10000sss

**Pattern:**
```
Bit 7 6 5 4 3 2 1 0
    [1][0][0][0][0][s][s][s]
```

- Bits 7-6: Always 10
- Bits 5-3: Always 000
- Bits 2-0: Source register (sss)

**Example: ADD B (Add B to A)**
```
Binary: 1000 0000
Hex:    0x80
```

**Example: ADD M (Add memory to A)**
```
Binary: 1000 0110
Hex:    0x86
```

**Formula:** Opcode = 0x80 + source_register

| Register | B | C | D | E | H | L | M | A |
|----------|---|---|---|---|---|---|---|---|
| Opcode | 0x80 | 0x81 | 0x82 | 0x83 | 0x84 | 0x85 | 0x86 | 0x87 |

---

### ADI (Add Immediate to A): Pattern 11000110

**Pattern:**
```
Bit 7 6 5 4 3 2 1 0
    [1][1][0][0][0][1][1][0]
```

This is a **fixed opcode** with no register bits:
- **Opcode: 0xC6**
- **Next byte: immediate value** (the number to add)

**Example: ADI 0x42 (Add 0x42 to A)**
```
Machine code: C6 42
Assembly:    ADI 0x42
```

---

## Part 4: Multi-Byte Instructions

### LXI (Load Extended Immediate - 16-bit Pair): Pattern 00rr0001

**Pattern:**
```
Bit 7 6 5 4 3 2 1 0
    [0][0][r][r][0][0][0][1]
```

- Bits 7-6: Always 00
- Bits 5-4: Register pair (rr) = 00=BC, 01=DE, 10=HL, 11=SP
- Bits 3-0: Always 0001

**Encoding for register pairs:**
```
Value | Pair | Operation
------|------|------------------------------
00    | BC   | Load 16-bit value into B,C
01    | DE   | Load 16-bit value into D,E
10    | HL   | Load 16-bit value into H,L
11    | SP   | Load 16-bit value into SP
```

**Example 1: LXI H, 0x1234 (Load 0x1234 into HL)**

**Step 1: Identify the pair**
- rr = 10 (binary) for HL

**Step 2: Construct opcode byte**
```
Binary: 0010 0001
Hex:    0x21
```

**Step 3: Add the 16-bit immediate**
- The value 0x1234 is stored as: LSB first (0x34), then MSB (0x12)

**Machine code:** `21 34 12` (3 bytes)

**Example 2: LXI BC, 0x5678 (Load 0x5678 into BC)**

**Step 1: Identify the pair**
- rr = 00 (binary) for BC

**Step 2: Construct opcode byte**
```
Binary: 0000 0001
Hex:    0x01
```

**Step 3: Add the immediate**
- Value 0x5678: LSB (0x78), MSB (0x56)

**Machine code:** `01 78 56` (3 bytes)

**Complete LXI Table:**

| Pair | Opcode | Example | Machine Code |
|------|--------|---------|--------------|
| BC | 0x01 | LXI BC, 0x2000 | 01 00 20 |
| DE | 0x11 | LXI DE, 0x3000 | 11 00 30 |
| HL | 0x21 | LXI HL, 0x1000 | 21 00 10 |
| SP | 0x31 | LXI SP, 0x00FF | 31 FF 00 |

---

## Part 5: Conditional Jump Instructions

### Conditional Jump Pattern: 11ccc010

**Pattern:**
```
Bit 7 6 5 4 3 2 1 0
    [1][1][c][c][c][0][1][0]
```

- Bits 7-6: Always 11
- Bits 5-3: Condition code (ccc)
- Bits 2-0: Always 010

**Condition Codes:**

```
Code (Binary) | Mnemonic | Meaning
--------------|----------|---------------------------
000           | JZ       | Jump if Zero = 1
001           | JNZ      | Jump if Zero = 0
010           | JC       | Jump if Carry = 1
011           | JNC      | Jump if Carry = 0
100           | JPE      | Jump if Parity = 1 (Even)
101           | JPO      | Jump if Parity = 0 (Odd)
110           | JP       | Jump if Sign = 0 (Positive)
111           | JM       | Jump if Sign = 1 (Minus)
```

**Example 1: JZ 0x1000 (Jump if Zero)**

**Step 1: Condition code**
- ccc = 000 (binary)

**Step 2: Construct opcode byte**
```
Binary: 1100 1010
Hex:    0xCA
```

**Step 3: Add the 16-bit address**
- 0x1000 = LSB (0x00), MSB (0x10)

**Machine code:** `CA 00 10` (3 bytes)

**Example 2: JC 0x2000 (Jump if Carry)**

**Step 1: Condition code**
- ccc = 010 (binary)

**Step 2: Construct opcode byte**
```
Binary: 1101 1010
Hex:    0xDA
```

**Step 3: Add address**
- 0x2000 = LSB (0x00), MSB (0x20)

**Machine code:** `DA 00 20` (3 bytes)

**Complete Conditional Jump Table:**

| Instruction | Opcode | Meaning |
|-------------|--------|---------|
| JZ | 0xCA | Jump if Zero flag = 1 |
| JNZ | 0xC2 | Jump if Zero flag = 0 |
| JC | 0xDA | Jump if Carry flag = 1 |
| JNC | 0xD2 | Jump if Carry flag = 0 |
| JPE | 0xEA | Jump if Parity flag = 1 (even) |
| JPO | 0xE2 | Jump if Parity flag = 0 (odd) |
| JP | 0xF2 | Jump if Sign flag = 0 (positive) |
| JM | 0xFA | Jump if Sign flag = 1 (minus/negative) |

**Pattern:** All conditional jumps have opcodes in the format `1100xxx0` and `11x1x010`

---

## Part 6: PUSH and POP Instructions

### PUSH Pattern: 11rr0101

**Pattern:**
```
Bit 7 6 5 4 3 2 1 0
    [1][1][r][r][0][1][0][1]
```

- Bits 7-6: Always 11
- Bits 5-4: Register pair (rr)
- Bits 3-0: Always 0101

**Register Pair Encoding for PUSH/POP:**

```
Code (rr) | Pair | Meaning
----------|------|------------------
00        | BC   | Push/Pop BC
01        | DE   | Push/Pop DE
10        | HL   | Push/Pop HL
11        | PSW  | Push/Pop A and Flags
```

**Example 1: PUSH B (Push BC onto stack)**

**Step 1: Register pair**
- rr = 00 (binary)

**Step 2: Construct opcode**
```
Binary: 1100 0101
Hex:    0xC5
```

**Machine code:** `C5` (1 byte)

**Example 2: PUSH PSW (Push A and Flags onto stack)**

**Step 1: Register pair**
- rr = 11 (binary) for PSW

**Step 2: Construct opcode**
```
Binary: 1111 0101
Hex:    0xF5
```

**Machine code:** `F5` (1 byte)

### POP Pattern: 11rr0001

**Pattern:**
```
Bit 7 6 5 4 3 2 1 0
    [1][1][r][r][0][0][0][1]
```

- Bits 7-6: Always 11
- Bits 5-4: Register pair (rr)
- Bits 3-0: Always 0001

**Same register pair encoding as PUSH**

**Example: POP H (Pop HL from stack)**

**Step 1: Register pair**
- rr = 10 (binary)

**Step 2: Construct opcode**
```
Binary: 1110 0001
Hex:    0xE1
```

**Machine code:** `E1` (1 byte)

**Complete PUSH/POP Table:**

| Instruction | Opcode | Pair | Effect |
|-------------|--------|------|--------|
| PUSH B | 0xC5 | BC | (SP-2:SP) ← BC, SP ← SP-2 |
| PUSH D | 0xD5 | DE | (SP-2:SP) ← DE, SP ← SP-2 |
| PUSH H | 0xE5 | HL | (SP-2:SP) ← HL, SP ← SP-2 |
| PUSH PSW | 0xF5 | AF | (SP-2:SP) ← AF, SP ← SP-2 |
| POP B | 0xC1 | BC | BC ← (SP:SP+2), SP ← SP+2 |
| POP D | 0xD1 | DE | DE ← (SP:SP+2), SP ← SP+2 |
| POP H | 0xE1 | HL | HL ← (SP:SP+2), SP ← SP+2 |
| POP PSW | 0xF1 | AF | AF ← (SP:SP+2), SP ← SP+2 |

---

## Part 7: Fixed Opcode Instructions

Some instructions have no variable operands and are simply fixed hex values:

| Instruction | Opcode | Binary | Meaning |
|-------------|--------|--------|---------|
| NOP | 0x00 | 0000 0000 | No operation |
| HLT | 0x76 | 0111 0110 | Halt execution |
| JMP | 0xC3 | 1100 0011 | Unconditional jump (next 2 bytes = address) |
| CALL | 0xCD | 1100 1101 | Call subroutine (next 2 bytes = address) |
| RET | 0xC9 | 1100 1001 | Return from subroutine |
| EXX | 0xD9 | 1101 1001 | Exchange shadow registers (Z80 only) |
| DAA | 0x27 | 0010 0111 | Decimal Adjust Accumulator |

---

## Part 8: VBA Implementation Pattern

Here's how to implement opcode calculation in your VBA code:

### General Pattern - Calculate Opcode from Components

```vba
' Example: Calculate MOV opcode from destination and source
Public Function CalculateMOVOpcode(ByVal destination As Long, ByVal source As Long) As Long
    ' Pattern: 01mmmss
    ' Bits 7-6: 01 (binary) = 0x40 in decimal
    ' Bits 5-3: destination (0-7)
    ' Bits 2-0: source (0-7)
    
    CalculateMOVOpcode = &H40 + (destination * 8) + source
    
    ' Example:
    ' MOV A, B: destination=7, source=0
    ' Result: 0x40 + (7*8) + 0 = 0x40 + 56 = 0x78
End Function

' Test the function
Sub TestOpcodeCalculation()
    Dim opcode As Long
    
    ' MOV A, B
    opcode = CalculateMOVOpcode(7, 0)
    MsgBox "MOV A, B = 0x" & hex$(opcode)  ' Should print 0x78
    
    ' MOV B, C
    opcode = CalculateMOVOpcode(0, 1)
    MsgBox "MOV B, C = 0x" & hex$(opcode)  ' Should print 0x41
    
    ' MOV H, M
    opcode = CalculateMOVOpcode(4, 6)
    MsgBox "MOV H, M = 0x" & hex$(opcode)  ' Should print 0x66
End Sub
```

### Conditional Jump Opcode Calculation

```vba
' Calculate conditional jump opcode from condition code
Public Function CalculateConditionalJumpOpcode(ByVal condition As Long) As Long
    ' Pattern: 11ccc010
    ' Bits 7-6: 11 (binary) = 0xC0
    ' Bits 5-3: condition code (0-7)
    ' Bits 2-0: 010 (binary) = 0x02
    
    ' This becomes: 0xC0 + (condition * 8) + 0x02
    ' = 0xC2 + (condition * 8)
    
    CalculateConditionalJumpOpcode = &HC2 + (condition * 8)
    
    ' Example:
    ' JZ (condition=0): 0xC2 + 0 = 0xC2
    ' JC (condition=2): 0xC2 + 16 = 0xD2
    ' JM (condition=7): 0xC2 + 56 = 0xFA
End Function

Sub TestConditionalJumps()
    Dim opcodes As Object
    Set opcodes = CreateObject("Scripting.Dictionary")
    
    ' Build conditional jump table
    Dim conditions() As Variant
    conditions = Array("JZ", "JNZ", "JC", "JNC", "JPE", "JPO", "JP", "JM")
    
    Dim i As Long
    For i = 0 To 7
        Dim opcode As Long
        opcode = CalculateConditionalJumpOpcode(i)
        opcodes(conditions(i)) = opcode
        Debug.Print conditions(i) & " = 0x" & hex$(opcode)
    Next i
End Sub
```

### INR/DCR Opcode Calculation

```vba
Public Function CalculateINROpcode(ByVal regNum As Long) As Long
    ' Pattern: 00ddd100
    ' Formula: 0x04 + (register * 8)
    CalculateINROpcode = &H4 + (regNum * 8)
End Function

Public Function CalculateDCROpcode(ByVal regNum As Long) As Long
    ' Pattern: 00ddd101
    ' Formula: 0x05 + (register * 8)
    CalculateDCROpcode = &H5 + (regNum * 8)
End Function

Sub TestIncrementDecrement()
    Dim i As Long
    Dim regs() As String
    regs = Split("B,C,D,E,H,L,M,A", ",")
    
    For i = 0 To 7
        Debug.Print "INR " & regs(i) & " = 0x" & hex$(CalculateINROpcode(i))
        Debug.Print "DCR " & regs(i) & " = 0x" & hex$(CalculateDCROpcode(i))
    Next i
End Sub
```

### ADD Opcode Calculation

```vba
Public Function CalculateADDOpcode(ByVal sourceReg As Long) As Long
    ' Pattern: 10000sss
    ' Formula: 0x80 + source
    CalculateADDOpcode = &H80 + sourceReg
End Function

Sub TestADD()
    Dim regs() As String
    regs = Split("B,C,D,E,H,L,M,A", ",")
    
    Dim i As Long
    For i = 0 To 7
        Debug.Print "ADD " & regs(i) & " = 0x" & hex$(CalculateADDOpcode(i))
    Next i
End Sub
```

---

## Part 9: Reverse Engineering - Opcode to Instruction

Sometimes you need to decode an opcode back to the instruction:

```vba
Public Function DecodeOpcode(ByVal opcode As Long) As String
    Dim result As String
    Dim bits7_6 As Long, bits5_3 As Long, bits2_0 As Long
    Dim regs() As String
    regs = Split("B,C,D,E,H,L,M,A", ",")
    
    ' Extract bit ranges
    bits7_6 = (opcode \ 64) And 3         ' Bits 7-6
    bits5_3 = (opcode \ 8) And 7          ' Bits 5-3
    bits2_0 = opcode And 7                 ' Bits 2-0
    
    ' Decode based on bit patterns
    Select Case bits7_6
        Case 0 ' 00xxxxxx
            If bits2_0 = 4 Then
                result = "INR " & regs(bits5_3)
            ElseIf bits2_0 = 5 Then
                result = "DCR " & regs(bits5_3)
            End If
            
        Case 1 ' 01xxxxxx
            ' MOV instruction
            result = "MOV " & regs(bits5_3) & ", " & regs(bits2_0)
            
        Case 2 ' 10xxxxxx
            If bits5_3 = 0 Then
                result = "ADD " & regs(bits2_0)
            ElseIf bits5_3 = 2 Then
                result = "ANA " & regs(bits2_0)
            ElseIf bits5_3 = 4 Then
                result = "ANA " & regs(bits2_0)  ' OR
            End If
            
        Case 3 ' 11xxxxxx
            If bits5_3 = 0 And bits2_0 = 1 Then
                result = "LXI (various)"
            ElseIf bits5_3 = 0 And bits2_0 = 5 Then
                result = "PUSH (various)"
            End If
    End Select
    
    If result = "" Then
        result = "Unknown opcode: 0x" & hex$(opcode)
    End If
    
    DecodeOpcode = result
End Function

Sub TestDecode()
    Debug.Print DecodeOpcode(&H78)  ' Should print MOV A, B
    Debug.Print DecodeOpcode(&H41)  ' Should print MOV B, C
    Debug.Print DecodeOpcode(&H3C)  ' Should print INR A
    Debug.Print DecodeOpcode(&H0D)  ' Should print DCR C
End Sub
```

---

## Part 10: Quick Reference Formulas

```
Instruction | Pattern        | Formula
------------|----------------|----------------------------------------------
MOV         | 01mmmss        | 0x40 + (destination × 8) + source
INR         | 00ddd100       | 0x04 + (register × 8)
DCR         | 00ddd101       | 0x05 + (register × 8)
ADD         | 10000sss       | 0x80 + source
ADI         | 11000110       | 0xC6 (fixed)
ANA         | 10100sss       | 0xA0 + source
ORA         | 10110sss       | 0xB0 + source
XRA         | 10101sss       | 0xA8 + source
JZ/JC/etc   | 11ccc010       | 0xC2 + (condition × 8)
JMP         | 11000011       | 0xC3 (fixed)
CALL        | 11001101       | 0xCD (fixed)
RET         | 11001001       | 0xC9 (fixed)
PUSH        | 11rr0101       | 0xC5 + (pair × 8)
POP         | 11rr0001       | 0xC1 + (pair × 8)
LXI         | 00rr0001       | 0x01 + (pair × 16)
HLT         | 01110110       | 0x76 (fixed)
NOP         | 00000000       | 0x00 (fixed)
```

---

## Summary

The key insight is that the 8080 instruction set is **systematically encoded**:

1. **Fixed bits** identify the instruction type
2. **Variable bits** encode operands (registers, register pairs, etc.)
3. **Operands are encoded as small integers** (0-7 for 8-bit registers, 0-3 for pairs)

This systematic approach allows:
- Easy calculation of opcodes from mnemonics
- Easy decoding of opcodes back to mnemonics
- Compact representation (8-bit or 16-bit instructions)
- Hardware simplicity (minimal decoding logic needed)

By understanding these bit patterns, you can confidently calculate any 8080 opcode and verify your emulator's instruction implementations!

---

**Document Version:** 1.0  
**Last Updated:** March 7, 2026  
**For Use With:** Z80_Model_Current_2026-03-06_2239.xlsm
