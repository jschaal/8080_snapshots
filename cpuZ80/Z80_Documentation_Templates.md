# Z80/8080 Emulator - Documentation Templates & Instruction Reference

This document provides reusable templates for documenting your VBA code and reference material for 8080/Z80 instruction implementation.

---

## Part 1: Instruction Documentation Template

Use this template when adding comments to instruction implementations. Copy and customize:

```vba
' ==============================================================================
' Instruction: [MNEMONIC] [operands]
' 
' Hardware: 8080 / Z80 (Identify which CPU)
' Opcode:   [hex value(s)] (e.g., 80h + register for ADD r)
' Cycles:   [8080 cycles] / [Z80 cycles] (e.g., 4 cycles)
' Bytes:    [1-3] (instruction length)
'
' Syntax:   [MNEMONIC] [operand1], [operand2]
' Example:  ADD B  (Add B register to A)
'
' Operation:
'   [Describe what the CPU does in plain English]
'   Example:   A ← A + B
'   Example:   Flags updated per ALU rules
'
' Flags Affected:
'   CY    = [1 if overflow, 0 otherwise]
'   P     = [1 if even parity, 0 if odd]
'   AC    = [1 if bit 3 overflow, 0 otherwise]
'   Z     = [1 if result is 0, 0 otherwise]
'   S     = [1 if result bit 7 is set, 0 otherwise]
'   N     = [Z80 only: 1 after SUB/SBC, 0 after ADD/ADC] (used for DAA)
'   H     = [Half-carry from bit 3]
'
' Z80 Differences from 8080:
'   [List any behavior differences, if applicable]
'   Example: "Z80 resets N flag; 8080 does not have N flag"
'
' Timing Notes:
'   - Memory access adds cycles
'   - Jumps are [Y+Z] cycles if branch taken, [Y] if not taken
'
' Related Instructions:
'   - [List similar or complementary instructions]
'   - Example: ADI (add immediate), ADC (add with carry)
'
' Error Codes (base [XXX]):
'   [XXX]: [error condition 0] (typically: missing operand)
'   [XXX+1]: [error condition 1] (typically: invalid register/operand)
'   [XXX+2]: [error condition 2] (if applicable)
'   See "Error Code Reference" section below for full details
'
' Example Code:
'   MVI A, 0x50       ; A = 0x50
'   ADD B             ; Add B (if B=0x60, result=0xB0, CY=0)
'   JC CARRY_SET      ; Jump if carry (won't jump)
'
' Common Errors:
'   - Forgetting that [operand] is [type], not [other type]
'   - Flag [X] is NOT affected by this instruction
'   - [Other common mistakes]
'
' ==============================================================================
Public Function [INSTRUCTION_NAME](ByVal op1 As String, Optional ByVal op2 As String = "") As Long
    Dim errorBase As Long: errorBase = [XXX]
    Dim result As Long
    result = 0
    
    ' Validate operands
    ' If op1 = "" Then SetError errorBase, "...": [INSTRUCTION_NAME] = errorBase: Exit Function
    
    ' [Description of implementation logic]
    
    ' Update affected flags
    '   (Detail how each flag is computed)
    
    ' Return status code (0 = success, non-zero = error)
    [INSTRUCTION_NAME] = result
End Function
```

---

## Error Code Pattern Reference

Each instruction uses an `errorBase` constant to generate error codes. Return codes are reported in **Trace column 19 (Err)**.

### Standard Error Offset Pattern

| Offset | Pattern | Typical Usage |
|--------|---------|---------------|
| Base+0 | Always first | Missing operand(s) |
| Base+1 | Always second | Invalid register/operand type |
| Base+2 | Optional | Invalid value range / unresolved label |
| Base+3 | Optional | Out of bounds / memory access error |
| Base+4+ | Instruction-specific | Custom validation errors |

### Return Values

- **0** = Success ✓
- **Negative** = Special signal (e.g., -100 for HLT)
- **Positive** = Error; return errorBase + offset

---

## Part 2: 8080 Instruction Reference

### Group 1: Data Transfer Instructions

#### MOV (Move Data Between Registers or Memory)

```vba
' ==============================================================================
' MOV destination, source
' Opcode:   01mmmss (m=destination, s=source; values 0-7 for B,C,D,E,H,L,M,A)
' 8080:     1 byte, 4-10 cycles (7 if M involved)
' Z80:      1 byte, 4-7 cycles
' 
' Operation:   destination ← source
' Flags:       None affected
' 
' Examples:
'   MOV A,B     - Copy B to A (opcode: 78h)
'   MOV H,M     - Copy memory at (HL) to H (opcode: 64h)
'   MOV M,C     - Copy C to memory at (HL) (opcode: 71h)
'
' Error Codes (base 350):
'   350: Missing operands (op1 or op2 is empty)
'   351: Invalid register(s) (source or destination register not recognized)
' ==============================================================================
```

#### MVI (Move Immediate Data to Register or Memory)

```vba
' ==============================================================================
' MVI destination, data
' Opcode:   06dd 10dd (dd = 0-7 for B,C,D,E,H,L,M,A; second byte is immediate)
' 8080:     2 bytes, 7-10 cycles
' Z80:      2 bytes, 7-10 cycles
'
' Operation:   destination ← immediate byte
' Flags:       None affected
'
' Examples:
'   MVI A, 42h      - Load 0x42 into A (opcode: 3Eh 42h)
'   MVI H, 0x10     - Load 0x10 into H (opcode: 26h 10h)
'
' Error Codes (base 360):
'   360: Missing operands (op1 or op2 is empty)
'   361: Invalid hex value (op2 is not valid hex)
'   362: Invalid register (op1 register name not recognized)
' ==============================================================================
```

#### LXI (Load Extended Immediate - 16-bit Pair)

```vba
' ==============================================================================
' LXI pair, address
' Opcode:   00rr0001 (rr = 0-3 for BC,DE,HL,SP; next 2 bytes are 16-bit value)
' 8080:     3 bytes, 10 cycles
' Z80:      3 bytes, 10 cycles
'
' Operation:   pair ← 16-bit address (LSB first, then MSB)
' Flags:       None affected
'
' Examples:
'   LXI H, 1000h    - Load HL = 0x1000 (opcode: 21h 00h 10h)
'   LXI BC, 2000h   - Load BC = 0x2000 (opcode: 01h 00h 20h)
'
' Notes:
'   - SP can be loaded with LXI SP, address
'   - Operands are always in little-endian order (LSB, MSB)
'
' Error Codes (base 340):
'   340: Missing operands (op1 or op2 is empty)
'   341: Invalid register pair (must be B, D, H, or SP)
'   342: Invalid address (not valid hex or label doesn't resolve)
' ==============================================================================
```

### Group 2: Arithmetic Instructions

#### ADD (Add Register to A)

```vba
' ==============================================================================
' ADD source
' Opcode:   10000ss (ss = 0-7 for B,C,D,E,H,L,M,A)
' 8080:     1 byte, 4 cycles (7 if M)
' Z80:      1 byte, 4 cycles (7 if M)
'
' Operation:   A ← A + source
' Flags Affected:
'   CY = 1 if sum > 255 (overflow)
'   P  = 1 if result has even parity
'   AC = 1 if carry from bit 3
'   Z  = 1 if result is 0
'   S  = 1 if bit 7 is set (result is "negative")
'   N  = 0 (Z80 only; indicates ADD, not SUB)
'
' Examples:
'   ADD A           - Double A (A = A + A)
'   ADD M           - Add memory at (HL) to A
'   ADD B           - Add B to A
'
' Error Codes (base 120):
'   120: Missing operand (op1 is empty)
'   121: Invalid register (op1 register name not recognized)
' ==============================================================================
```

#### ADI (Add Immediate to A)

```vba
' ==============================================================================
' ADI data
' Opcode:   11000110 (next byte is immediate)
' 8080:     2 bytes, 7 cycles
' Z80:      2 bytes, 7 cycles
'
' Operation:   A ← A + immediate
' Flags:       Same as ADD (CY, P, AC, Z, S, N)
'
' Examples:
'   ADI 10h         - Add 0x10 to A (opcode: C6h 10h)
'
' Error Codes (base 130):
'   130: Missing operand (op1 is empty)
'   131: Invalid hex immediate (not valid hex or label resolves to invalid/out-of-range)
' ==============================================================================
```

#### ADC (Add with Carry) - 8080/Z80

```vba
' ==============================================================================
' ADC source
' Opcode:   10001ss (ss = 0-7)
' 8080:     1 byte, 4 cycles (7 if M)
' Z80:      1 byte, 4 cycles (7 if M)
'
' Operation:   A ← A + source + CY (includes carry-in)
' Flags:       Same as ADD
'
' Notes:
'   - Essential for multi-byte addition
'   - N flag must be cleared before ADC to indicate addition in progress
'
' Error Codes (base 110):
'   110: Missing operand (op1 is empty)
'   111: Invalid register (op1 register name not recognized)
' ==============================================================================
```

#### ACI (Add Immediate to A with Carry)

```vba
' ==============================================================================
' ACI data
' Opcode:   11001110 (next byte is immediate)
' 8080:     2 bytes, 7 cycles
' Z80:      2 bytes, 7 cycles
'
' Operation:   A ← A + immediate + CY (includes carry-in)
' Flags:       Same as ADC
'
' Notes:
'   - Immediate version of ADC
'   - Essential for multi-byte addition with immediate values
'
' Examples:
'   ACI 0x10        - Add 0x10 + carry to A
'
' Error Codes (base 100):
'   100: Missing operand (op1 is empty)
'   101: Invalid hex immediate (not valid hex or out-of-range)
' ==============================================================================
```

#### CMP (Compare Register with A)

```vba
' ==============================================================================
' CMP source
' Opcode:   10111ss (ss = 0-7)
' 8080:     1 byte, 4 cycles (7 if M)
' Z80:      1 byte, 4 cycles (7 if M)
'
' Operation:   A - source (perform subtraction, discard result, set flags)
' Flags Affected:
'   Z  = 1 if A == source
'   S  = 1 if result is negative
'   P  = 1 if even parity
'   CY = 1 if A < source (borrow)
'   AC = 1 if borrow from bit 3
'
' Notes:
'   - Does NOT modify A; only sets flags
'   - Useful for conditional jumps after comparison
'
' Examples:
'   MVI A, 0x50     - Load A with 0x50
'   CMP B           - Compare A with B (don't modify A)
'   JZ  EQUAL       - Jump if equal (Z flag set)
'
' Common Uses:
'   - Testing equality: CMP reg, then JZ
'   - Testing less-than: CMP reg, then JC
'
' Error Codes (base 190):
'   190: Missing operand (op1 is empty)
'   191: Invalid register (op1 register name not recognized)
' ==============================================================================
```

#### CPI (Compare Immediate with A)

```vba
' ==============================================================================
' CPI data
' Opcode:   11111110 (next byte is immediate)
' 8080:     2 bytes, 7 cycles
' Z80:      2 bytes, 7 cycles
'
' Operation:   A - immediate (perform subtraction, discard result, set flags)
' Flags:       Same as CMP
'
' Notes:
'   - Immediate version of CMP
'   - Does NOT modify A
'
' Examples:
'   CPI 0x50        - Compare A with 0x50
'   JZ  FOUND       - Jump if A was 0x50
'
' Error Codes (base 200):
'   200: Missing operand (op1 is empty)
'   201: Invalid hex immediate (not valid hex or out-of-range)
' ==============================================================================
```

#### INR (Increment Register or Memory)

```vba
' ==============================================================================
' INR register
' Opcode:   00ddd100 (ddd = 0-7 for B,C,D,E,H,L,M,A)
' 8080:     1 byte, 5 cycles (10 if M)
' Z80:      1 byte, 4 cycles (10 if M)
'
' Operation:   register ← register + 1
' Flags Affected:
'   Z  = 1 if result is 0
'   S  = 1 if result is 0x80 or higher
'   P  = 1 if even parity
'   AC = 1 if carry from bit 3
'   Note: CY (carry) is NOT affected
'
' Examples:
'   INR A           - Increment A
'   INR M           - Increment byte at (HL)
'   INR H           - Increment H
'
' Error Codes (base 280):
'   280: Missing operand (op1 is empty)
'   281: Invalid register (op1 register name not recognized)
' ==============================================================================
```

#### DCR (Decrement Register or Memory)

```vba
' ==============================================================================
' DCR register
' Opcode:   00ddd101 (ddd = 0-7)
' 8080:     1 byte, 5 cycles (10 if M)
' Z80:      1 byte, 4 cycles (10 if M)
'
' Operation:   register ← register - 1
' Flags:       Same as INR (Z, S, P, AC; NOT CY)
'
' Notes:
'   - Unlike SUB, DCR does not affect Carry flag
'   - Useful for loop counters
'
' Error Codes (base 230):
'   230: Missing operand (op1 is empty)
'   231: Invalid register (op1 register name not recognized)
' ==============================================================================
```

### Group 3: Logical Instructions

#### ANA (AND Register with A)

```vba
' ==============================================================================
' ANA source
' Opcode:   10100ss (ss = 0-7)
' 8080:     1 byte, 4 cycles (7 if M)
' Z80:      1 byte, 4 cycles (7 if M)
'
' Operation:   A ← A AND source (bitwise AND)
' Flags Affected:
'   All flags set based on result (Z, S, P)
'   CY = 0 (always cleared)
'   AC = 1 (always set, Z80 behavior)
'
' Examples:
'   ANA A           - Logical AND of A with itself (result = A, flags set)
'   ANA B           - Mask A by AND-ing with B
'   ANA M           - AND A with memory at (HL)
'
' Common Uses:
'   - Bit masking
'   - Testing bit patterns
'   - Clearing high nibble: ANA 0Fh
'
' Error Codes (base 140):
'   140: Missing operand (op1 is empty)
'   141: Invalid register (op1 register name not recognized)
' ==============================================================================
```

#### ANI (AND Immediate with A)

```vba
' ==============================================================================
' ANI data
' Opcode:   11100110 (next byte is immediate)
' 8080:     2 bytes, 7 cycles
' Z80:      2 bytes, 7 cycles
'
' Operation:   A ← A AND immediate
' Flags:       Same as ANA (Z, S, P; CY = 0, AC = 1)
'
' Examples:
'   ANI 0x0F        - Clear high nibble of A (mask with 0x0F)
'   ANI 0x80        - Test bit 7 (result 0 if not set)
'
' Common Uses:
'   - Bit pattern testing
'   - Clearing specific bit groups
'   - Extract low/high nibble
'
' Error Codes (base 150):
'   150: Missing operand (op1 is empty)
'   151: Invalid hex immediate (not valid hex or out-of-range)
' ==============================================================================
```

#### ORA (OR Register with A)

```vba
' ==============================================================================
' ORA source
' Opcode:   10110ss (ss = 0-7)
' 8080:     1 byte, 4 cycles (7 if M)
' Z80:      1 byte, 4 cycles (7 if M)
'
' Operation:   A ← A OR source (bitwise OR)
' Flags Affected:
'   Z, S, P set based on result
'   CY = 0 (always cleared)
'   AC = 0 (always cleared)
'
' Examples:
'   ORA B           - OR A with B
'   ORA A           - Sets flags on A, CY cleared
'   ORA M           - OR A with memory at (HL)
'
' Common Uses:
'   - Combining bit patterns
'   - Setting specific bits: ORA 80h (set bit 7)
'
' Error Codes (base 380):
'   380: Missing operand (op1 is empty)
' ==============================================================================
```

#### XRA (XOR Register with A)

```vba
' ==============================================================================
' XRA source
' Opcode:   10101ss (ss = 0-7)
' 8080:     1 byte, 4 cycles (7 if M)
' Z80:      1 byte, 4 cycles (7 if M)
'
' Operation:   A ← A XOR source (exclusive OR)
' Flags:       Z, S, P set; CY = 0, AC = 0
'
' Examples:
'   XRA A           - XOR A with itself (result = 0, all flags updated)
'   XRA B           - XOR A with B (toggle bits where B has 1s)
'
' Common Uses:
'   - Clear register: XRA A (faster than MVI A, 0)
'   - Toggle bits selectively
'   - Checksum/parity computation
'
' Error Codes (base 400):
'   400: Missing operand (op1 is empty)
'   401: Invalid register (op1 register name not recognized)
' ==============================================================================
```

#### SUB (Subtract Register from A)

```vba
' ==============================================================================
' SUB source
' Opcode:   10010ss (ss = 0-7)
' 8080:     1 byte, 4 cycles (7 if M)
' Z80:      1 byte, 4 cycles (7 if M)
'
' Operation:   A ← A - source
' Flags Affected:
'   CY = 1 if A < source (borrow)
'   P  = 1 if even parity
'   AC = 1 if borrow from bit 3
'   Z  = 1 if result is 0
'   S  = 1 if result is negative (bit 7 set)
'   N  = 1 (Z80 only; indicates SUB, not ADD)
'
' Examples:
'   SUB B           - Subtract B from A
'   SUB M           - Subtract memory at (HL) from A
'
' Error Codes (base 580):
'   580: Missing operand (op1 is empty)
'   581: Invalid register (op1 register name not recognized)
' ==============================================================================
```

#### SUI (Subtract Immediate from A)

```vba
' ==============================================================================
' SUI data
' Opcode:   11010110 (next byte is immediate)
' 8080:     2 bytes, 7 cycles
' Z80:      2 bytes, 7 cycles
'
' Operation:   A ← A - immediate
' Flags:       Same as SUB
'
' Examples:
'   SUI 0x10        - Subtract 0x10 from A
'
' Error Codes (base 590):
'   590: Missing operand (op1 is empty)
'   591: Invalid hex immediate (not valid hex or out-of-range)
' ==============================================================================
```

#### SBB (Subtract Register from A with Borrow)

```vba
' ==============================================================================
' SBB source
' Opcode:   10011ss (ss = 0-7)
' 8080:     1 byte, 4 cycles (7 if M)
' Z80:      1 byte, 4 cycles (7 if M)
'
' Operation:   A ← A - source - CY (includes carry-in as borrow)
' Flags:       Same as SUB
'
' Notes:
'   - Essential for multi-byte subtraction
'   - N flag must be set before SBB to indicate subtraction in progress
'
' Examples:
'   SUB B           - Start multi-byte subtraction
'   SBB C           - Continue with second byte (includes borrow)
'
' Error Codes (base 570):
'   570: Missing operand (op1 is empty)
'   571: Invalid register (op1 register name not recognized)
' ==============================================================================
```

#### ORI (OR Immediate with A)

```vba
' ==============================================================================
' ORI data
' Opcode:   11110110 (next byte is immediate)
' 8080:     2 bytes, 7 cycles
' Z80:      2 bytes, 7 cycles
'
' Operation:   A ← A OR immediate
' Flags:       Z, S, P set; CY = 0, AC = 0
'
' Examples:
'   ORI 0x80        - Set bit 7 of A (A = A | 0x80)
'
' Common Uses:
'   - Setting specific bit groups in A
'   - Combining bitmasks
'
' Error Codes (base 390):
'   390: Missing operand (op1 is empty)
'   391: Invalid hex immediate (not valid hex or out-of-range)
' ==============================================================================
```

### Group 4: Control Flow Instructions

#### JMP (Unconditional Jump)

```vba
' ==============================================================================
' JMP address
' Opcode:   11000011 (next 2 bytes are target address, LSB first)
' 8080:     3 bytes, 10 cycles
' Z80:      3 bytes, 10 cycles
'
' Operation:   PC ← address (jump unconditionally)
' Flags:       None affected
'
' Examples:
'   JMP 1000h       - Jump to address 0x1000 (opcode: C3h 00h 10h)
'
' Error Codes (base 310):
'   310: Missing operand (op1 is empty)
'   311: Label not found (referenced label does not exist)
'   312: Label resolves to invalid hex (address cannot be converted to hex)
'   313: Out-of-range address (address is beyond memory size)
' ==============================================================================
```

#### JC / JNC / JZ / JNZ / JP / JM / JPE / JPO (Conditional Jumps)

```vba
' ==============================================================================
' Conditional Jump Instructions
'
' Mnemonic | Opcode | Jump If...           | Flag Test
' ---------|--------|----------------------|-------------
' JC       | DA     | Carry = 1            | Jump if carry flag is SET
' JNC      | D2     | Carry = 0            | Jump if NO carry
' JZ       | CA     | Zero = 1             | Jump if result was ZERO
' JNZ      | C2     | Zero = 0             | Jump if NOT zero
' JP       | F2     | Sign = 0             | Jump if POSITIVE (sign bit clear)
' JM       | FA     | Sign = 1             | Jump if MINUS/NEGATIVE
' JPE      | EA     | Parity = 1           | Jump if EVEN parity
' JPO      | E2     | Parity = 0           | Jump if ODD parity
'
' 8080:     3 bytes each, 10 cycles if jump taken, 7 if not
' Z80:      3 bytes each, 10 cycles if jump taken, 7 if not
'
' Operation:   If condition true: PC ← address; else: PC ← PC + 3
' Flags:       None affected (test done before jump)
'
' Examples:
'   JC  1000h       - Jump to 0x1000 if carry flag is set
'   JNZ LOOP        - Jump back to LOOP label if NOT zero (continue loop)
'   JP  POSITIVE    - Jump if result was positive (high bit clear)
' ==============================================================================
```

#### CALL (Subroutine Call)

```vba
' ==============================================================================
' CALL address
' Opcode:   11001101 (next 2 bytes are subroutine address)
' 8080:     3 bytes, 17 cycles
' Z80:      3 bytes, 17 cycles
'
' Operation:
'   1. Push next instruction address (PC + 3) to stack
'   2. Jump to subroutine address
'   Stack grows downward: SP ← SP - 2
'
' Flags:       None affected
'
' Stack Effect:
'   Before:  SP = 0x200, memory[0x1FE:0x1FF] undefined
'   After:   SP = 0x1FE, memory[0x1FE] = PC_low, memory[0x1FF] = PC_high
'
' Examples:
'   CALL 3000h      - Call subroutine at 0x3000 (opcode: CDh 00h 30h)
'   CALL PRINT_HEX  - Call procedure labeled PRINT_HEX
'
' Notes:
'   - Return address is 16-bit (supports full 64K address space)
'   - Must be paired with RET to return properly
'
' Error Codes (base 160):
'   160: Missing operand (op1 is empty)
'   161: Label not found (referenced label/procedure does not exist)
' ==============================================================================
```

#### RET (Return from Subroutine)

```vba
' ==============================================================================
' RET
' Opcode:   11001001 (single byte)
' 8080:     1 byte, 10 cycles
' Z80:      1 byte, 10 cycles
'
' Operation:
'   1. Pop return address from stack into PC
'   2. Resume execution at return address
'   Stack grows upward: SP ← SP + 2
'
' Flags:       None affected
'
' Stack Effect:
'   Before:  SP = 0x1FE, memory[0x1FE] = 0x00, memory[0x1FF] = 0x10
'   After:   SP = 0x200, PC = 0x1000
'
' Examples:
'   RET             - Return from current subroutine
'
' Notes:
'   - Pops 16-bit address (LSB first, then MSB)
'   - Usually used at end of subroutine body
'
' Error Codes (base 460):
'   460: No address on stack (stack is empty - stack underflow)
'   461: Invalid return frame (stack top is not a CALL frame)
' ==============================================================================
```

### Group 5: Stack Instructions

#### PUSH (Push Register Pair onto Stack)

```vba
' ==============================================================================
' PUSH register_pair
' Opcode:   11rr0101 (rr = 0-3 for BC,DE,HL,PSW)
' 8080:     1 byte, 11 cycles
' Z80:      1 byte, 11 cycles
'
' Operation:
'   1. SP ← SP - 2
'   2. memory[SP] ← register_high
'   3. memory[SP+1] ← register_low
'
' Flags:       None affected
'
' Stack Diagram (grows downward, toward lower addresses):
'   Before:  SP = 0x200 ▲
'                     │
'   After:   SP = 0x1FE ▲ (pushed 2 bytes here)
'
' Examples:
'   PUSH B          - Push BC onto stack (opcode: C5h)
'   PUSH PSW        - Push A and Flags onto stack (opcode: F5h)
'   PUSH H          - Push HL onto stack (opcode: E5h)
'
' Notes:
'   - PSW = (A register + Flags): high byte = A, low byte = flags
'   - Always pushes 2 bytes
'
' Error Codes (base 430):
'   430: Missing operand (op1 is empty)
'   431: Invalid register (op1 register name not recognized)
' ==============================================================================
```

#### POP (Pop Register Pair from Stack)

```vba
' ==============================================================================
' POP register_pair
' Opcode:   11rr0001 (rr = 0-3 for BC,DE,HL,PSW)
' 8080:     1 byte, 10 cycles
' Z80:      1 byte, 10 cycles
'
' Operation:
'   1. register_low ← memory[SP]
'   2. register_high ← memory[SP+1]
'   3. SP ← SP + 2
'
' Flags:       If PSW: Flags restored from stack; else: no effect
'
' Stack Diagram (grows upward, toward higher addresses):
'   Before:  SP = 0x1FE ▼ (pop from here)
'                     │
'   After:   SP = 0x200 ▼
'
' Examples:
'   POP B           - Pop BC from stack (opcode: C1h)
'   POP PSW         - Pop Flags and A from stack (opcode: F1h)
'   POP H           - Pop HL from stack (opcode: E1h)
'
' Complement to PUSH - always restores 2 bytes.
'
' Error Codes (base 420):
'   420: Missing operand (op1 is empty)
'   421: Invalid register (op1 register name not recognized)
' ==============================================================================
```

### Group 6: Special Instructions

#### HLT (Halt Execution)

```vba
' ==============================================================================
' HLT
' Opcode:   01110110 (single byte)
' 8080:     1 byte, 5 cycles (final)
' Z80:      1 byte, 4 cycles (final)
'
' Operation:   CPU halts; PC does not advance
' Flags:       None affected
'
' Notes:
'   - Terminates program execution
'   - Equivalent to END or program exit
'   - In your emulator, raises "Halt" error code for clean termination
'
' Examples:
'   HLT             - End program
'
' Return Code (special):
'   -100: Program halted (normal termination signal)
' ==============================================================================
```

---

## Part 3: Z80 Extended Instructions Reference

### Z80 Index Register Instructions (IX, IY)

#### Indexed Addressing Mode: (IX+d) and (IY+d)

```vba
' ==============================================================================
' Indexed Addressing: (IX+d) and (IY+d)
'
' Purpose:    Access memory at (Index Register + signed displacement)
' Example:    LD A, (IX+5)     - Load A from memory at IX + 5
'             LD (IY-2), B     - Store B to memory at IY - 2
'
' Operand (d) is a signed 8-bit value (-128 to +127)
'
' 8080:       NOT AVAILABLE
' Z80:        Available with DD or FD prefix
'   DD = IX prefix
'   FD = IY prefix
'
' Execution:  address = Index_Reg + sign_extend(d)
'             Then perform operation at [address]
'
' Cycles:     Generally +4 cycles vs non-indexed variant
'             Example: LD A, (HL) = 7 cycles
'                      LD A, (IX+0) = 15 cycles (DD+opcode+displacement)
'
' Notes:
'   - Displacement is signed; allows negative offsets
'   - Essential for Z80 array/structure access
'   - Requires 2-byte opcode: DD/FD, then regular opcode
' ==============================================================================
```

### Z80 Bit Manipulation Instructions (CB Prefix)

#### BIT (Test Bit)

```vba
' ==============================================================================
' BIT bit_pos, source
' Opcode:   CB 01bbsss (bb = bit 0-7, sss = source register)
' 8080:     NOT AVAILABLE
' Z80:      2 bytes, 8 cycles (12 if memory)
'
' Operation:   Test bit in source; set Zero flag if bit is 0
' Flags Affected:
'   Z = 1 if bit is 0 (bit not set)
'   Z = 0 if bit is 1 (bit is set)
'   S, P affected per result
'
' Examples:
'   BIT 0, A        - Test bit 0 of A (check if LSB is set)
'   BIT 7, B        - Test bit 7 of B (check sign bit)
'   BIT 3, M        - Test bit 3 of memory at (HL)
'
' Common Use:
'   BIT 0, A
'   JZ IS_EVEN      - Jump if bit 0 clear (even number)
' ==============================================================================
```

#### SET (Set Bit)

```vba
' ==============================================================================
' SET bit_pos, dest
' Opcode:   CB 11bbddd (bb = bit 0-7, ddd = dest register)
' 8080:     NOT AVAILABLE
' Z80:      2 bytes, 8 cycles (15 if memory)
'
' Operation:   Set specified bit to 1 in destination
' Flags:       None affected
'
' Examples:
'   SET 7, A        - Set bit 7 of A (0x80 | A)
'   SET 0, C        - Set bit 0 of C
'   SET 4, M        - Set bit 4 of memory at (HL)
' ==============================================================================
```

#### RES (Reset Bit)

```vba
' ==============================================================================
' RES bit_pos, dest
' Opcode:   CB 10bbddd (bb = bit 0-7, ddd = dest register)
' 8080:     NOT AVAILABLE
' Z80:      2 bytes, 8 cycles (15 if memory)
'
' Operation:   Clear specified bit to 0 in destination
' Flags:       None affected
'
' Examples:
'   RES 7, A        - Clear bit 7 of A (A & 0x7F)
'   RES 0, H        - Clear bit 0 of H
' ==============================================================================
```

### Z80 Rotate/Shift Instructions (CB Prefix)

#### RLC (Rotate Left Circular)

```vba
' ==============================================================================
' RLC dest
' Opcode:   CB 00000ddd (ddd = destination)
' 8080:     NOT AVAILABLE
' Z80:      2 bytes, 8 cycles (15 if memory)
'
' Operation:
'   Rotate left; bit 7 moves to Carry; Carry moves to bit 0
'   Before:  [bit7][bit6]...[bit0] CY
'   After:   [bit6]...[bit0][CY']   [old_bit7]
'
' Flags:       CY = old bit 7, others per result
'
' Example:
'   LD A, 0x80      ; A = 1000 0000 (bit 7 set)
'   RLC A           ; A = 0000 0001, CY = 1
' ==============================================================================
```

#### RRC (Rotate Right Circular)

```vba
' ==============================================================================
' RRC dest
' Opcode:   CB 00001ddd
' 8080:     NOT AVAILABLE
' Z80:      2 bytes, 8 cycles (15 if memory)
'
' Operation:
'   Rotate right; bit 0 moves to Carry; Carry moves to bit 7
'   Before:  CY [bit7]...[bit1][bit0]
'   After:   [old_bit0] [bit7]...[bit1]  [CY']
'
' Flags:       CY = old bit 0, others per result
' ==============================================================================
```

#### SLA (Shift Left Arithmetic)

```vba
' ==============================================================================
' SLA dest
' Opcode:   CB 00100ddd
' 8080:     NOT AVAILABLE
' Z80:      2 bytes, 8 cycles (15 if memory)
'
' Operation:
'   Shift left; bit 0 becomes 0; bit 7 moves to Carry
'   Before:  [bit7][bit6]...[bit1][bit0]
'   After:   [bit6]...[bit1][bit0][0] <- Carry gets old bit7
'
' Flags:       CY = old bit 7, Z,S,P per result
'
' Notes:
'   - Equivalent to multiply by 2
'   - Sign bit preserved implicitly (next SLA will rotate it out)
' ==============================================================================
```

#### SRL (Shift Right Logical)

```vba
' ==============================================================================
' SRL dest
' Opcode:   CB 00111ddd
' 8080:     NOT AVAILABLE
' Z80:      2 bytes, 8 cycles (15 if memory)
'
' Operation:
'   Shift right; bit 7 becomes 0; bit 0 moves to Carry
'   Before:  [bit7][bit6]...[bit1][bit0]
'   After:   [0][bit7]...[bit2][bit1] <- Carry gets old bit0
'
' Flags:       CY = old bit 0, Z,S,P per result
'
' Notes:
'   - Equivalent to unsigned divide by 2
' ==============================================================================
```

#### SRA (Shift Right Arithmetic)

```vba
' ==============================================================================
' SRA dest
' Opcode:   CB 00101ddd
' 8080:     NOT AVAILABLE
' Z80:      2 bytes, 8 cycles (15 if memory)
'
' Operation:
'   Shift right; bit 7 preserved (sign-extending); bit 0 to Carry
'   Before:  [bit7][bit6]...[bit1][bit0]
'   After:   [bit7][bit7]...[bit2][bit1] <- Carry gets old bit0
'
' Flags:       CY = old bit 0, Z,S,P per result
'
' Notes:
'   - Preserves sign (arithmetic right shift)
'   - Equivalent to signed divide by 2
'   - Negative numbers shift in with 1s (preserved sign)
' ==============================================================================
```

### Z80 Register Exchange Instructions

#### EXX (Exchange Primary and Shadow Registers)

```vba
' ==============================================================================
' EXX
' Opcode:   11011001 (single byte)
' 8080:     NOT AVAILABLE
' Z80:      1 byte, 4 cycles
'
' Operation:
'   Swap (B,C) ← → (B',C')
'   Swap (D,E) ← → (D',E')
'   Swap (H,L) ← → (H',L')
'   Note: A and F are NOT swapped; use separate EX AF, AF' for that
'
' Flags:       None affected (flags not swapped unless using EX AF, AF')
'
' Usage:
'   EXX             - Quickly swap register set for interrupt handler
'   EX AF, AF'      - Swap A and F with shadow versions
'
' Example:
'   Original:    B=0x42, C=0x99, B'=0x11, C'=0x22
'   After EXX:   B=0x11, C=0x22, B'=0x42, C'=0x99
' ==============================================================================
```

#### EX (Exchange Register Pairs)

```vba
' ==============================================================================
' EX dest, source
' Variants:
'   EX DE, HL       - Swap DE with HL (opcode: EBh)
'   EX AF, AF'      - Swap A and F with shadows (opcode: 08h)
'   EX (SP), HL     - Swap (SP) with HL (1-byte stack operation)
'
' 8080:     Limited support (varies)
' Z80:      Full support (1-2 bytes, 4-19 cycles)
'
' Operation:
'   Exchanges values of two registers
'
' Examples:
'   EX DE, HL       - Swap DE and HL values
'   EX (SP), HL     - Pop HL from stack, push original HL to stack
' ==============================================================================
```

---

## Part 3.5: Error Code Reference

Your emulator uses error codes in the trace output (column 19: "Err") to indicate the result of instruction execution. Here are the error codes used:

```vba
' ==============================================================================
' Error Codes (errcode parameter in TraceAppend)
' ==============================================================================
'
' Value | Meaning                      | When Set
' ------|------------------------------|-------------------------------------------
'   0   | SUCCESS (no error)           | Instruction executed normally
'   
'  -1   | HLT (Halt)                   | HLT instruction encountered
'       |                              | Program terminated cleanly
'       |
'  -2   | INVALID REGISTER             | Invalid register name/operand
'       |                              | (not found in register dictionary)
'       |
'  -3   | INVALID MEMORY ADDRESS       | Memory access out of bounds
'       |                              | (address >= MemSize or < 0)
'       |
'  -4   | STACK UNDERFLOW              | POP from empty stack
'       |                              | (SP would exceed top of memory)
'       |
'  -5   | STACK OVERFLOW               | PUSH to full stack
'       |                              | (SP would go below 0)
'       |
'  -6   | INVALID OPCODE               | Unrecognized instruction
'       |                              | (not in RunOpcode case statements)
'       |
'  -7   | INVALID OPERAND              | Operand parsing/validation failed
'       |                              | (bad register name, invalid data)
'       |
'  -8   | LABEL NOT FOUND              | JMP/CALL to undefined label
'       |                              | (label not registered)
'       |
'  -9   | UNIMPLEMENTED INSTRUCTION    | Z80 instruction not yet implemented
'       |                              | (CB/ED/DD/FD prefix without handler)
'       |
' -10   | IMMEDIATE VALUE OUT OF RANGE | Immediate operand > 255
'       |                              | (for 8-bit loads)
'       |
' -11   | PROCEDURE NOT FOUND          | CALL to undefined procedure
'       |                              | (procedure not in registry)
'       |
' Positive values (> 0) are reserved for user-defined status codes or
' warnings in future implementations.
' ==============================================================================
```

### Usage in Trace Output

When **Trace** is enabled, each instruction execution creates a row in the Trace sheet with:
- Column 19 (**Err**): Error code from execution
  - **0** = Instruction succeeded
  - **Non-zero** = Error occurred; execution may have halted

### Error Handling in Your Code

When calling `TraceAppend()`, pass the error code as the **errcode** parameter:

```vba
' Example: Trace an instruction that succeeded
Call modTrace.TraceAppend(cpu, pc_dec, opcode, op1, op2, 0)

' Example: Trace an instruction that failed  
Call modTrace.TraceAppend(cpu, pc_dec, opcode, op1, op2, -2)  ' Invalid register
```

### Implementing Error Codes

In your instruction functions, return the appropriate error code:

```vba
Public Function ADD(ByVal regName As String) As Long
    Dim sourceValue As Long
    Dim result As Long
    
    ' Validate register
    If Not pRegs.Exists(regName) Then
        ADD = -2  ' Invalid register error
        Exit Function
    End If
    
    ' Get operand value
    sourceValue = pRegs(regName)
    
    ' Perform addition
    result = (pRegs("A") + sourceValue) And 255
    pRegs("A") = result
    
    ' Update flags...
    SetCarryFlag result
    SetZeroFlag result
    ' etc.
    
    ADD = 0  ' Success
End Function
```

---

## Part 4: Flag Reference

### 8080 Flags

```vba
' ==============================================================================
' 8080 Flag Register (Flags stored as individual bits)
'
' Bit | Name  | Set (1) If...                  | Clear (0) If...
' ----|-------|--------------------------------|---------------------------------
'  0  | CY    | Carry/Borrow out               | No carry/borrow
'     |       | (result > 255 for ADD)         | (result fits in 8 bits)
' ----|-------|--------------------------------|---------------------------------
'  2  | P     | Result has EVEN number of      | Result has ODD number of
'     |       | set bits (even parity)         | set bits (odd parity)
' ----|-------|--------------------------------|---------------------------------
'  4  | AC    | Carry out of bit 3             | No carry from bit 3
'     |       | (BCD half-carry)               |
' ----|-------|--------------------------------|---------------------------------
'  6  | Z     | Result is 0x00                 | Result is non-zero
' ----|-------|--------------------------------|---------------------------------
'  7  | S     | Bit 7 of result is 1           | Bit 7 of result is 0
'     |       | (negative in 2's complement)   | (positive)
' ----|-------|--------------------------------|---------------------------------
'
' PSW (Program Status Word) = (A register << 8) | (Flags register)
'   - Used by PUSH PSW / POP PSW
'   - Combines A and Flags for stack operation
' ==============================================================================
```

### Z80 Flag Additions

```vba
' ==============================================================================
' Z80 Additional Flags
'
' Bit | Name  | Purpose
' ----|-------|-------------------------------
'  1  | N     | Add/Subtract flag (used by DAA)
'     |       | 0 = result from ADD/ADC operation
'     |       | 1 = result from SUB/SBC operation
' ----|-------|-------------------------------
'  3  | 3     | Undocumented flag (copy of bit 3 of result)
' ----|-------|-------------------------------
'  5  | 5     | Undocumented flag (copy of bit 5 of result)
' ----|-------|-------------------------------
'
' Note: Bits 3 and 5 are undocumented but used in advanced Z80 code.
' Typical implementation: copy result bits 3 and 5 to flag register.
'
' DAA (Decimal Adjust Accumulator) uses:
'   - CY = carry from BCD operation
'   - AC = half-carry from BCD operation
'   - N = indicates last operation was subtract (Z80 DAA variant)
' ==============================================================================
```

---

## Part 5: Common Implementation Patterns

### Pattern: Calculate Parity Flag

```vba
' Algorithm: Count set bits; if even, P = 1; if odd, P = 0
Function SetParity(ByVal result As Long) As Integer
    Dim bitCount As Integer: bitCount = 0
    Dim temp As Long: temp = result And 255 ' mask to 8 bits
    
    Do While temp > 0
        If (temp And 1) Then bitCount = bitCount + 1
        temp = temp \ 2
    Loop
    
    ' Even number of bits set => P = 1
    SetParity = IIf((bitCount Mod 2) = 0, 1, 0)
End Function
```

### Pattern: Calculate Sign Flag

```vba
' Algorithm: S = bit 7 of result
Function SetSign(ByVal result As Long) As Integer
    SetSign = IIf((result And 128) <> 0, 1, 0)
End Function
```

### Pattern: Calculate Zero Flag

```vba
' Algorithm: Z = 1 if (result AND 0xFF) = 0
Function SetZero(ByVal result As Long) As Integer
    SetZero = IIf((result And 255) = 0, 1, 0)
End Function
```

### Pattern: Calculate Carry Flag (Add)

```vba
' Algorithm: For ADD, CY = 1 if result > 255
Function SetCarryAdd(ByVal op1 As Long, ByVal op2 As Long) As Integer
    SetCarryAdd = IIf(((op1 And 255) + (op2 And 255)) > 255, 1, 0)
End Function
```

### Pattern: Calculate Carry Flag (Subtract)

```vba
' Algorithm: For SUB, CY = 1 if op1 < op2 (borrow)
Function SetCarrySub(ByVal op1 As Long, ByVal op2 As Long) As Integer
    SetCarrySub = IIf((op1 And 255) < (op2 And 255), 1, 0)
End Function
```

### Pattern: Calculate Aux Carry (Half-Carry)

```vba
' Algorithm: AC = 1 if carry from bit 3 to bit 4
Function SetAuxCarry(ByVal op1 As Long, ByVal op2 As Long) As Integer
    Dim bit3Only As Long: bit3Only = 8 ' 0x08, bits 0-3
    Dim result As Long: result = (op1 And bit3Only) + (op2 And bit3Only)
    SetAuxCarry = IIf(result > bit3Only, 1, 0)
End Function
```

---

## Part 6: Testing Checklist for Each Instruction

Use this checklist when implementing and testing a new instruction:

```markdown
## Instruction: [NAME]

- [ ] **Opcode Verification**
  - [ ] Correct hex value(s) documented
  - [ ] All addressing modes covered (if applicable)

- [ ] **Operation**
  - [ ] Correct registers modified
  - [ ] Correct memory accessed (if applicable)
  - [ ] Operand order correct

- [ ] **Flags**
  - [ ] Carry flag set correctly (if affected)
  - [ ] Zero flag set correctly
  - [ ] Sign flag set correctly
  - [ ] Parity flag set correctly
  - [ ] Aux Carry set correctly (if affected)
  - [ ] N flag set correctly (Z80)

- [ ] **Edge Cases**
  - [ ] Result = 0x00 (minimum)
  - [ ] Result = 0xFF (maximum)
  - [ ] Result = 0x80 (sign bit only)
  - [ ] Result = 0x7F (max positive)
  - [ ] Overflow/Underflow behavior

- [ ] **Memory Access** (if applicable)
  - [ ] (HL) addressing works
  - [ ] (IX+d) / (IY+d) displacement correct (Z80)
  - [ ] Bounds checking performed

- [ ] **Z80 Differences** (if applicable)
  - [ ] N flag set per Z80 spec (if different from 8080)
  - [ ] Any undocumented flag behavior documented

- [ ] **Test Cases Written**
  - [ ] Basic operation test
  - [ ] Flag-setting test
  - [ ] Edge case test
  - [ ] Memory operation test (if applicable)
  - [ ] Timing validation (cycle count)

- [ ] **Documentation Added**
  - [ ] Instruction header comment added
  - [ ] Operation description clear
  - [ ] All flags documented
  - [ ] Examples provided
  - [ ] Z80 vs 8080 differences noted
```

---

## Part 7: Complete Error Code Reference

This section documents all error codes generated by your instruction implementations. Each instruction function returns an error code (0 for success, positive integer for error).

### Error Code Lookup Table

| Instruction | ErrorBase | Offsets | Status |
|-------------|-----------|---------|--------|
| ACI | 100 | 0, 1 | ✓ |
| ADD | 120 | 0, 1 | ✓ |
| ADI | 130 | 0, 1 | ✓ |
| ADC | 110 | 0, 1 | ✓ |
| ANA | 140 | 0, 1 | ✓ |
| ANI | 150 | 0, 1 | ✓ |
| CALL | 160 | 0, 1 | ✓ |
| CMP | 190 | 0, 1 | ✓ |
| CPI | 200 | 0, 1 | ✓ |
| DCR | 230 | 0, 1 | ✓ |
| EQU | 300 | 0, 1, 2, 3 | ✓ |
| HLT | -100 | Special | ✓ |
| INR | 280 | 0, 1 | ✓ |
| JMP | 310 | 0, 1, 2, 3 | ✓ |
| LXI | 340 | 0, 1, 2 | ✓ |
| MOV | 350 | 0, 1 | ✓ |
| MVI | 360 | 0, 1, 2 | ✓ |
| ORA | 380 | 0 | ✓ |
| ORI | 390 | 0, 1 | ✓ |
| POP | 420 | 0, 1 | ✓ |
| PUSH | 430 | 0, 1 | ✓ |
| RET | 460 | 0, 1 | ✓ |
| SUB | 580 | 0, 1 | ✓ |
| SUI | 590 | 0, 1 | ✓ |
| SBB | 570 | 0, 1 | ✓ |
| XRA | 400 | 0, 1 | ✓ |

### Single-Operand Arithmetic/Logic Instructions

#### ADD - Add Register to A (Base 120)

```
120: Missing operand (op1 is empty)
121: Invalid register (register not recognized)
```

#### ADI - Add Immediate to A (Base 130)

```
130: Missing operand (op1 is empty)
131: Invalid hex immediate (value not hex or label unresolved)
```

#### ADC - Add with Carry (Base 110)

```
110: Missing operand (op1 is empty)
111: Invalid register (register not recognized)
```

#### ACI - Add Immediate with Carry (Base 100)

```
100: Missing operand (op1 is empty)
101: Invalid hex immediate (value not hex)
```

#### ANA - Logical AND Register (Base 140)

```
140: Missing operand (op1 is empty)
141: Invalid register (register not recognized)
```

#### ANI - Logical AND Immediate (Base 150)

```
150: Missing operand (op1 is empty)
151: Invalid hex immediate (value not hex)
```

#### CMP - Compare Register (Base 190)

```
190: Missing operand (op1 is empty)
191: Invalid register (register not recognized)
```

#### CPI - Compare Immediate (Base 200)

```
200: Missing operand (op1 is empty)
201: Invalid hex immediate (value not hex)
```

#### INR - Increment Register (Base 280)

```
280: Missing operand (op1 is empty)
281: Invalid register (register not recognized)
```

#### DCR - Decrement Register (Base 230)

```
230: Missing operand (op1 is empty)
231: Invalid register (register not recognized)
```

#### ORA - Logical OR Register (Base 380)

```
380: Missing operand (op1 is empty)
```

#### ORI - Logical OR Immediate (Base 390)

```
390: Missing operand (op1 is empty)
391: Invalid hex immediate (value not hex or label unresolved)
```

#### SUB - Subtract Register (Base 580)

```
580: Missing operand (op1 is empty)
581: Invalid register (register not recognized)
```

#### SUI - Subtract Immediate (Base 590)

```
590: Missing operand (op1 is empty)
591: Invalid hex immediate (value not hex)
```

#### SBB - Subtract with Borrow (Base 570)

```
570: Missing operand (op1 is empty)
571: Invalid register (register not recognized)
```

#### XRA - Logical XOR Register (Base 400)

```
400: Missing operand (op1 is empty)
401: Invalid register (register not recognized)
```

### Two-Operand Instructions

#### MOV - Move Register to Register (Base 350)

```
350: Missing operands (op1 or op2 empty)
351: Invalid register(s) (source or destination not recognized)
```

#### MVI - Move Immediate to Register (Base 360)

```
360: Missing operands (op1 or op2 empty)
361: Invalid hex value (immediate not hex)
362: Invalid register (register not recognized)
```

#### LXI - Load 16-bit Immediate (Base 340)

```
340: Missing operands (op1 or op2 empty)
341: Invalid register pair (must be B, D, H, or SP)
342: Invalid address (address not hex or label unresolved)
```

### Three+ Operand / Multi-Parameter Instructions

#### CALL - Call Subroutine (Base 160)

```
160: Missing operand (op1 empty)
161: Label not found (referenced label/procedure doesn't exist)
```

#### JMP - Unconditional Jump (Base 310)

```
310: Missing operand (op1 empty)
311: Label not found (referenced label doesn't exist)
312: Label resolves to invalid hex (address invalid)
313: Out-of-range address (address beyond memory)
```

#### PUSH - Push Register (Base 430)

```
430: Missing operand (op1 empty)
431: Invalid register (register not recognized)
```

#### POP - Pop Register (Base 420)

```
420: Missing operand (op1 empty)
421: Invalid register (register not recognized)
```

#### RET - Return from Subroutine (Base 460)

```
460: No address on stack (stack underflow)
461: Invalid return frame (stack top not CALL frame)
```

### Assembler Pseudo-Instructions

#### EQU - Define Constant (Base 300)

```
300: Missing value (op1 empty)
301: Missing label name (label empty)
302: Duplicate label (label already defined)
303: Invalid hex value (value not hex)
```

### Special/Control Instructions

#### HLT - Halt (Base -100)

```
-100: Program halted (normal program termination)
```

---

### Using Error Codes with TraceAppend()

When logging execution in trace output, pass the error code as the `errcode` parameter:

```vba
' In your RunOpcode or instruction dispatcher:
Dim errcode As Long
errcode = cpu.ADD("B")  ' Returns 0 on success, errorBase+offset on error
Call modTrace.TraceAppend(cpu, pc, "ADD", "B", "", errcode)
```

### Debugging with Error Codes

**Example Trace Output:**
```
Step | PC   | Op  | OP1 | OP2  | ... | Err
-----|------|-----|-----|------|-----|-----
  10 | 0042 | MVI | A   | GGGG | ... | 361
  11 | 0043 | ADD | A   |      | ... | 120
  12 | 0044 | JMP | XXXX|      | ... | 311
```

**Interpretation:**
- **Row 10, Err=361:** MVI instruction, errorBase=360, offset=1 → Invalid hex value
- **Row 11, Err=120:** ADD instruction, errorBase=120, offset=0 → Missing operand
- **Row 12, Err=311:** JMP instruction, errorBase=310, offset=1 → Label not found

**How to Fix:**
1. Find the error code in the Quick Reference table
2. Look up the instruction name and errorBase
3. Calculate offset: 361 - 360 = 1
4. Read the error message for that offset
5. Check OP1/OP2 columns for the problematic data

---

## Part 8: Comprehensive Implementation Checklist

When implementing a new instruction, follow this complete checklist:

### Step 1: Choose ErrorBase

- Pick a base number (10s) that doesn't conflict with existing instructions
- Check the Error Code Lookup Table above
- Common ranges: 100-199 (single-operand), 300-399 (two-operand), 400+ (special)

### Step 2: Document Errors in Header

```vba
' Errors (base XXX):
'   XXX: missing operand
'   XXX+1: invalid register
'   XXX+2: [additional error if needed]
```

### Step 3: Implement Validation

```vba
Dim errorBase As Long: errorBase = XXX

If op1 = "" Then SetError errorBase, "...": InstructionName = errorBase: Exit Function
If Not IsValidReg(op1) Then SetError errorBase + 1, "...": InstructionName = errorBase + 1: Exit Function
' Add more validations as needed
```

### Step 4: Update This Document

Add your instruction to the Error Code Lookup Table and create an entry in the appropriate section.

---

## References & Further Reading

- **Intel 8080 Datasheet** - Original specification
- **Zilog Z80 User Manual** - Z80 extensions and differences
- **Zaks, Rodnay. "Programming the Z80" (1982)** - Comprehensive reference
- **Sean Young's Z80 Guide** - Modern online reference
- **WikiChip Z80 Instruction Set** - Opcode table
- **http://www.z80.info** - Comprehensive Z80 documentation

---

**Document Version:** 1.0  
**Last Updated:** March 6, 2026  
**For Use With:** Z80_Model_Current_2026-03-06_2239.xlsm
