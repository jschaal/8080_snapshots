# Z80/8080 CPU Emulator - Error Code Reference

**Date:** March 7, 2026  
**Version:** 2.0 (Extracted from VBA Source)  
**Purpose:** Complete error code reference for all instruction implementations

---

## Overview

Each instruction function defines its own `errorBase` constant and generates error codes by adding numeric offsets to that base. This allows fine-grained error reporting in trace output and enables systematic debugging.

### Core Principle

```vba
Dim errorBase As Long: errorBase = XXX
If condition Then
    SetError errorBase + offset, "Description"
    InstructionName = errorBase + offset
    Exit Function
End If
InstructionName = 0  ' Success
```

### How to Read This Document

1. **Look up the instruction** from the trace "Op" column
2. **Find the errorBase** in the table below
3. **Read the error code offset** for that condition
4. **Interpret** the specific error message

---

## Quick Reference Table

| Instruction | ErrorBase | Offsets | Status |
|-------------|-----------|---------|--------|
| ADD | 120 | 0, 1 | ✓ Complete |
| ADI | 130 | 0, 1 | ✓ Complete |
| ACI | 100 | 0, 1 | ✓ Complete |
| ADC | 110 | 0, 1 | ✓ Complete |
| ANA | 140 | 0, 1 | ✓ Complete |
| ANI | 150 | 0, 1 | ✓ Complete |
| CMP | 190 | 0, 1 | ✓ Complete |
| CPI | 200 | 0, 1 | ✓ Complete |
| CALL | 160 | 0, 1 | ✓ Complete |
| DCR | 230 | 0, 1 | ✓ Complete |
| EQU | 300 | 0, 1, 2, 3 | ✓ Complete |
| HLT | -100 | 0 | ✓ Complete |
| INR | 280 | 0, 1 | ✓ Complete |
| JMP | 310 | 0, 1, 2, 3 | ✓ Complete |
| LXI | 340 | 0, 1, 2 | ✓ Complete |
| MOV | 350 | 0, 1 | ✓ Complete |
| MVI | 360 | 0, 1, 2 | ✓ Complete |
| ORA | 380 | 0 | ✓ Complete |
| ORI | 390 | 0, 1 | ✓ Complete |
| POP | 420 | 0, 1 | ✓ Complete |
| PUSH | 430 | 0, 1 | ✓ Complete |
| RET | 460 | 0, 1 | ✓ Complete |
| SUB | 580 | 0, 1 | ✓ Complete |
| SUI | 590 | 0, 1 | ✓ Complete |
| SBB | 570 | 0, 1 | ✓ Complete |
| XRA | 400 | 0, 1 | ✓ Complete |

---

## Detailed Error Code Reference

### ADD - Add Register to Accumulator

**Base:** 120 | **Syntax:** `ADD reg`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 120 | 0 | Missing operand | op1 is empty |
| 121 | 1 | Invalid register | Register name not recognized |

**Returns:** 0 on success

---

### ADI - Add Immediate to Accumulator

**Base:** 130 | **Syntax:** `ADI imm8`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 130 | 0 | Missing operand | op1 is empty |
| 131 | 1 | Invalid hex immediate | Value is not valid hex or label resolves to invalid/out-of-range |

**Returns:** 0 on success

---

### ACI - Add Immediate to Accumulator with Carry

**Base:** 100 | **Syntax:** `ACI imm8`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 100 | 0 | Missing operand | op1 is empty |
| 101 | 1 | Invalid hex immediate | Value is not valid hex |

**Returns:** 0 on success

---

### ADC - Add Register to Accumulator with Carry

**Base:** 110 | **Syntax:** `ADC reg`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 110 | 0 | Missing operand | op1 is empty |
| 111 | 1 | Invalid register | Register name not recognized |

**Returns:** 0 on success

---

### ANA - Logical AND Register with Accumulator

**Base:** 140 | **Syntax:** `ANA reg`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 140 | 0 | Missing operand | op1 is empty |
| 141 | 1 | Invalid register | Register name not recognized |

**Returns:** 0 on success

---

### ANI - Logical AND Immediate with Accumulator

**Base:** 150 | **Syntax:** `ANI imm8`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 150 | 0 | Missing operand | op1 is empty |
| 151 | 1 | Invalid hex immediate | Value is not valid hex |

**Returns:** 0 on success

---

### CMP - Compare Register with Accumulator

**Base:** 190 | **Syntax:** `CMP reg`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 190 | 0 | Missing operand | op1 is empty |
| 191 | 1 | Invalid register | Register name not recognized |

**Returns:** 0 on success

---

### CPI - Compare Immediate with Accumulator

**Base:** 200 | **Syntax:** `CPI imm8`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 200 | 0 | Missing operand | op1 is empty |
| 201 | 1 | Invalid hex immediate | Value is not valid hex |

**Returns:** 0 on success

---

### CALL - Call Subroutine

**Base:** 160 | **Syntax:** `CALL addr|label` or `CALL procedure_name`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 160 | 0 | Missing operand | op1 is empty |
| 161 | 1 | Label not found | Referenced label/procedure does not exist |

**Returns:** 0 on success

---

### DCR - Decrement Register

**Base:** 230 | **Syntax:** `DCR reg`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 230 | 0 | Missing operand | op1 is empty |
| 231 | 1 | Invalid register | Register name not recognized |

**Returns:** 0 on success

---

### EQU - Define Constant/Label

**Base:** 300 | **Syntax:** `label EQU value`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 300 | 0 | Missing value | op1 (value) is empty |
| 301 | 1 | Missing label name | Label name is empty |
| 302 | 2 | Duplicate label | Label already defined |
| 303 | 3 | Invalid hex value | Value is not valid hex |

**Returns:** 0 on success

---

### HLT - Halt Processor

**Base:** -100 | **Syntax:** `HLT`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| -100 | 0 | Program halted | Normal halt signal; end of execution |

**Returns:** -100 (program termination signal)

---

### INR - Increment Register

**Base:** 280 | **Syntax:** `INR reg`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 280 | 0 | Missing operand | op1 is empty |
| 281 | 1 | Invalid register | Register name not recognized |

**Returns:** 0 on success

---

### JMP - Unconditional Jump

**Base:** 310 | **Syntax:** `JMP addr|label`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 310 | 0 | Missing operand | op1 is empty |
| 311 | 1 | Label not found | Referenced label does not exist |
| 312 | 2 | Label resolves to invalid hex | Address cannot be converted to hex |
| 313 | 3 | Out-of-range address | Address is beyond memory size |

**Returns:** 0 on success

---

### LXI - Load 16-bit Immediate into Register Pair

**Base:** 340 | **Syntax:** `LXI rp, addr|label` where rp = B|D|H|SP

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 340 | 0 | Missing operands | op1 or op2 is empty |
| 341 | 1 | Invalid register pair | Register is not B, D, H, or SP |
| 342 | 2 | Invalid address | Address is not valid hex or label doesn't resolve |

**Returns:** 0 on success

**Special Note:** When register pair is SP, also sets stack start pointer in UI.

---

### MOV - Move Register to Register

**Base:** 350 | **Syntax:** `MOV dest, src`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 350 | 0 | Missing operands | op1 or op2 is empty (need both source and destination) |
| 351 | 1 | Invalid register(s) | Source or destination register not recognized |

**Returns:** 0 on success

---

### MVI - Move Immediate to Register

**Base:** 360 | **Syntax:** `MVI reg, imm8`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 360 | 0 | Missing operands | op1 or op2 is empty |
| 361 | 1 | Invalid hex value | Immediate value is not valid hex |
| 362 | 2 | Invalid register | Register name not recognized |

**Returns:** 0 on success

---

### ORA - Logical OR Register with Accumulator

**Base:** 380 | **Syntax:** `ORA reg`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 380 | 0 | Missing operand | op1 is empty |

**Returns:** 0 on success

---

### ORI - Logical OR Immediate with Accumulator

**Base:** 390 | **Syntax:** `ORI imm8`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 390 | 0 | Missing operand | op1 is empty |
| 391 | 1 | Invalid hex immediate | Value is not valid hex or label resolves to invalid/out-of-range |

**Returns:** 0 on success

---

### POP - Pop Register from Stack

**Base:** 420 | **Syntax:** `POP reg`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 420 | 0 | Missing operand | op1 is empty |
| 421 | 1 | Invalid register | Register name not recognized |

**Returns:** 0 on success

---

### PUSH - Push Register onto Stack

**Base:** 430 | **Syntax:** `PUSH reg`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 430 | 0 | Missing operand | op1 is empty |
| 431 | 1 | Invalid register | Register name not recognized |

**Returns:** 0 on success

---

### RET - Return from Subroutine

**Base:** 460 | **Syntax:** `RET`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 460 | 0 | No address on stack | Stack is empty (stack underflow) |
| 461 | 1 | Invalid return frame | Stack top is not a CALL frame |

**Returns:** 0 on success

---

### SUB - Subtract Register from Accumulator

**Base:** 580 | **Syntax:** `SUB reg`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 580 | 0 | Missing operand | op1 is empty |
| 581 | 1 | Invalid register | Register name not recognized |

**Returns:** 0 on success

---

### SUI - Subtract Immediate from Accumulator

**Base:** 590 | **Syntax:** `SUI imm8`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 590 | 0 | Missing operand | op1 is empty |
| 591 | 1 | Invalid hex immediate | Value is not valid hex |

**Returns:** 0 on success

---

### SBB - Subtract Register from Accumulator with Borrow

**Base:** 570 | **Syntax:** `SBB reg`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 570 | 0 | Missing operand | op1 is empty |
| 571 | 1 | Invalid register | Register name not recognized |

**Returns:** 0 on success

---

### XRA - Logical XOR Register with Accumulator

**Base:** 400 | **Syntax:** `XRA reg`

| Code | Offset | Error | Description |
|------|--------|-------|-------------|
| 400 | 0 | Missing operand | op1 is empty |
| 401 | 1 | Invalid register | Register name not recognized |

**Returns:** 0 on success

---

## Error Code Patterns

### Pattern 1: Single-Operand Instructions (e.g., ADD, INR, DCR)

```
Base+0: Missing operand
Base+1: Invalid register
```

Examples: ADD, ADI, ANA, ANI, CMP, CPI, DCR, INR, ORA, ORI, SUB, SUI, XRA

### Pattern 2: Two-Operand Instructions (e.g., MOV, MVI)

```
Base+0: Missing operands
Base+1: Invalid register(s)
Base+2: Additional validation (optional)
```

Examples: MOV, MVI, LXI

### Pattern 3: Label-Based Instructions (e.g., JMP, CALL)

```
Base+0: Missing operand
Base+1: Label not found
Base+2: Invalid address format
Base+3: Out of range
```

Examples: JMP, CALL

### Pattern 4: Stack Operations (e.g., PUSH, POP, RET)

```
Base+0: Missing operand / Empty stack
Base+1: Invalid register / Invalid frame
```

Examples: POP, PUSH, RET

---

## Using SetError() in Your Code

The `SetError()` function logs error information for debugging:

```vba
' In your instruction function:
SetError errorBase + offset, "Descriptive message"

' Example:
SetError 340, "LXI: Missing Operands: B "
SetError 341, "LXI: Invalid Register: X"
```

This creates an audit trail of all errors encountered during execution.

---

## Trace Output Column 19: Error Codes

When tracing is enabled, column 19 (Err) records the return value of each instruction:

```
Step | PC   | Op  | OP1 | OP2  | A  | B  | ... | Err | MemAddr | ...
-----|------|-----|-----|------|----|----|-----|-----|---------|-----
  42 | 1234 | LXI | B   |      |    |    | ... | 340 |         |
  43 | 1235 | MOV | A   | C    |    |    | ... | 0   |         |
  44 | 1236 | JMP | XXXX|      |    |    | ... | 311 |         |
```

**Interpretation:**
- **Row 42:** LXI error 340 = missing operand
- **Row 43:** MOV error 0 = success
- **Row 44:** JMP error 311 = label not found

---

## Debugging with Error Codes

### Step 1: Find the Error in Trace
```
Filter Trace sheet: Err <> 0
```

### Step 2: Identify the Instruction
```
Look at Op column for that row
Example: LXI
```

### Step 3: Find the Error Base
```
Look up LXI in Quick Reference Table
ErrorBase = 340
```

### Step 4: Calculate Offset
```
Error Code - ErrorBase = Offset
342 - 340 = 2 (offset +2)
```

### Step 5: Look Up Description
```
LXI offset +2 = "Invalid address"
Likely cause: Address in OP2 is not valid hex or label doesn't exist
```

### Step 6: Check Trace Details
```
Look at OP1 and OP2 columns for that row
Verify: Is register pair valid? Is address/label valid?
```

---

## Best Practices

### 1. Always Return 0 on Success

```vba
' At the end of every instruction function:
InstructionName = 0
```

### 2. Check Operands First

```vba
' Order of checks:
' 1. Empty operands
' 2. Invalid types (register names, hex format)
' 3. Invalid ranges (out of bounds, unresolved labels)
' 4. Instruction-specific validations
```

### 3. Exit Early on Error

```vba
If condition Then
    SetError errorBase + N, "Description"
    InstructionName = errorBase + N
    Exit Function
End If
```

### 4. Use Consistent Error Patterns

Follow the patterns above so users can predict error codes across instructions.

---

## Summary Table

| Return Value | Meaning |
|--------------|---------|
| 0 | ✓ Success |
| Negative (e.g., -100) | Special signal (e.g., HLT) |
| 100-199 | ACI/ADC/ADD error |
| 200-299 | ADI/ANA/ANI/CMP/CPI/DCR/INR error |
| 300-399 | EQU/JMP/LXI/MOV/MVI error |
| 400-499 | ORA/ORI/POP/PUSH/RET/XRA error |
| 500-599 | SUB/SUI/SBB error |

---

**Document Version:** 2.0  
**Last Updated:** March 7, 2026  
**Status:** Complete extraction from VBA source  
**For Use With:** Z80_Model_Current_2026-03-06_2239.xlsm

