# Z80 Prefix Instructions: How They Work

## The Problem CB Prefix Solves

The Intel 8080 has **no bit manipulation instructions**. You can't directly test, set, or clear individual bits in a register or memory location.

The Z80 adds bit operations: **BIT, RES, SET** using a special encoding mechanism called **prefixes**.

## What is a Prefix?

A **prefix** is a special byte that modifies how the NEXT byte is interpreted.

- **Prefix byte**: A special code (CB, ED, DD, FD) that says "interpret the next byte differently"
- **Following byte**: The actual instruction encoding

### Example: BIT 0, A (Test bit 0 in register A)

In Z80 machine code: **CB 47**

Breaking this down:
```
CB  = Prefix byte (says: "interpret next byte as a bit operation")
47  = Base opcode (encodes which bit operation, which bit, which register)
```

**Total: 2 bytes**

## How CB Prefix Bit Operations Work

### The CB Prefix Group

Instructions with CB prefix handle bit manipulation on ANY register or memory location:
- BIT n, reg - Test bit n in register
- RES n, reg - Reset (clear) bit n in register  
- SET n, reg - Set bit n in register

Where:
- **n** = bit position (0-7)
- **reg** = A, B, C, D, E, H, L, or (HL)

### The Base Opcode Encoding

After the CB prefix, the next byte encodes:
1. **Which operation** (BIT, RES, SET) - uses upper 2 bits
2. **Which bit** (0-7) - uses middle 3 bits
3. **Which register** - uses lower 3 bits

**Byte format: [op_bits][bit_bits][reg_bits]** or in binary: **OOBBBRR** where:
- OO = operation (00=BIT, 01=RES, 10=SET)
- BBB = bit position (000=bit 0, 001=bit 1, ..., 111=bit 7)
- RR = register (000=B, 001=C, 010=D, 011=E, 100=H, 101=L, 110=(HL), 111=A)

### Example: BIT 0, A

```
CB          Prefix (tells Z80: "next byte is a bit operation")
47          Base opcode

Breaking down 47 in binary: 0100 0111
            op bits: 01    (BIT operation)
            bit #:   000   (bit 0)
            reg:     111   (register A)

So: BIT (01) bit 0 (000) from A (111) = 01 000 111 = 0x47 ✓
```

### Example: SET 5, H

```
CB          Prefix
FA          Base opcode (11111010)

Breaking down FA in binary: 1111 1010
            op bits: 11    (SET operation)
            bit #:   101   (bit 5)
            reg:     010   (register H)

So: SET (11) bit 5 (101) in H (010) = 11 101 010 = 0xFA ✓
```

### Example: RES 2, (HL)

```
CB          Prefix
96          Base opcode (10010110)

Breaking down 96 in binary: 1001 0110
            op bits: 10    (RES operation)
            bit #:   010   (bit 2)
            reg:     110   ((HL) memory)

So: RES (10) bit 2 (010) from (HL) (110) = 10 010 110 = 0x96 ✓
```

## Other Prefix Examples

### ED Prefix - Extended Operations

**ED** prefix encodes extended operations like:
- Block operations: LDIR (ED B0), LDDR (ED B8), CPIR (ED B1), CPDR (ED B9)
- I/O: IN, OUT
- Advanced 16-bit arithmetic: ADC HL, SBC HL

Example: **LDIR** (Load, Increment, Repeat)
```
ED          Prefix (extended instruction)
B0          Specific operation code for LDIR
```

**Total: 2 bytes**

### DD Prefix - IX Register Operations

**DD** prefix means "use IX register instead of HL" for the following instruction.

Example: **LD A, (IX+5)** (Load A from IX plus offset 5)
```
DD          Prefix (use IX)
7E          Base opcode for LD A, (HL) - but with IX substitution
05          Signed offset byte
```

**Total: 3 bytes** (because offset is an additional byte)

### FD Prefix - IY Register Operations

**FD** prefix means "use IY register instead of HL" for the following instruction.

Example: **LD (IY-10), A** (Store A at IY minus offset 10)
```
FD          Prefix (use IY)
77          Base opcode for LD (HL), A - but with IY substitution
F6          Signed offset byte (-10 in two's complement = 0xF6)
```

**Total: 3 bytes**

## How the Compiler Handles Prefixes

### In the Opcode Table

Each prefixed instruction has two entries:
1. **Prefix byte** - the prefix code (CB, ED, DD, FD)
2. **Base opcode** - what follows the prefix

Example from your Z80 opcodes table:
```
Mnemonic  OP1  OP2     Hex     Bytes  Prefix  Base_Opcode
BIT       0    A       47      2      CB      47
BIT       0    (HL)    46      2      CB      46
SET       5    H       FA      2      CB      FA
LDIR                   B0      2      ED      B0
LD        A    (IX+5)  7E      3      DD      7E
```

### How the Encoder Works

When the compiler sees: **BIT 0, A**

1. **Lookup**: Find "BIT|0|A" in opcode table
2. **Get info**: Prefix=CB, Base_Opcode=47, Bytes=2
3. **Encode**:
   ```
   Output byte 1: 0xCB (the prefix)
   Output byte 2: 0x47 (the base opcode)
   ```

When the compiler sees: **LD A, (IX+5)**

1. **Lookup**: Find "LD|A|(IX+5)" in opcode table
2. **Get info**: Prefix=DD, Base_Opcode=7E, Bytes=3
3. **Encode**:
   ```
   Output byte 1: 0xDD (the prefix for IX)
   Output byte 2: 0x7E (the base opcode for LD A)
   Output byte 3: 0x05 (the offset +5)
   ```

## Summary

| Prefix | Purpose | Example | Machine Code |
|--------|---------|---------|--------------|
| **CB** | Bit operations | BIT 0, A | CB 47 |
| **ED** | Extended ops | LDIR | ED B0 |
| **DD** | IX register | LD A,(IX+5) | DD 7E 05 |
| **FD** | IY register | LD (IY-10),A | FD 77 F6 |

The key insight: **The prefix changes how the next byte is interpreted**. Without the prefix, 0x47 means "MOV A,A" (8080 instruction). With the CB prefix, 0x47 means "BIT 0,A" (Z80 bit operation).

