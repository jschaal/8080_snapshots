# Using ALU Helpers for All Add/Subtract Variants

**Date:** March 7, 2026  
**Purpose:** Show exactly how the ALU helpers handle all add/subtract instruction variants  

---

## Quick Answer: YES!

The ALU helper functions are **designed to work with all variants**. Here's how the `includeCarry` and `includeBorrow` parameters work:

```vba
' Addition family:
HlpALU_Addition(operand, False)  ' ADD, ADI         (no carry)
HlpALU_Addition(operand, True)   ' ADC, ACI         (with carry)

' Subtraction family:
HlpALU_Subtraction(operand, False)  ' SUB_, SUI    (no borrow)
HlpALU_Subtraction(operand, True)   ' SBB, SBI     (with borrow)
```

---

## The Logic

### Addition Helper Parameter Meaning

```vba
HlpALU_Addition(operandValue, includeCarry)
                              ↑
                              └─→ False = Don't add carry flag
                                  True  = Add carry flag to result
```

**When `includeCarry = False` (ADD, ADI):**
```
A ← A + operand
```

**When `includeCarry = True` (ADC, ACI):**
```
A ← A + operand + CY
```

---

### Subtraction Helper Parameter Meaning

```vba
HlpALU_Subtraction(operandValue, includeBorrow)
                                 ↑
                                 └─→ False = Don't include borrow
                                     True  = Include borrow in result
```

**When `includeBorrow = False` (SUB_, SUI):**
```
A ← A - operand
```

**When `includeBorrow = True` (SBB, SBI):**
```
A ← A - operand - CY
```

---

## Complete Examples: All Add/Subtract Variants

### 1. ADD (Add Register to A)
```vba
Public Function ADD(ByVal op1 As String) As Long
    Dim errorBase As Long: errorBase = 120
    Dim src As Variant

    src = HlpResolveReg8(op1, errorBase, "ADD")
    If IsEmpty(src) Then ADD = errorBase: Exit Function

    HlpALU_Addition CLng(src), False  ' ← No carry included
    ADD = 0
End Function
```

---

### 2. ADI (Add Immediate to A)
```vba
Public Function ADI(ByVal op1 As String) As Long
    Dim errorBase As Long: errorBase = 130
    Dim imm8 As Variant

    imm8 = HlpResolveImm8(op1, errorBase, "ADI")
    If IsEmpty(imm8) Then ADI = errorBase: Exit Function

    HlpALU_Addition CLng(imm8), False  ' ← No carry included
    ADI = 0
End Function
```

---

### 3. ADC (Add Register with Carry to A)
```vba
Public Function ADC(ByVal op1 As String) As Long
    Dim errorBase As Long: errorBase = 110
    Dim src As Variant

    src = HlpResolveReg8(op1, errorBase, "ADC")
    If IsEmpty(src) Then ADC = errorBase: Exit Function

    HlpALU_Addition CLng(src), True   ' ← WITH carry included!
    ADC = 0
End Function
```

**Key Difference:** Notice the `True` parameter — this makes ADC include the carry flag in the calculation.

---

### 4. ACI (Add Immediate with Carry to A)
```vba
Public Function ACI(ByVal op1 As String) As Long
    Dim errorBase As Long: errorBase = 100
    Dim imm8 As Variant

    imm8 = HlpResolveImm8(op1, errorBase, "ACI")
    If IsEmpty(imm8) Then ACI = errorBase: Exit Function

    HlpALU_Addition CLng(imm8), True  ' ← WITH carry included!
    ACI = 0
End Function
```

**Key Difference:** Same as ADC but with immediate value instead of register.

---

### 5. SUB_ (Subtract Register from A)
```vba
Public Function SUB_(ByVal op1 As String) As Long
    Dim errorBase As Long: errorBase = 580
    Dim src As Variant

    src = HlpResolveReg8(op1, errorBase, "SUB")
    If IsEmpty(src) Then SUB_ = errorBase: Exit Function

    HlpALU_Subtraction CLng(src), False  ' ← No borrow included
    SUB_ = 0
End Function
```

---

### 6. SUI (Subtract Immediate from A)
```vba
Public Function SUI(ByVal op1 As String) As Long
    Dim errorBase As Long: errorBase = 590
    Dim imm8 As Variant

    imm8 = HlpResolveImm8(op1, errorBase, "SUI")
    If IsEmpty(imm8) Then SUI = errorBase: Exit Function

    HlpALU_Subtraction CLng(imm8), False  ' ← No borrow included
    SUI = 0
End Function
```

---

### 7. SBB (Subtract Register with Borrow from A)
```vba
Public Function SBB(ByVal op1 As String) As Long
    Dim errorBase As Long: errorBase = 570
    Dim src As Variant

    src = HlpResolveReg8(op1, errorBase, "SBB")
    If IsEmpty(src) Then SBB = errorBase: Exit Function

    HlpALU_Subtraction CLng(src), True   ' ← WITH borrow included!
    SBB = 0
End Function
```

**Key Difference:** Notice the `True` parameter — this makes SBB include the carry flag (borrow) in the calculation.

---

### 8. SBI (Subtract Immediate with Borrow from A)
```vba
Public Function SBI(ByVal op1 As String) As Long
    Dim errorBase As Long: errorBase = 560
    Dim imm8 As Variant

    imm8 = HlpResolveImm8(op1, errorBase, "SBI")
    If IsEmpty(imm8) Then SBI = errorBase: Exit Function

    HlpALU_Subtraction CLng(imm8), True  ' ← WITH borrow included!
    SBI = 0
End Function
```

**Key Difference:** Same as SBB but with immediate value instead of register.

---

## Pattern Summary

### Addition Instructions (4 total)

| Instruction | Register? | With Carry? | Helper Call |
|-------------|-----------|-------------|------------|
| ADD | ✅ Yes | ❌ No | `HlpALU_Addition(src, False)` |
| ADI | ❌ Imm | ❌ No | `HlpALU_Addition(imm, False)` |
| ADC | ✅ Yes | ✅ Yes | `HlpALU_Addition(src, True)` |
| ACI | ❌ Imm | ✅ Yes | `HlpALU_Addition(imm, True)` |

### Subtraction Instructions (4 total)

| Instruction | Register? | With Borrow? | Helper Call |
|-------------|-----------|--------------|------------|
| SUB_ | ✅ Yes | ❌ No | `HlpALU_Subtraction(src, False)` |
| SUI | ❌ Imm | ❌ No | `HlpALU_Subtraction(imm, False)` |
| SBB | ✅ Yes | ✅ Yes | `HlpALU_Subtraction(src, True)` |
| SBI | ❌ Imm | ✅ Yes | `HlpALU_Subtraction(imm, True)` |

---

## How the Helper Functions Handle Carry/Borrow Internally

### Inside HlpALU_Addition()

```vba
Private Sub HlpALU_Addition(ByVal operandValue As Long, Optional ByVal includeCarry As Boolean = False)
    Dim a As Long, b As Long, sum As Long, result As Long
    Dim ac As Long, cy As Long
    
    a = pRegs("A") And 255
    b = operandValue And 255
    
    ' ← THIS IS THE KEY LOGIC:
    ' If includeCarry is True, add the current carry flag to operand
    If includeCarry Then
        b = b + pFlags("Carry")  ' b becomes (operand + old carry)
    End If
    
    ' Now sum = A + B (where B might include the carry)
    sum = a + b
    result = sum And 255
    
    ' Calculate and set all flags...
    cy = IIf(sum > 255, 1, 0)
    ' ... rest of flag calculation
End Sub
```

**Example Execution for ADC B (when A=0x50, B=0x40, CY=1):**
```
1. a = 0x50
2. b = 0x40
3. includeCarry = True, so: b = 0x40 + 1 = 0x41
4. sum = 0x50 + 0x41 = 0x91
5. result = 0x91
6. CY = 0 (sum did not overflow)
7. All flags calculated
8. A now contains 0x91
```

### Inside HlpALU_Subtraction()

```vba
Private Sub HlpALU_Subtraction(ByVal operandValue As Long, Optional ByVal includeBorrow As Boolean = False)
    Dim a As Long, b As Long, diff As Long, result As Long
    Dim borrow As Long, acBorrow As Long
    
    a = pRegs("A") And 255
    b = operandValue And 255
    
    ' ← THIS IS THE KEY LOGIC:
    ' If includeBorrow is True, add the current carry flag (borrow) to operand
    If includeBorrow Then
        b = b + pFlags("Carry")  ' b becomes (operand + old borrow)
    End If
    
    ' Calculate borrow and perform subtraction
    borrow = IIf(a < b, 1, 0)
    ' ... rest of subtraction logic
End Sub
```

**Example Execution for SBB B (when A=0x50, B=0x20, CY=1):**
```
1. a = 0x50
2. b = 0x20
3. includeBorrow = True, so: b = 0x20 + 1 = 0x21
4. borrow = IIf(0x50 < 0x21, 1, 0) = 0 (no borrow needed)
5. diff = 0x50 - 0x21 = 0x2F
6. result = 0x2F
7. CY = 0 (no borrow)
8. All flags calculated
9. A now contains 0x2F
```

---

## Real-World Usage Example: Multi-Byte Addition

Here's why you need both ADD and ADC:

```vba
' Add two 16-bit numbers: BC + DE → HL
' 
' BC = 0x1234, DE = 0x5678
' Expected: HL = 0x68AC

Private Sub AddTwoWords()
    ' First add low bytes (C + E)
    RegSet "A", pRegs("C")              ' A = 0x34
    HlpALU_Addition pRegs("E"), False   ' ADD E
    RegSet "L", pRegs("A")              ' L = result (0xAC), CY = 0
    
    ' Then add high bytes with carry (B + D + CY)
    RegSet "A", pRegs("B")              ' A = 0x12
    HlpALU_Addition pRegs("D"), True    ' ADC D (includes carry from low byte!)
    RegSet "H", pRegs("A")              ' H = result (0x68)
    
    ' Now HL = 0x68AC ✓
End Sub
```

Similarly for multi-byte subtraction, you'd use SUB_ for the first byte and SBB for subsequent bytes.

---

## Complete Implementation Checklist

### ✅ Using Helpers for ADD & ADC
- [ ] ADD uses `HlpALU_Addition(src, False)`
- [ ] ADI uses `HlpALU_Addition(imm, False)`
- [ ] ADC uses `HlpALU_Addition(src, True)`
- [ ] ACI uses `HlpALU_Addition(imm, True)`

### ✅ Using Helpers for SUB_ & SBB
- [ ] SUB_ uses `HlpALU_Subtraction(src, False)`
- [ ] SUI uses `HlpALU_Subtraction(imm, False)`
- [ ] SBB uses `HlpALU_Subtraction(src, True)`
- [ ] SBI uses `HlpALU_Subtraction(imm, True)`

### ✅ Testing Multi-Byte Operations
- [ ] Test ADD + ADC sequence
- [ ] Test SUB_ + SBB sequence
- [ ] Verify carry/borrow flag is properly propagated
- [ ] Test boundary cases (0xFF + 0x01, 0x00 - 0x01)

---

## The Beauty of This Design

One helper function handles **4 instruction variants**:

```vba
' These 4 instructions all use the SAME helper:
HlpALU_Addition(operand, False)  ← ADD, ADI
HlpALU_Addition(operand, True)   ← ADC, ACI

' These 4 instructions all use the SAME helper:
HlpALU_Subtraction(operand, False)  ← SUB_, SUI
HlpALU_Subtraction(operand, True)   ← SBB, SBI
```

**Benefits:**
- ✅ **Single source of truth** for arithmetic logic
- ✅ **Eliminates duplication** — no copy/paste errors
- ✅ **Easier to debug** — fix one place, all 4 instructions work correctly
- ✅ **Consistent behavior** across register and immediate variants
- ✅ **Carry/Borrow properly propagated** for multi-byte operations

---

## Yes, ACI and SUI Work Perfectly Too!

Here's the final proof — your complete refactored instructions:

```vba
' ADD FAMILY (all 4 use the same HlpALU_Addition helper)

Public Function ADD(ByVal op1 As String) As Long
    Dim src As Variant
    src = HlpResolveReg8(op1, 120, "ADD")
    If IsEmpty(src) Then ADD = 120: Exit Function
    HlpALU_Addition CLng(src), False
    ADD = 0
End Function

Public Function ADI(ByVal op1 As String) As Long
    Dim imm8 As Variant
    imm8 = HlpResolveImm8(op1, 130, "ADI")
    If IsEmpty(imm8) Then ADI = 130: Exit Function
    HlpALU_Addition CLng(imm8), False
    ADI = 0
End Function

Public Function ADC(ByVal op1 As String) As Long
    Dim src As Variant
    src = HlpResolveReg8(op1, 110, "ADC")
    If IsEmpty(src) Then ADC = 110: Exit Function
    HlpALU_Addition CLng(src), True  ' ← WITH CARRY
    ADC = 0
End Function

Public Function ACI(ByVal op1 As String) As Long
    Dim imm8 As Variant
    imm8 = HlpResolveImm8(op1, 100, "ACI")
    If IsEmpty(imm8) Then ACI = 100: Exit Function
    HlpALU_Addition CLng(imm8), True  ' ← WITH CARRY
    ACI = 0
End Function

' SUB FAMILY (all 4 use the same HlpALU_Subtraction helper)

Public Function SUB_(ByVal op1 As String) As Long
    Dim src As Variant
    src = HlpResolveReg8(op1, 580, "SUB")
    If IsEmpty(src) Then SUB_ = 580: Exit Function
    HlpALU_Subtraction CLng(src), False
    SUB_ = 0
End Function

Public Function SUI(ByVal op1 As String) As Long
    Dim imm8 As Variant
    imm8 = HlpResolveImm8(op1, 590, "SUI")
    If IsEmpty(imm8) Then SUI = 590: Exit Function
    HlpALU_Subtraction CLng(imm8), False
    SUI = 0
End Function

Public Function SBB(ByVal op1 As String) As Long
    Dim src As Variant
    src = HlpResolveReg8(op1, 570, "SBB")
    If IsEmpty(src) Then SBB = 570: Exit Function
    HlpALU_Subtraction CLng(src), True  ' ← WITH BORROW
    SBB = 0
End Function

Public Function SBI(ByVal op1 As String) As Long
    Dim imm8 As Variant
    imm8 = HlpResolveImm8(op1, 560, "SBI")
    If IsEmpty(imm8) Then SBI = 560: Exit Function
    HlpALU_Subtraction CLng(imm8), True  ' ← WITH BORROW
    SBI = 0
End Function
```

**Look at that!** 8 complete instruction implementations, each only 6-7 lines, all using just 2 helper functions. This is **excellent code reuse**. ✅

---

## Summary

**Q: Can I use helpers for ACI, SUI, SBB?**  
**A: YES! 100%**

The helpers are designed specifically for this. The boolean parameters tell the helper whether to include carry/borrow:
- `False` = basic operation (ADD, ADI, SUB_, SUI)
- `True` = with carry/borrow (ADC, ACI, SBB, SBI)

Same helper, different parameter = different behavior. Perfect!

---

**Document Version:** 1.0  
**Last Updated:** March 7, 2026

