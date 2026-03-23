Option Explicit
'================================================================================
' Class:        cls6510CPU
' Purpose:      CPU core for the Excel/VBA MOS 6510 emulator.
'               Mirrors the architecture of clsDecCPU exactly:
'               - Register & flag storage (Dictionary-based, same pattern)
'               - Stack helper (reuses clsDecStack unchanged)
'               - Label helper (reuses clsLabels unchanged)
'               - Memory helper (reuses clsMemory unchanged)
'               - All 56 6510 mnemonics implemented as Public Functions
'               - RunOpcode dispatcher (drop-in compatible with decExecute)
'               - UI deferral (RegistersDirty / FlagsDirty / StackDirty)
'               - HlpFlagPack / HlpFlagUnpack for Trace sheet SR column
'
' 6510 Registers (vs 8080):
'   A            - 8-bit accumulator          (same)
'   X, Y         - 8-bit index registers      (replaces B,C,D,E,H,L pairs)
'   PC           - 16-bit program counter     (same)
'   SP           - 8-bit stack pointer        (fixed page $01: $0100-$01FF)
'   SR / P       - 8-bit status register      (replaces pFlags dictionary,
'                                              but we keep the same Dict pattern)
'
' 6510 Flags (SR bit layout):  N V - B D I Z C
'   Bit 7  N  Negative   (replaces Sign)
'   Bit 6  V  oVerflow   (new; no 8080 equivalent)
'   Bit 5  -  Unused     (always 1 in hardware; we ignore)
'   Bit 4  B  Break      (set by BRK; software interrupt marker)
'   Bit 3  D  Decimal    (BCD mode; ADC/SBC behave differently when D=1)
'   Bit 2  I  Interrupt  (IRQ disable)
'   Bit 1  Z  Zero       (same)
'   Bit 0  C  Carry      (same; NOTE: borrow sense is INVERTED vs 8080)
'
' Key differences from 8080 you need to keep in mind:
'   - No register pairs (no BC/DE/HL).  Zero-page ($00-$FF) is used for
'     16-bit pointer storage via indirect addressing modes.
'   - Carry flag for SBC is inverted: C=1 means NO borrow (set CLC before SBC).
'   - All branches are relative (-128..+127 from next instruction).
'   - Stack is ALWAYS page $01 ($0100-$01FF).  SP wraps in that page.
'   - JSR pushes PC-1 (address of last byte of JSR); RTS pops and adds 1.
'   - BRK pushes PC+1 (one past the padding byte) then SR with B=1.
'
' Addressing Modes encoded in op1/op2 columns (mirrors your existing model):
'   IMM      #nn          op1 = hex literal e.g. "4A"
'   ZP       nn           op1 = 1-2 hex digit address e.g. "42"
'   ZP,X     nn,X         op1 = "42,X"
'   ZP,Y     nn,Y         op1 = "42,Y"
'   ABS      nnnn         op1 = 4-digit hex address e.g. "C000"
'   ABS,X    nnnn,X       op1 = "C000,X"
'   ABS,Y    nnnn,Y       op1 = "C000,Y"
'   IND      (nnnn)       op1 = "(C000)"  (JMP only)
'   (ZP,X)   (nn,X)       op1 = "(42,X)"
'   (ZP),Y   (nn),Y       op1 = "(42),Y"
'   IMP/ACC  (no operand) op1 = "" or "A"
'
' Notes:
'   - clsDecStack, clsLabels, clsMemory, clsAddrList, clsAddrRecord are
'     shared unchanged with the 8080/Z80 model.
'   - modGlobals: add  Public gDecCPU6510 As cls6510CPU  (same singleton
'     pattern as gDecCPU).
'   - decExecute: add a SelectEngine branch that calls Execute6510 instead
'     of Execute when the CPU sheet has a "6510" radio/cell selected.
'   - Trace sheet columns are identical; SR replaces the 8080 F register.
'
' Source:       Initial build (2026-03-22) - based on clsDecCPU v2.08
'================================================================================

'--- State ---
Private pRegs   As Object          ' Dictionary: A, X, Y, PC, SP (decimal Long)
Private pFlags  As Object          ' Dictionary: N, V, B, D, I, Z, C (0/1 ints)
Private pStack  As New clsDecStack
Private pLabels As New clsLabels
Private pMemory As clsMemory
Private pAddressList As clsAddrList

'--- UI deferral (headless runs) ---
Private mUIRegsDirty  As Boolean
Private mUIFlagsDirty As Boolean
Private mUIStackDirty As Boolean

'--- Error state (mirrors clsDecCPU.SetError / Err pattern) ---
Private pLastErrCode As Long
Private pLastErrMsg  As String

'================================================================================
' Class_Initialize
'================================================================================
Private Sub Class_Initialize()
    Set pRegs = CreateObject("Scripting.Dictionary")
    pRegs.Add "A",  0
    pRegs.Add "X",  0
    pRegs.Add "Y",  0
    pRegs.Add "PC", 0
    pRegs.Add "SP", &HFF         ' points to top of stack page; wraps in $01xx

    Set pFlags = CreateObject("Scripting.Dictionary")
    pFlags.Add "N", 0            ' Negative
    pFlags.Add "V", 0            ' oVerflow
    pFlags.Add "B", 0            ' Break
    pFlags.Add "D", 0            ' Decimal
    pFlags.Add "I", 1            ' Interrupt disable (set on reset)
    pFlags.Add "Z", 0            ' Zero
    pFlags.Add "C", 0            ' Carry

    Set pAddressList = gAddressList
    Set pMemory      = gMemory
End Sub
' --- Class_Initialize ---

'================================================================================
' RunOpcode
' Drop-in replacement for clsDecCPU.RunOpcode.
' Called by decExecute (or a new Execute6510 module) with the same signature.
' Returns 0 on success; ERR_EXEC_END / ERR_BAD_OPCODE / negative on halt/error.
'================================================================================
Public Function RunOpcode(ByVal opcode As String, _
    ByVal op1 As String, ByVal op2 As String, _
    Optional ByVal label As String = "") As Long

    Dim result As Long
    result = 0

    Select Case opcode
        '----------------------------------------------------------------------
        ' Load / Store
        '----------------------------------------------------------------------
        Case "LDA": result = Me.LDA(op1)
        Case "LDX": result = Me.LDX(op1)
        Case "LDY": result = Me.LDY(op1)
        Case "STA": result = Me.STA(op1)
        Case "STX": result = Me.STX(op1)
        Case "STY": result = Me.STY(op1)

        '----------------------------------------------------------------------
        ' Register transfers
        '----------------------------------------------------------------------
        Case "TAX": result = Me.TAX()
        Case "TAY": result = Me.TAY()
        Case "TXA": result = Me.TXA()
        Case "TYA": result = Me.TYA()
        Case "TSX": result = Me.TSX()
        Case "TXS": result = Me.TXS()

        '----------------------------------------------------------------------
        ' Stack
        '----------------------------------------------------------------------
        Case "PHA": result = Me.PHA()
        Case "PHP": result = Me.PHP()
        Case "PLA": result = Me.PLA()
        Case "PLP": result = Me.PLP()

        '----------------------------------------------------------------------
        ' Arithmetic
        '----------------------------------------------------------------------
        Case "ADC": result = Me.ADC(op1)
        Case "SBC": result = Me.SBC(op1)
        Case "INC": result = Me.INC(op1)
        Case "INX": result = Me.INX()
        Case "INY": result = Me.INY()
        Case "DEC": result = Me.DEC(op1)
        Case "DEX": result = Me.DEX()
        Case "DEY": result = Me.DEY()

        '----------------------------------------------------------------------
        ' Logical
        '----------------------------------------------------------------------
        Case "AND": result = Me.AND_(op1)
        Case "ORA": result = Me.ORA(op1)
        Case "EOR": result = Me.EOR(op1)
        Case "BIT": result = Me.BIT(op1)

        '----------------------------------------------------------------------
        ' Shifts & Rotates
        '----------------------------------------------------------------------
        Case "ASL": result = Me.ASL(op1)
        Case "LSR": result = Me.LSR(op1)
        Case "ROL": result = Me.ROL(op1)
        Case "ROR": result = Me.ROR(op1)

        '----------------------------------------------------------------------
        ' Compare
        '----------------------------------------------------------------------
        Case "CMP": result = Me.CMP(op1)
        Case "CPX": result = Me.CPX(op1)
        Case "CPY": result = Me.CPY(op1)

        '----------------------------------------------------------------------
        ' Branches (all relative; op1 = target label or hex offset)
        '----------------------------------------------------------------------
        Case "BCC": result = Me.BCC(op1)
        Case "BCS": result = Me.BCS(op1)
        Case "BEQ": result = Me.BEQ(op1)
        Case "BMI": result = Me.BMI(op1)
        Case "BNE": result = Me.BNE(op1)
        Case "BPL": result = Me.BPL(op1)
        Case "BVC": result = Me.BVC(op1)
        Case "BVS": result = Me.BVS(op1)

        '----------------------------------------------------------------------
        ' Jumps & Subroutines
        '----------------------------------------------------------------------
        Case "JMP": result = Me.JMP(op1)
        Case "JSR": result = Me.JSR(op1)
        Case "RTS": result = Me.RTS()
        Case "RTI": result = Me.RTI()
        Case "BRK": result = Me.BRK()

        '----------------------------------------------------------------------
        ' Flag operations
        '----------------------------------------------------------------------
        Case "CLC": result = Me.CLC()
        Case "CLD": result = Me.CLD()
        Case "CLI": result = Me.CLI()
        Case "CLV": result = Me.CLV()
        Case "SEC": result = Me.SEC()
        Case "SED": result = Me.SED()
        Case "SEI": result = Me.SEI()

        '----------------------------------------------------------------------
        ' Misc / pseudo-ops
        '----------------------------------------------------------------------
        Case "NOP":             result = Me.NOP()
        Case "ORG", "DB", "DS": result = 0          ' assembler directives
        Case "EQU":             result = Me.EQU_(op1, label)
        Case "END":             result = ERR_EXEC_END
        Case "HLT", "STP":     result = Me.HLT()   ' unofficial/WDC stop

        Case Else
            SetError ERR_BAD_OPCODE, "6510 RunOpcode: Unknown opcode: " & opcode
            result = ERR_BAD_OPCODE
    End Select

    RunOpcode = result
End Function
' --- RunOpcode ---

'================================================================================
' LOAD / STORE
'================================================================================

' ------------------------------------------------------------------------------
' LDA  Load Accumulator
' Modes: IMM, ZP, ZP,X, ABS, ABS,X, ABS,Y, (ZP,X), (ZP),Y
' Flags: N, Z
' ------------------------------------------------------------------------------
Public Function LDA(ByVal op1 As String) As Long
    Dim v As Long
    v = HlpReadOperand(op1, 700, "LDA")
    If pLastErrCode <> 0 Then LDA = pLastErrCode: Exit Function
    SetReg "A", v
    HlpFlagNZ v
    LDA = 0
End Function

' ------------------------------------------------------------------------------
' LDX  Load X Register
' Modes: IMM, ZP, ZP,Y, ABS, ABS,Y
' Flags: N, Z
' ------------------------------------------------------------------------------
Public Function LDX(ByVal op1 As String) As Long
    Dim v As Long
    v = HlpReadOperand(op1, 710, "LDX")
    If pLastErrCode <> 0 Then LDX = pLastErrCode: Exit Function
    SetReg "X", v
    HlpFlagNZ v
    LDX = 0
End Function

' ------------------------------------------------------------------------------
' LDY  Load Y Register
' Modes: IMM, ZP, ZP,X, ABS, ABS,X
' Flags: N, Z
' ------------------------------------------------------------------------------
Public Function LDY(ByVal op1 As String) As Long
    Dim v As Long
    v = HlpReadOperand(op1, 720, "LDY")
    If pLastErrCode <> 0 Then LDY = pLastErrCode: Exit Function
    SetReg "Y", v
    HlpFlagNZ v
    LDY = 0
End Function

' ------------------------------------------------------------------------------
' STA  Store Accumulator
' Modes: ZP, ZP,X, ABS, ABS,X, ABS,Y, (ZP,X), (ZP),Y
' Flags: none
' ------------------------------------------------------------------------------
Public Function STA(ByVal op1 As String) As Long
    HlpWriteOperand op1, pRegs("A") And 255, 730, "STA"
    If pLastErrCode <> 0 Then STA = pLastErrCode: Exit Function
    STA = 0
End Function

' ------------------------------------------------------------------------------
' STX  Store X Register
' Modes: ZP, ZP,Y, ABS
' Flags: none
' ------------------------------------------------------------------------------
Public Function STX(ByVal op1 As String) As Long
    HlpWriteOperand op1, pRegs("X") And 255, 740, "STX"
    If pLastErrCode <> 0 Then STX = pLastErrCode: Exit Function
    STX = 0
End Function

' ------------------------------------------------------------------------------
' STY  Store Y Register
' Modes: ZP, ZP,X, ABS
' Flags: none
' ------------------------------------------------------------------------------
Public Function STY(ByVal op1 As String) As Long
    HlpWriteOperand op1, pRegs("Y") And 255, 750, "STY"
    If pLastErrCode <> 0 Then STY = pLastErrCode: Exit Function
    STY = 0
End Function

'================================================================================
' REGISTER TRANSFERS
'================================================================================

Public Function TAX() As Long
    SetReg "X", pRegs("A") And 255
    HlpFlagNZ pRegs("X")
    TAX = 0
End Function

Public Function TAY() As Long
    SetReg "Y", pRegs("A") And 255
    HlpFlagNZ pRegs("Y")
    TAY = 0
End Function

Public Function TXA() As Long
    SetReg "A", pRegs("X") And 255
    HlpFlagNZ pRegs("A")
    TXA = 0
End Function

Public Function TYA() As Long
    SetReg "A", pRegs("Y") And 255
    HlpFlagNZ pRegs("A")
    TYA = 0
End Function

' TSX: Transfer SP to X.  Note SP is 8-bit (0-255).
Public Function TSX() As Long
    SetReg "X", pRegs("SP") And 255
    HlpFlagNZ pRegs("X")
    TSX = 0
End Function

' TXS: Transfer X to SP.  No flag changes.
Public Function TXS() As Long
    pRegs("SP") = pRegs("X") And 255
    mUIRegsDirty = True
    TXS = 0
End Function

'================================================================================
' STACK OPERATIONS
' Hardware stack: page $01.  SP is 8-bit; full-address = $0100 + SP.
' Push decrements SP first; pull increments SP first.
'================================================================================

' PHA  Push Accumulator
Public Function PHA() As Long
    HlpStackPush CByte(pRegs("A") And 255)
    PHA = 0
End Function

' PHP  Push Processor Status (SR)
' B flag is always set to 1 when pushed by PHP (hardware behaviour).
Public Function PHP() As Long
    Dim sr As Byte
    sr = HlpFlagPack(True)        ' True = set B bit
    HlpStackPush sr
    PHP = 0
End Function

' PLA  Pull Accumulator
' Flags: N, Z
Public Function PLA() As Long
    Dim v As Byte
    v = HlpStackPull()
    SetReg "A", CLng(v)
    HlpFlagNZ CLng(v)
    PLA = 0
End Function

' PLP  Pull Processor Status
' Restores all flags.  B and unused bits behave per hardware spec.
Public Function PLP() As Long
    Dim sr As Byte
    sr = HlpStackPull()
    HlpFlagUnpack sr
    PLP = 0
End Function

'================================================================================
' ARITHMETIC
'================================================================================

' ------------------------------------------------------------------------------
' ADC  Add with Carry
' Modes: IMM, ZP, ZP,X, ABS, ABS,X, ABS,Y, (ZP,X), (ZP),Y
' Flags: N, V, Z, C
' Decimal mode (D=1): result is BCD-adjusted.
' ------------------------------------------------------------------------------
Public Function ADC(ByVal op1 As String) As Long
    Dim m As Long
    m = HlpReadOperand(op1, 100, "ADC")
    If pLastErrCode <> 0 Then ADC = pLastErrCode: Exit Function
    HlpALU_Add m
    ADC = 0
End Function

' ------------------------------------------------------------------------------
' SBC  Subtract with Carry (Borrow)
' IMPORTANT: On the 6510, C=1 means NO borrow.
'            Always CLC before SBC if you want simple subtraction.
' Internally: A = A - M - (1 - C)  which equals  A + (~M) + C
' Flags: N, V, Z, C
' ------------------------------------------------------------------------------
Public Function SBC(ByVal op1 As String) As Long
    Dim m As Long
    m = HlpReadOperand(op1, 110, "SBC")
    If pLastErrCode <> 0 Then SBC = pLastErrCode: Exit Function
    ' Convert subtraction to addition of one's-complement + carry
    HlpALU_Add (m Xor 255)        ' same as ADC(~M) - carry already in play
    SBC = 0
End Function

' ------------------------------------------------------------------------------
' INC  Increment Memory
' Modes: ZP, ZP,X, ABS, ABS,X
' Flags: N, Z
' ------------------------------------------------------------------------------
Public Function INC(ByVal op1 As String) As Long
    Dim addr As Long, v As Long
    addr = HlpResolveWriteAddr(op1, 120, "INC")
    If pLastErrCode <> 0 Then INC = pLastErrCode: Exit Function
    v = (CLng(pMemory.addr(addr)) + 1) And 255
    pMemory.addr(addr) = CByte(v)
    HlpFlagNZ v
    INC = 0
End Function

' INX  Increment X
Public Function INX() As Long
    Dim v As Long
    v = (pRegs("X") + 1) And 255
    SetReg "X", v
    HlpFlagNZ v
    INX = 0
End Function

' INY  Increment Y
Public Function INY() As Long
    Dim v As Long
    v = (pRegs("Y") + 1) And 255
    SetReg "Y", v
    HlpFlagNZ v
    INY = 0
End Function

' ------------------------------------------------------------------------------
' DEC  Decrement Memory
' Modes: ZP, ZP,X, ABS, ABS,X
' Flags: N, Z
' ------------------------------------------------------------------------------
Public Function DEC(ByVal op1 As String) As Long
    Dim addr As Long, v As Long
    addr = HlpResolveWriteAddr(op1, 130, "DEC")
    If pLastErrCode <> 0 Then DEC = pLastErrCode: Exit Function
    v = (CLng(pMemory.addr(addr)) - 1) And 255
    pMemory.addr(addr) = CByte(v)
    HlpFlagNZ v
    DEC = 0
End Function

' DEX  Decrement X
Public Function DEX() As Long
    Dim v As Long
    v = (pRegs("X") - 1) And 255
    SetReg "X", v
    HlpFlagNZ v
    DEX = 0
End Function

' DEY  Decrement Y
Public Function DEY() As Long
    Dim v As Long
    v = (pRegs("Y") - 1) And 255
    SetReg "Y", v
    HlpFlagNZ v
    DEY = 0
End Function

'================================================================================
' LOGICAL
'================================================================================

' AND  Logical AND with Accumulator
Public Function AND_(ByVal op1 As String) As Long
    Dim m As Long
    m = HlpReadOperand(op1, 200, "AND")
    If pLastErrCode <> 0 Then AND_ = pLastErrCode: Exit Function
    Dim r As Long
    r = (pRegs("A") And m) And 255
    SetReg "A", r
    HlpFlagNZ r
    AND_ = 0
End Function

' ORA  Logical OR with Accumulator
Public Function ORA(ByVal op1 As String) As Long
    Dim m As Long
    m = HlpReadOperand(op1, 210, "ORA")
    If pLastErrCode <> 0 Then ORA = pLastErrCode: Exit Function
    Dim r As Long
    r = (pRegs("A") Or m) And 255
    SetReg "A", r
    HlpFlagNZ r
    ORA = 0
End Function

' EOR  Exclusive OR with Accumulator
Public Function EOR(ByVal op1 As String) As Long
    Dim m As Long
    m = HlpReadOperand(op1, 220, "EOR")
    If pLastErrCode <> 0 Then EOR = pLastErrCode: Exit Function
    Dim r As Long
    r = (pRegs("A") Xor m) And 255
    SetReg "A", r
    HlpFlagNZ r
    EOR = 0
End Function

' ------------------------------------------------------------------------------
' BIT  Bit Test
' Modes: ZP, ABS
' Operation:   N = M7, V = M6, Z = (A AND M) == 0
' NOTE: does NOT alter A.
' ------------------------------------------------------------------------------
Public Function BIT(ByVal op1 As String) As Long
    Dim m As Long
    m = HlpReadOperand(op1, 230, "BIT")
    If pLastErrCode <> 0 Then BIT = pLastErrCode: Exit Function
    pFlags("N") = IIf((m And &H80) <> 0, 1, 0)
    pFlags("V") = IIf((m And &H40) <> 0, 1, 0)
    pFlags("Z") = IIf((pRegs("A") And m) = 0, 1, 0)
    mUIFlagsDirty = True
    BIT = 0
End Function

'================================================================================
' SHIFTS AND ROTATES
' op1 = "A" for accumulator mode, else an address expression.
'================================================================================

' ASL  Arithmetic Shift Left  (C <- [76543210] <- 0)
' Flags: N, Z, C
Public Function ASL(ByVal op1 As String) As Long
    Dim v As Long, r As Long
    If UCase$(Trim$(op1)) = "A" Or op1 = "" Then
        v = pRegs("A") And 255
        pFlags("C") = IIf((v And &H80) <> 0, 1, 0)
        r = (v * 2) And 255
        SetReg "A", r
    Else
        Dim addr As Long
        addr = HlpResolveWriteAddr(op1, 300, "ASL")
        If pLastErrCode <> 0 Then ASL = pLastErrCode: Exit Function
        v = CLng(pMemory.addr(addr)) And 255
        pFlags("C") = IIf((v And &H80) <> 0, 1, 0)
        r = (v * 2) And 255
        pMemory.addr(addr) = CByte(r)
    End If
    HlpFlagNZ r
    mUIFlagsDirty = True
    ASL = 0
End Function

' LSR  Logical Shift Right  (0 -> [76543210] -> C)
' Flags: N=0, Z, C
Public Function LSR(ByVal op1 As String) As Long
    Dim v As Long, r As Long
    If UCase$(Trim$(op1)) = "A" Or op1 = "" Then
        v = pRegs("A") And 255
        pFlags("C") = v And 1
        r = v \ 2
        SetReg "A", r
    Else
        Dim addr As Long
        addr = HlpResolveWriteAddr(op1, 310, "LSR")
        If pLastErrCode <> 0 Then LSR = pLastErrCode: Exit Function
        v = CLng(pMemory.addr(addr)) And 255
        pFlags("C") = v And 1
        r = v \ 2
        pMemory.addr(addr) = CByte(r)
    End If
    HlpFlagNZ r
    mUIFlagsDirty = True
    LSR = 0
End Function

' ROL  Rotate Left through Carry  (C <- [76543210] <- C)
' Flags: N, Z, C
Public Function ROL(ByVal op1 As String) As Long
    Dim v As Long, r As Long, oldC As Long
    oldC = pFlags("C") And 1
    If UCase$(Trim$(op1)) = "A" Or op1 = "" Then
        v = pRegs("A") And 255
        pFlags("C") = IIf((v And &H80) <> 0, 1, 0)
        r = ((v * 2) Or oldC) And 255
        SetReg "A", r
    Else
        Dim addr As Long
        addr = HlpResolveWriteAddr(op1, 320, "ROL")
        If pLastErrCode <> 0 Then ROL = pLastErrCode: Exit Function
        v = CLng(pMemory.addr(addr)) And 255
        pFlags("C") = IIf((v And &H80) <> 0, 1, 0)
        r = ((v * 2) Or oldC) And 255
        pMemory.addr(addr) = CByte(r)
    End If
    HlpFlagNZ r
    mUIFlagsDirty = True
    ROL = 0
End Function

' ROR  Rotate Right through Carry  (C -> [76543210] -> C)
' Flags: N, Z, C
Public Function ROR(ByVal op1 As String) As Long
    Dim v As Long, r As Long, oldC As Long
    oldC = pFlags("C") And 1
    If UCase$(Trim$(op1)) = "A" Or op1 = "" Then
        v = pRegs("A") And 255
        pFlags("C") = v And 1
        r = (v \ 2) Or (oldC * 128)
        SetReg "A", r
    Else
        Dim addr As Long
        addr = HlpResolveWriteAddr(op1, 330, "ROR")
        If pLastErrCode <> 0 Then ROR = pLastErrCode: Exit Function
        v = CLng(pMemory.addr(addr)) And 255
        pFlags("C") = v And 1
        r = (v \ 2) Or (oldC * 128)
        pMemory.addr(addr) = CByte(r)
    End If
    HlpFlagNZ r
    mUIFlagsDirty = True
    ROR = 0
End Function

'================================================================================
' COMPARE
'================================================================================

' CMP  Compare A with memory  (A - M, discard result)
' Flags: N, Z, C  (C=1 if A >= M, i.e. no borrow)
Public Function CMP(ByVal op1 As String) As Long
    Dim m As Long
    m = HlpReadOperand(op1, 400, "CMP")
    If pLastErrCode <> 0 Then CMP = pLastErrCode: Exit Function
    HlpALU_Compare pRegs("A") And 255, m
    CMP = 0
End Function

' CPX  Compare X with memory
Public Function CPX(ByVal op1 As String) As Long
    Dim m As Long
    m = HlpReadOperand(op1, 410, "CPX")
    If pLastErrCode <> 0 Then CPX = pLastErrCode: Exit Function
    HlpALU_Compare pRegs("X") And 255, m
    CPX = 0
End Function

' CPY  Compare Y with memory
Public Function CPY(ByVal op1 As String) As Long
    Dim m As Long
    m = HlpReadOperand(op1, 420, "CPY")
    If pLastErrCode <> 0 Then CPY = pLastErrCode: Exit Function
    HlpALU_Compare pRegs("Y") And 255, m
    CPY = 0
End Function

'================================================================================
' BRANCHES
' op1 is a label name (resolved by pLabels) or a hex PC-relative offset.
' All branches are PC-relative; the offset is signed (-128..+127) from
' the byte AFTER the branch instruction.  In your row-based model we use
' label resolution and jump directly to the target address row, same as JMP.
'================================================================================

Public Function BCC(ByVal op1 As String) As Long
    BCC = HlpBranch(pFlags("C") = 0, op1, 500, "BCC")
End Function

Public Function BCS(ByVal op1 As String) As Long
    BCS = HlpBranch(pFlags("C") = 1, op1, 501, "BCS")
End Function

Public Function BEQ(ByVal op1 As String) As Long
    BEQ = HlpBranch(pFlags("Z") = 1, op1, 502, "BEQ")
End Function

Public Function BMI(ByVal op1 As String) As Long
    BMI = HlpBranch(pFlags("N") = 1, op1, 503, "BMI")
End Function

Public Function BNE(ByVal op1 As String) As Long
    BNE = HlpBranch(pFlags("Z") = 0, op1, 504, "BNE")
End Function

Public Function BPL(ByVal op1 As String) As Long
    BPL = HlpBranch(pFlags("N") = 0, op1, 505, "BPL")
End Function

Public Function BVC(ByVal op1 As String) As Long
    BVC = HlpBranch(pFlags("V") = 0, op1, 506, "BVC")
End Function

Public Function BVS(ByVal op1 As String) As Long
    BVS = HlpBranch(pFlags("V") = 1, op1, 507, "BVS")
End Function

'================================================================================
' JUMPS AND SUBROUTINES
'================================================================================

' ------------------------------------------------------------------------------
' JMP  Jump to address
' Modes: ABS (op1 = "C000"), IND (op1 = "(C000)")
' No flags affected.
' ------------------------------------------------------------------------------
Public Function JMP(ByVal op1 As String) As Long
    Dim target As Long
    op1 = Trim$(op1)
    If Left$(op1, 1) = "(" And Right$(op1, 1) = ")" Then
        ' Indirect: read 16-bit little-endian pointer from address
        Dim ptrAddr As Long
        ptrAddr = usrHexToDec(Mid$(op1, 2, Len(op1) - 2)) And &HFFFF&
        Dim lo As Long, hi As Long
        lo = CLng(pMemory.addr(ptrAddr))
        ' 6510 page-wrap bug: (xxFF) wraps to (xx00), not (xx+1)00
        hi = CLng(pMemory.addr((ptrAddr And &HFF00&) Or ((ptrAddr + 1) And &HFF&)))
        target = (hi * 256&) Or lo
    Else
        target = HlpResolveAddr(op1, 600, "JMP")
        If pLastErrCode <> 0 Then JMP = pLastErrCode: Exit Function
    End If
    pRegs("PC") = target And &HFFFF&
    mUIRegsDirty = True
    JMP = 0
End Function

' ------------------------------------------------------------------------------
' JSR  Jump to Subroutine
' Pushes (PC - 1) onto stack (address of last byte of JSR instruction).
' In the row-based model "PC" at call time == address of JSR row, so we
' push (PC) and RTS adds 1 — net effect is identical to hardware.
' ------------------------------------------------------------------------------
Public Function JSR(ByVal op1 As String) As Long
    Dim target As Long
    target = HlpResolveAddr(op1, 610, "JSR")
    If pLastErrCode <> 0 Then JSR = pLastErrCode: Exit Function

    ' Push current PC (return address - 1 matches hardware; RTS adds 1)
    Dim retAddr As Long
    retAddr = pRegs("PC") And &HFFFF&
    HlpStackPush CByte((retAddr \ 256) And 255)   ' high byte first
    HlpStackPush CByte(retAddr And 255)            ' low byte

    pRegs("PC") = target And &HFFFF&
    mUIRegsDirty = True
    mUIStackDirty = True
    JSR = 0
End Function

' ------------------------------------------------------------------------------
' RTS  Return from Subroutine
' Pops 16-bit address (low then high) and adds 1.
' ------------------------------------------------------------------------------
Public Function RTS() As Long
    Dim lo As Long, hi As Long
    lo = CLng(HlpStackPull())
    hi = CLng(HlpStackPull())
    pRegs("PC") = ((hi * 256&) Or lo) + 1
    pRegs("PC") = pRegs("PC") And &HFFFF&
    mUIRegsDirty = True
    mUIStackDirty = True
    RTS = 0
End Function

' ------------------------------------------------------------------------------
' RTI  Return from Interrupt
' Pops SR (flags), then low byte of PC, then high byte of PC.
' Unlike RTS, does NOT add 1 to the popped address.
' ------------------------------------------------------------------------------
Public Function RTI() As Long
    Dim sr As Byte, lo As Long, hi As Long
    sr = HlpStackPull()
    HlpFlagUnpack sr
    lo = CLng(HlpStackPull())
    hi = CLng(HlpStackPull())
    pRegs("PC") = (hi * 256&) Or lo
    mUIRegsDirty = True
    mUIFlagsDirty = True
    mUIStackDirty = True
    RTI = 0
End Function

' ------------------------------------------------------------------------------
' BRK  Force Break (software interrupt)
' Pushes PC+1 (skip padding byte) then SR with B=1; sets I=1; jumps to
' IRQ/BRK vector at $FFFE/$FFFF.  In the emulator we treat it like HLT.
' ------------------------------------------------------------------------------
Public Function BRK() As Long
    ' Push PC+2 (one past padding byte)
    Dim retAddr As Long
    retAddr = (pRegs("PC") + 2) And &HFFFF&
    HlpStackPush CByte((retAddr \ 256) And 255)
    HlpStackPush CByte(retAddr And 255)
    ' Push SR with B=1
    Dim sr As Byte
    sr = HlpFlagPack(True)
    HlpStackPush sr
    ' Set I flag
    pFlags("I") = 1
    ' Signal halt so decExecute stops the run
    SetError ERR_EXEC_END, "BRK: software interrupt at PC=" & _
        Right$("0000" & Hex$(pRegs("PC")), 4)
    BRK = ERR_EXEC_END
End Function

' ------------------------------------------------------------------------------
' HLT  Halt (unofficial / WDC STP)
' Signals normal termination.
' ------------------------------------------------------------------------------
Public Function HLT() As Long
    SetError ERR_EXEC_END, "HLT: Program Halted"
    HLT = ERR_EXEC_END
End Function

'================================================================================
' FLAG OPERATIONS
'================================================================================

Public Function CLC() As Long: pFlags("C") = 0: mUIFlagsDirty = True: CLC = 0: End Function
Public Function CLD() As Long: pFlags("D") = 0: mUIFlagsDirty = True: CLD = 0: End Function
Public Function CLI() As Long: pFlags("I") = 0: mUIFlagsDirty = True: CLI = 0: End Function
Public Function CLV() As Long: pFlags("V") = 0: mUIFlagsDirty = True: CLV = 0: End Function
Public Function SEC() As Long: pFlags("C") = 1: mUIFlagsDirty = True: SEC = 0: End Function
Public Function SED() As Long: pFlags("D") = 1: mUIFlagsDirty = True: SED = 0: End Function
Public Function SEI() As Long: pFlags("I") = 1: mUIFlagsDirty = True: SEI = 0: End Function

'================================================================================
' MISC / PSEUDO-OPS
'================================================================================

Public Function NOP() As Long
    NOP = 0
End Function

' EQU_ (trailing underscore to avoid VBA keyword clash with clsDecCPU.EQU)
Public Function EQU_(ByVal op1 As String, ByVal label As String) As Long
    Dim errorBase As Long: errorBase = 800
    If op1 = "" Then
        SetError errorBase, "EQU: Missing operand"
        EQU_ = errorBase: Exit Function
    End If
    If label = "" Then
        SetError errorBase + 1, "EQU: Missing label"
        EQU_ = errorBase + 1: Exit Function
    End If
    Dim record As clsLabelRecord
    Set record = pLabels.GetLabel(label)
    If record Is Nothing Then
        SetError errorBase + 2, "EQU: Unknown label: " & label
        EQU_ = errorBase + 2: Exit Function
    End If
    If Not usrValIsHexString(op1) Then
        SetError errorBase + 3, "EQU: Invalid value: " & op1
        EQU_ = errorBase + 3: Exit Function
    End If
    EQU_ = 0
End Function

'================================================================================
' PUBLIC REGISTER / FLAG ACCESSORS (same interface as clsDecCPU)
'================================================================================

Public Function Reg(ByVal regName As String) As Long
    regName = UCase$(Trim$(regName))
    If pRegs.Exists(regName) Then
        Reg = CLng(pRegs(regName)) And &HFFFF&
    Else
        Reg = 0
    End If
End Function

Public Sub SetReg(ByVal regName As String, ByVal value As Long)
    regName = UCase$(Trim$(regName))
    If pRegs.Exists(regName) Then
        pRegs(regName) = value And &HFFFF&
    End If
    mUIRegsDirty = True
End Sub

Public Sub SetPC(ByVal pcVal As Long)
    pRegs("PC") = pcVal And &HFFFF&
    mUIRegsDirty = True
End Sub

' IncPC: advance PC by 1 row (mirrors clsDecCPU.IncPC)
Public Function IncPC() As Long
    pRegs("PC") = (pRegs("PC") + 1) And &HFFFF&
    mUIRegsDirty = True
    IncPC = pRegs("PC")
End Function

' GetCurrentIdx: row index within the program grid (PC - Line0_dec)
Public Function GetCurrentIdx() As Long
    Dim line0 As Long
    line0 = CLng(ThisWorkbook.Worksheets("CPU").Range("Line0_dec").value)
    GetCurrentIdx = CLng(pRegs("PC")) - line0
End Function

' SR as a packed byte (for Trace sheet)
Public Function GetSR() As Byte
    GetSR = HlpFlagPack(False)
End Function

'--- Dirty flags (same property names as clsDecCPU for UI code compatibility) ---
Public Property Get RegistersDirty() As Boolean:  RegistersDirty = mUIRegsDirty:  End Property
Public Property Let RegistersDirty(v As Boolean): mUIRegsDirty = v:               End Property
Public Property Get FlagsDirty() As Boolean:      FlagsDirty = mUIFlagsDirty:     End Property
Public Property Let FlagsDirty(v As Boolean):     mUIFlagsDirty = v:              End Property
Public Property Get StackDirty() As Boolean:      StackDirty = mUIStackDirty:     End Property
Public Property Let StackDirty(v As Boolean):     mUIStackDirty = v:              End Property

'--- Expose flag values individually (used by RefreshFlags in decExecute) ---
Public Function Flag(ByVal flagName As String) As Long
    flagName = UCase$(Trim$(flagName))
    If pFlags.Exists(flagName) Then
        Flag = CLng(pFlags(flagName))
    Else
        Flag = 0
    End If
End Function

'================================================================================
' STACK REFRESH (same signature as clsDecCPU.RefreshStack)
' Writes the stack display area on the CPU sheet.
'================================================================================
Public Sub RefreshStack(Optional ByVal force As Boolean = False)
    If Not (mUIStackDirty Or force) Then Exit Sub
    ' Delegate to the shared clsDecStack display helper if you wire it up,
    ' or write the raw SP / stack-page bytes directly to the sheet here.
    ' Minimal implementation: just show SP in the StackStart named range.
    Dim wsCPU As Worksheet
    Set wsCPU = ThisWorkbook.Worksheets("CPU")
    wsCPU.Range("StackStart").value = Right$("00" & Hex$(pRegs("SP") And 255), 2)
    mUIStackDirty = False
End Sub

'================================================================================
' INTERNAL HELPER FUNCTIONS
'================================================================================

'------------------------------------------------------------------------------
' HlpFlagPack
' Packs the pFlags dictionary into the 6510 SR byte.
' Bit layout:  N V 1 B D I Z C
' breakBit: PHP and BRK set B=1 when pushing; PLP/RTI restore B from stack.
'------------------------------------------------------------------------------
Private Function HlpFlagPack(Optional ByVal breakBit As Boolean = False) As Byte
    Dim sr As Long
    sr = 0
    If pFlags("C") <> 0 Then sr = sr Or &H1
    If pFlags("Z") <> 0 Then sr = sr Or &H2
    If pFlags("I") <> 0 Then sr = sr Or &H4
    If pFlags("D") <> 0 Then sr = sr Or &H8
    If breakBit Or (pFlags("B") <> 0) Then sr = sr Or &H10
    sr = sr Or &H20               ' bit 5 always 1 in hardware
    If pFlags("V") <> 0 Then sr = sr Or &H40
    If pFlags("N") <> 0 Then sr = sr Or &H80
    HlpFlagPack = CByte(sr And 255)
End Function

'------------------------------------------------------------------------------
' HlpFlagUnpack
' Restores pFlags from a packed SR byte (used by PLP, RTI).
'------------------------------------------------------------------------------
Private Sub HlpFlagUnpack(ByVal sr As Byte)
    pFlags("C") = IIf((sr And &H1)  <> 0, 1, 0)
    pFlags("Z") = IIf((sr And &H2)  <> 0, 1, 0)
    pFlags("I") = IIf((sr And &H4)  <> 0, 1, 0)
    pFlags("D") = IIf((sr And &H8)  <> 0, 1, 0)
    pFlags("B") = IIf((sr And &H10) <> 0, 1, 0)
    pFlags("V") = IIf((sr And &H40) <> 0, 1, 0)
    pFlags("N") = IIf((sr And &H80) <> 0, 1, 0)
    mUIFlagsDirty = True
End Sub

'------------------------------------------------------------------------------
' HlpFlagNZ
' Sets N and Z flags from an 8-bit result value.  Used by most instructions.
'------------------------------------------------------------------------------
Private Sub HlpFlagNZ(ByVal result As Long)
    result = result And 255
    pFlags("Z") = IIf(result = 0, 1, 0)
    pFlags("N") = IIf((result And &H80) <> 0, 1, 0)
    mUIFlagsDirty = True
End Sub

'------------------------------------------------------------------------------
' HlpALU_Add
' Core ADC logic: A = A + operand + C.
' Handles decimal mode (D=1) with BCD correction.
' Flags: N, V, Z, C
'------------------------------------------------------------------------------
Private Sub HlpALU_Add(ByVal m As Long)
    Dim a As Long, sum As Long
    a = pRegs("A") And 255
    m = m And 255

    sum = a + m + (pFlags("C") And 1)

    ' Overflow: set if sign of result differs from both operands' sign
    pFlags("V") = IIf(((Not (a Xor m)) And (a Xor sum) And &H80) <> 0, 1, 0)
    pFlags("C") = IIf(sum > 255, 1, 0)

    ' BCD correction when D=1
    If pFlags("D") = 1 Then
        If (a And &HF) + (m And &HF) + (pFlags("C") And 1) > 9 Then
            sum = sum + 6
        End If
        If sum > &H99 Then
            sum = sum + 96
            pFlags("C") = 1
        End If
    End If

    sum = sum And 255
    SetReg "A", sum
    HlpFlagNZ sum
    ' C already set above; NZ updates N and Z
End Sub

'------------------------------------------------------------------------------
' HlpALU_Compare
' Computes reg - m, sets N/Z/C, discards result.
' C=1 means reg >= m (no borrow) — opposite sense to 8080 CY.
'------------------------------------------------------------------------------
Private Sub HlpALU_Compare(ByVal regVal As Long, ByVal m As Long)
    Dim diff As Long
    regVal = regVal And 255
    m = m And 255
    diff = regVal - m
    pFlags("C") = IIf(regVal >= m, 1, 0)
    pFlags("Z") = IIf(diff = 0, 1, 0)
    pFlags("N") = IIf((diff And &H80) <> 0, 1, 0)
    mUIFlagsDirty = True
End Sub

'------------------------------------------------------------------------------
' HlpBranch
' Common implementation for all conditional branches.
' condition: Boolean — True = take the branch.
' op1: target label name or hex relative offset.
'------------------------------------------------------------------------------
Private Function HlpBranch(ByVal condition As Boolean, ByVal op1 As String, _
    ByVal errorBase As Long, ByVal mnem As String) As Long
    If Not condition Then
        HlpBranch = 0
        Exit Function
    End If
    Dim target As Long
    target = HlpResolveAddr(op1, errorBase, mnem)
    If pLastErrCode <> 0 Then HlpBranch = pLastErrCode: Exit Function
    pRegs("PC") = target And &HFFFF&
    mUIRegsDirty = True
    HlpBranch = 0
End Function

'------------------------------------------------------------------------------
' HlpResolveAddr
' Resolves a label or hex address string to a decimal address.
' Same semantics as your existing 8080 address resolution.
'------------------------------------------------------------------------------
Private Function HlpResolveAddr(ByVal op1 As String, ByVal errorBase As Long, _
    ByVal mnem As String) As Long
    pLastErrCode = 0
    op1 = Trim$(UCase$(op1))
    If op1 = "" Then
        SetError errorBase, mnem & ": Missing address operand"
        HlpResolveAddr = 0
        Exit Function
    End If

    ' Try label lookup first
    Dim addr As Long
    addr = pLabels.GetAddress(op1)
    If addr >= 0 Then
        HlpResolveAddr = addr
        Exit Function
    End If

    ' Fall back to hex literal
    If usrValIsHexString(op1) Then
        HlpResolveAddr = usrHexToDec(op1) And &HFFFF&
    Else
        SetError errorBase + 1, mnem & ": Invalid address: " & op1
        HlpResolveAddr = 0
    End If
End Function

'------------------------------------------------------------------------------
' HlpResolveWriteAddr
' Resolves the effective memory address for a write instruction (STA, INC, etc.)
' handling ZP, ZP,X, ZP,Y, ABS, ABS,X, ABS,Y operand syntax.
'------------------------------------------------------------------------------
Private Function HlpResolveWriteAddr(ByVal op1 As String, ByVal errorBase As Long, _
    ByVal mnem As String) As Long
    pLastErrCode = 0
    op1 = Trim$(UCase$(op1))
    HlpResolveWriteAddr = HlpEffectiveAddress(op1, errorBase, mnem)
End Function

'------------------------------------------------------------------------------
' HlpReadOperand
' Reads the effective 8-bit value for all addressing modes.
' Handles: #nn (IMM), nn (ZP), nn,X (ZP,X), nn,Y (ZP,Y),
'          nnnn (ABS), nnnn,X (ABS,X), nnnn,Y (ABS,Y),
'          (nn,X) (indexed indirect), (nn),Y (indirect indexed).
'------------------------------------------------------------------------------
Private Function HlpReadOperand(ByVal op1 As String, ByVal errorBase As Long, _
    ByVal mnem As String) As Long
    pLastErrCode = 0
    op1 = Trim$(op1)

    ' Immediate: #nn
    If Left$(op1, 1) = "#" Then
        Dim immStr As String
        immStr = Mid$(op1, 2)
        If Not usrValIsHexString(immStr) Then
            SetError errorBase, mnem & ": Invalid immediate: " & op1
            HlpReadOperand = 0: Exit Function
        End If
        HlpReadOperand = usrHexToDec(immStr) And 255
        Exit Function
    End If

    ' All other modes: resolve to an address then read memory
    Dim addr As Long
    addr = HlpEffectiveAddress(UCase$(op1), errorBase, mnem)
    If pLastErrCode <> 0 Then HlpReadOperand = 0: Exit Function
    HlpReadOperand = CLng(pMemory.addr(addr)) And 255
End Function

'------------------------------------------------------------------------------
' HlpWriteOperand
' Writes an 8-bit value to the effective address for all non-immediate modes.
'------------------------------------------------------------------------------
Private Sub HlpWriteOperand(ByVal op1 As String, ByVal value As Long, _
    ByVal errorBase As Long, ByVal mnem As String)
    pLastErrCode = 0
    Dim addr As Long
    addr = HlpEffectiveAddress(UCase$(Trim$(op1)), errorBase, mnem)
    If pLastErrCode <> 0 Then Exit Sub
    pMemory.addr(addr) = CByte(value And 255)
End Sub

'------------------------------------------------------------------------------
' HlpEffectiveAddress
' Central address resolver for all non-immediate modes.
' Returns the 16-bit effective address as a Long.
'
' Syntax accepted in op1 (upper-cased by caller):
'   nn            Zero page
'   nn,X          Zero page indexed X
'   nn,Y          Zero page indexed Y
'   nnnn          Absolute
'   nnnn,X        Absolute indexed X
'   nnnn,Y        Absolute indexed Y
'   (nn,X)        Indexed indirect (pre-indexed)
'   (nn),Y        Indirect indexed (post-indexed)
'   LABELNAME     Resolved to address via pLabels
'------------------------------------------------------------------------------
Private Function HlpEffectiveAddress(ByVal op1 As String, ByVal errorBase As Long, _
    ByVal mnem As String) As Long
    pLastErrCode = 0
    Dim addr As Long
    Dim baseAddr As Long
    Dim idxStr As String
    Dim commaPos As Long

    '--- (nn,X)  Indexed Indirect ---
    If Left$(op1, 1) = "(" And InStr(op1, ",X)") > 0 Then
        Dim zpStr As String
        zpStr = Mid$(op1, 2, InStr(op1, ",X)") - 2)
        baseAddr = usrHexToDec(zpStr) And 255
        Dim ptr As Long
        ptr = (baseAddr + (pRegs("X") And 255)) And 255
        addr = CLng(pMemory.addr(ptr)) Or (CLng(pMemory.addr((ptr + 1) And 255)) * 256&)
        HlpEffectiveAddress = addr And &HFFFF&
        Exit Function
    End If

    '--- (nn),Y  Indirect Indexed ---
    If Left$(op1, 1) = "(" And Right$(op1, 2) = "),Y" Then
        Dim zpStr2 As String
        zpStr2 = Mid$(op1, 2, Len(op1) - 4)
        baseAddr = usrHexToDec(zpStr2) And 255
        Dim ptr2 As Long
        ptr2 = CLng(pMemory.addr(baseAddr)) Or (CLng(pMemory.addr((baseAddr + 1) And 255)) * 256&)
        addr = (ptr2 + (pRegs("Y") And 255)) And &HFFFF&
        HlpEffectiveAddress = addr
        Exit Function
    End If

    '--- Check for index suffix (,X or ,Y) ---
    commaPos = InStr(op1, ",")
    If commaPos > 0 Then
        idxStr = Trim$(Mid$(op1, commaPos + 1))
        op1 = Left$(op1, commaPos - 1)
        baseAddr = HlpParseAddrOrLabel(op1, errorBase, mnem)
        If pLastErrCode <> 0 Then HlpEffectiveAddress = 0: Exit Function
        If idxStr = "X" Then
            If Len(op1) <= 2 Then           ' zero page,X wraps in page 0
                addr = (baseAddr + (pRegs("X") And 255)) And 255
            Else
                addr = (baseAddr + (pRegs("X") And 255)) And &HFFFF&
            End If
        ElseIf idxStr = "Y" Then
            If Len(op1) <= 2 Then           ' zero page,Y wraps in page 0
                addr = (baseAddr + (pRegs("Y") And 255)) And 255
            Else
                addr = (baseAddr + (pRegs("Y") And 255)) And &HFFFF&
            End If
        Else
            SetError errorBase, mnem & ": Unknown index register: " & idxStr
            HlpEffectiveAddress = 0: Exit Function
        End If
        HlpEffectiveAddress = addr
        Exit Function
    End If

    '--- No index: zero page or absolute or label ---
    HlpEffectiveAddress = HlpParseAddrOrLabel(op1, errorBase, mnem)
End Function

'------------------------------------------------------------------------------
' HlpParseAddrOrLabel
' Returns decimal address from a hex string or label name.
'------------------------------------------------------------------------------
Private Function HlpParseAddrOrLabel(ByVal s As String, ByVal errorBase As Long, _
    ByVal mnem As String) As Long
    If usrValIsHexString(s) Then
        HlpParseAddrOrLabel = usrHexToDec(s) And &HFFFF&
    Else
        Dim a As Long
        a = pLabels.GetAddress(s)
        If a >= 0 Then
            HlpParseAddrOrLabel = a
        Else
            SetError errorBase, mnem & ": Unresolved address: " & s
            HlpParseAddrOrLabel = 0
        End If
    End If
End Function

'------------------------------------------------------------------------------
' HlpStackPush / HlpStackPull
' 6510 hardware stack: always page $01 ($0100–$01FF).
' SP is 8-bit; push = write to $0100+SP then decrement; pull = increment then read.
'------------------------------------------------------------------------------
Private Sub HlpStackPush(ByVal value As Byte)
    Dim sp As Long
    sp = pRegs("SP") And 255
    pMemory.addr(&H100& + sp) = value
    pRegs("SP") = (sp - 1) And 255
    mUIStackDirty = True
End Sub

Private Function HlpStackPull() As Byte
    Dim sp As Long
    sp = ((pRegs("SP") And 255) + 1) And 255
    pRegs("SP") = sp
    HlpStackPull = pMemory.addr(&H100& + sp)
    mUIStackDirty = True
End Function

'------------------------------------------------------------------------------
' SetError  (mirrors clsDecCPU.SetError)
'------------------------------------------------------------------------------
Private Sub SetError(ByVal code As Long, ByVal msg As String)
    pLastErrCode = code
    pLastErrMsg  = msg
    If code <> ERR_EXEC_END Then
        Debug.Print "cls6510CPU Error " & code & ": " & msg
    End If
End Sub

'================================================================================
' END cls6510CPU
'================================================================================
