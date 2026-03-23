Option Explicit
'================================================================================
' FILE: Compiler6510_wiring_changes.bas
'
' This file describes the small changes needed in three existing modules to
' wire Assemble6510_ToMachine into the same places Assemble8080_ToMachine is.
' It is NOT a new module — apply the edits described below.
'================================================================================

'================================================================================
' CHANGE 1 — decExecute8080 module: ForceCompile
'
' ForceCompile currently always calls Assemble8080_ToMachine.
' Add a CPUMode branch so it calls the right assembler.
'
' FIND this sub in decExecute8080:
'
'   Public Sub ForceCompile(Optional ByVal compile As Boolean = False)
'       ...
'       If compile Then
'           Assemble8080_ToMachine
'           Exit Sub
'       End If
'       ...
'       If compile Then
'           Assemble8080_ToMachine
'       End If
'   End Sub
'
' REPLACE with this (two substitutions highlighted with ***):
'
'   Public Sub ForceCompile(Optional ByVal compile As Boolean = False)
'       Dim ofs_opcode As Long: ofs_opcode = Range("ofs_opcode").value
'       Dim memSize    As Long: memSize    = usrHexToDec(Range("MemSize").value)
'       Dim arrOpcode  As Variant
'       Dim EMPTY_RUN_LIMIT As Long: EMPTY_RUN_LIMIT = 8
'       Dim emptyRuns  As Long
'       Dim opcode     As String
'
'       If compile Then
'           *** AssembleForCPUMode      ' <-- CHANGED
'           Exit Sub
'       End If
'
'       Set base = Range("Line0").offset(0, ofs_opcode)
'       memSize = usrHexToDec(Range("MemSize").value)
'       arrOpcode = base.Resize(memSize, 1).value
'
'       Dim i As Long
'       i = 1
'       emptyRuns = 0
'       Do While i <= memSize And Not compile And (emptyRuns < EMPTY_RUN_LIMIT)
'           opcode = arrOpcode(i, 1)
'           If opcode = "DB" Then compile = True
'           If opcode = "" Then emptyRuns = emptyRuns + 1 Else emptyRuns = 0
'           i = i + 1
'       Loop
'
'       If compile Then
'           *** AssembleForCPUMode      ' <-- CHANGED
'       End If
'   End Sub
'================================================================================

'================================================================================
' CHANGE 2 — modGlobals: add AssembleForCPUMode helper
'
' PASTE this new Public sub into modGlobals alongside SelectEngine:
'
'   Public Sub AssembleForCPUMode()
'       Select Case CPUMode()
'           Case "6510": Assemble6510_ToMachine
'           Case Else:   Assemble8080_ToMachine
'       End Select
'   End Sub
'================================================================================

'================================================================================
' CHANGE 3 — TestRunner: CompileTest column already calls Assemble8080_ToMachine
'
' FIND in RunDynamicTests:
'
'   compileThis = (val(wsTest.Cells(nameRow, compileCol).value) = 1)
'   If compileThis Then Assemble8080_ToMachine
'
' REPLACE with:
'
'   compileThis = (val(wsTest.Cells(nameRow, compileCol).value) = 1)
'   If compileThis Then AssembleForCPUMode
'
' This means the Compile flag in the test table will use whichever assembler
' matches the test's CPU column — consistent with the CPU override added earlier.
'================================================================================

'================================================================================
' CHANGE 4 — CompileIfNeeded in Compiler8080
'
' FIND:
'   Assemble8080_ToMachine ' or whatever defaults you want
'
' REPLACE with:
'   AssembleForCPUMode
'================================================================================

'================================================================================
' HOW TO APPLY — STEP BY STEP
'
' 1. Add the new module:
'    Insert > Module, name it "Compiler6510", paste Compiler6510.bas into it.
'
' 2. Import the opcode sheet:
'    Open 6510_Op_to_Hex.xlsx, copy the "6510 Op to Hex" sheet into your
'    main workbook (right-click tab > Move or Copy > your workbook).
'    The sheet name must be exactly:  6510 Op to Hex
'
' 3. In modGlobals, add AssembleForCPUMode (Change 2 above).
'
' 4. In decExecute8080, update ForceCompile (Change 1 above) — two substitutions.
'
' 5. In TestRunner (RunDynamicTests), update the compile line (Change 3 above).
'
' 6. In Compiler8080 (CompileIfNeeded), update the assembly call (Change 4 above).
'
' 7. Compile check: Debug > Compile VBAProject — zero errors expected.
'
' 8. Verify:
'    a. Set CPUMode = 8080 — run existing tests — all pass as before.
'    b. Set CPUMode = 6510 — enter a simple 6510 program and click Assemble.
'       Expected: MemoryTable fills with correct bytes, errMessage = "Assemble complete (6510)".
'================================================================================

'================================================================================
' 6510 PSEUDO-OP REFERENCE (for writing programs and test cases)
'
' ORG  nnnn       Set program counter (same as 8080)
'                 e.g.  ORG  0200
'
' EQU  nnnn       Define constant (same as 8080, handled by clsLabels)
'                 e.g.  DELAY  EQU  FF
'
' DB   "string"   Define bytes from ASCII string + $00 terminator
'                 e.g.  MSG  DB  "HELLO"
'
' DB   nn         Define single raw byte
'                 e.g.  DB   EA       ; NOP byte
'
' DB   nn,nn,nn   Define multiple raw bytes (comma-separated hex or decimal)
'                 e.g.  DB   01,02,03
'
' DS   nn         Reserve nn bytes of $00  (n = hex)
'                 e.g.  DS   10       ; reserve 16 bytes
'                 Note: 6510 fills with $00, not $FF as the 8080 does
'
' ADDRESSING MODE COLUMN SYNTAX:
'   Immediate     #nn      LDA #4A    (load A with value $4A)
'   Zero page     nn       LDA 42     (load A from address $0042)
'   Zero page,X   nn,X     LDA 42,X   (load A from $0042+X)
'   Zero page,Y   nn,Y     LDX 42,Y   (load X from $0042+Y)
'   Absolute      nnnn     LDA C000   (load A from address $C000)
'   Absolute,X    nnnn,X   LDA C000,X (load A from $C000+X)
'   Absolute,Y    nnnn,Y   LDA C000,Y (load A from $C000+Y)
'   Indirect      (nnnn)   JMP (C000) (jump to address stored at $C000/$C001)
'   (Indirect,X)  (nn,X)   LDA (42,X) (zero-page pointer indexed by X)
'   (Indirect),Y  (nn),Y   LDA (42),Y (zero-page pointer, result indexed by Y)
'   Accumulator   A        ASL A      (or just leave OP1 blank)
'   Implied       (blank)  NOP        (no operand)
'
' BRANCHES — use a label name in OP1, same as JMP:
'   The assembler automatically computes the signed relative offset.
'   Branch range: -128 to +127 bytes from the instruction after the branch.
'   If the target is out of range, an assembly error is raised.
'================================================================================
