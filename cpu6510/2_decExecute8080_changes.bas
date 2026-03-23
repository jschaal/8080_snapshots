Option Explicit
'================================================================================
' FILE 2 OF 4  —  decExecute8080 changes
'
' This is NOT a complete replacement of decAssemble.
' It lists the EXACT edits to make inside the existing decAssemble module
' after you rename it.
'
' HOW TO APPLY — do these steps in order:
'
' STEP 1: Rename the module
'   - In the VBA Project Explorer, click on decAssemble
'   - Press F4 to open the Properties window
'   - Change the (Name) field from "decAssemble" to "decExecute8080"
'
' STEP 2: Rename the main sub
'   Find:    Public Sub decExecute()
'   Replace: Public Sub Execute8080()
'   (there is exactly one occurrence)
'
' STEP 3: Delete the private cache declarations
'   Find and DELETE these lines from the top of the module
'   (lines 6495-6508 in your source export):
'
'       Private gCacheValid As Boolean
'       Private gCacheLine0Dec As Long
'       Private gCacheCountRows As Long
'       Private gCacheMemStart As Long
'       Private gCacheMemEnd As Long
'       Private gCacheOfsOpcode As Long, gCacheOfsOp1 As Long, gCacheOfsOp2 As Long, gCacheOfsRowStat As Long, gCacheOfsLabel As Long
'       Private gArrOpcode As Variant
'       Private gArrRowStat As Variant
'       Private gArrOp1 As Variant
'       Private gArrOp2 As Variant
'       Private gArrLabel As Variant
'       Private gBreak As Boolean
'       Private gCurrentIter As Long
'
'   These are now declared Public in modGlobals. VBA will find them
'   project-wide automatically — no other changes needed inside the module.
'
' STEP 4: Replace the Private InvalidateCodeCache sub
'   Find and DELETE this entire sub (3 lines):
'
'       Private Sub InvalidateCodeCache()
'           gCacheValid = False
'       End Sub
'
'   Then find its ONE call site inside Execute8080 (in the Reset block):
'   Find:    InvalidateCodeCache
'   Replace: InvalidateExecCache
'
' THAT IS ALL.
' Every other line in the module stays exactly as-is.
' The variable names gCacheValid, gArrOpcode, gBreak etc. are unchanged
' in the body — they now resolve to the Public declarations in modGlobals
' instead of the deleted Private ones. VBA resolves them automatically.
'================================================================================
