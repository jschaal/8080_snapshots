Option Explicit
'================================================================================
' FILE 1 OF 4  —  modGlobals additions
'
' HOW TO APPLY:
'   Open modGlobals in the VBA editor.
'   Section A: paste declarations at the TOP of the module alongside the
'              existing Private pCPU / pMemory / pAddressList declarations.
'   Section B: paste functions anywhere after the existing gDecCPU / gMemory
'              property blocks.
'   Section C: REPLACE the existing SelectEngine sub entirely.
'================================================================================

' ============================================================================
' SECTION A  —  DECLARATIONS (paste at top of modGlobals)
' ============================================================================

' --- Shared execution cache (replaces the Private vars in decAssemble) ---
' These were Private in decAssemble purely because it was the only engine.
' Now that two engines share the same grid they belong here as Public.
Public gCacheValid      As Boolean
Public gCacheLine0Dec   As Long
Public gCacheCountRows  As Long
Public gCacheMemStart   As Long
Public gCacheMemEnd     As Long
Public gCacheOfsOpcode  As Long
Public gCacheOfsOp1     As Long
Public gCacheOfsOp2     As Long
Public gCacheOfsRowStat As Long
Public gCacheOfsLabel   As Long

Public gArrOpcode  As Variant
Public gArrRowStat As Variant
Public gArrOp1     As Variant
Public gArrOp2     As Variant
Public gArrLabel   As Variant

' --- Shared run-state (replaces Private gBreak / gCurrentIter in decAssemble) ---
Public gBreak       As Boolean
Public gCurrentIter As Long

' --- 6510 singleton (same pattern as existing Private pCPU) ---
Private p6510CPU As cls6510CPU

' ============================================================================
' SECTION B  —  FUNCTIONS (paste after existing gDecCPU / gMemory blocks)
' ============================================================================

'------------------------------------------------------------------------------
' InvalidateExecCache
' Public replacement for the Private InvalidateCodeCache in decAssemble.
' Both engines call this when the program grid changes.
'------------------------------------------------------------------------------
Public Sub InvalidateExecCache()
    gCacheValid = False
End Sub

'------------------------------------------------------------------------------
' g6510  —  6510 singleton accessor (mirrors existing gDecCPU property)
'------------------------------------------------------------------------------
Public Property Get g6510() As cls6510CPU
    If p6510CPU Is Nothing Then
        Set p6510CPU = New cls6510CPU
    End If
    Set g6510 = p6510CPU
End Property

'------------------------------------------------------------------------------
' ResetDecCPU6510
' Discards the 6510 instance; next g6510 call re-creates it.
' (mirrors existing ResetDecCPU)
'------------------------------------------------------------------------------
Public Sub ResetDecCPU6510()
    Set p6510CPU = Nothing
End Sub

'------------------------------------------------------------------------------
' CPUMode
' Reads the named range "CPUMode" from the CPU sheet.
' Returns "8080" as the safe default if the range is missing or blank.
'
' SETUP: add a cell to the CPU sheet, name it "CPUMode", give it a dropdown:
'   Data Validation > List > source:  8080,Z80,6510
'------------------------------------------------------------------------------
Public Function CPUMode() As String
    On Error Resume Next
    Dim v As String
    v = UCase$(Trim$(CStr(ThisWorkbook.Worksheets("CPU").Range("CPUMode").value)))
    If Err.Number <> 0 Or v = "" Then v = "8080"
    On Error GoTo 0
    CPUMode = v
End Function

'------------------------------------------------------------------------------
' IsCPU6510  —  convenience wrapper
'------------------------------------------------------------------------------
Public Function IsCPU6510() As Boolean
    IsCPU6510 = (CPUMode() = "6510")
End Function

' ============================================================================
' SECTION C  —  REPLACE existing SelectEngine with this
' ============================================================================

Public Sub SelectEngine()
    Select Case CPUMode()
        Case "6510": Execute6510   ' in decExecute6510
        Case Else:   Execute8080   ' in decExecute8080 (renamed from decAssemble)
    End Select
End Sub
