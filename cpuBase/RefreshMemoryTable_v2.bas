'================================================================================
' Sub:          RefreshMemoryTable
' Purpose:      Reads gMemory and writes the memory display table on the CPU
'               worksheet in a single bulk array assignment for maximum speed.
'
' Strategy:
'   1. Allocate a 2-D Variant array sized (totalRows x totalCols).
'   2. Fill every cell value in VBA memory (no sheet I/O in the loop).
'   3. Write the entire array to the sheet in ONE Range.Value assignment.
'   4. Clear only the exact range that was written (no over-clearing).
'
' Column layout written (left to right):
'   col 1  = Address  (decimal base address for the row)
'   col 2  = ASCII    (8-char preview: printable or ".")
'   col 3..10 = Hex bytes 0-7  (e.g. "3E", "00", "41" ...)
'
' Named ranges consumed:
'   MemoryTableAddress  - address column anchor cell (used for row/col position)
'   MemoryTable         - 8-column hex byte area (must start 2 cols right of above)
'   MemStart            - hex string first address  (e.g. "0100")
'   MemEnd              - hex string last  address  (e.g. "01FF")
'                         Falls back to MemStart + MemSize - 1 if MemEnd absent.
'================================================================================
Public Sub RefreshMemoryTable()

    Const BYTES_PER_ROW  As Long = 8
    Const COL_ADDR       As Long = 1   ' offset within our write array
    Const COL_ASCII      As Long = 2
    Const COL_BYTE0      As Long = 3   ' bytes occupy cols 3-10
    Const TOTAL_COLS     As Long = 10  ' 1 addr + 1 ascii + 8 bytes

    ' ── 1. Named-range anchors ────────────────────────────────────────────────
    Dim ws      As Worksheet
    Set ws = ThisWorkbook.Worksheets("CPU")

    Dim addrAnchor As Range
    Set addrAnchor = ws.Range("MemoryTableAddress") ' top-left of entire table

    ' ── 2. Memory window ──────────────────────────────────────────────────────
    Dim memStart As Long
    Dim memEnd   As Long
    memStart = usrHexToDec(CStr(ws.Range("MemStart").Value))

    On Error Resume Next
    Dim memEndRng As Range
    Set memEndRng = ws.Range("MemEnd")
    On Error GoTo 0
    If Not memEndRng Is Nothing Then
        memEnd = usrHexToDec(CStr(memEndRng.Value))
    Else
        memEnd = memStart + usrHexToDec(CStr(ws.Range("MemSize").Value)) - 1
    End If

    ' Cap at the physical table row count
    Dim maxPhysRows As Long
    maxPhysRows = ws.Range("MemoryTable").Rows.Count

    Dim totalRows As Long
    totalRows = ((memEnd - memStart) \ BYTES_PER_ROW) + 1
    If totalRows > maxPhysRows Then totalRows = maxPhysRows

    ' ── 3. Build 2-D array entirely in RAM ───────────────────────────────────
    Dim tbl() As Variant
    ReDim tbl(1 To totalRows, 1 To TOTAL_COLS)

    ' Pre-build hex lookup table: 0-255 -> "00".."FF"  (avoids Hex$+Right$ each call)
    Dim hexLookup(0 To 255) As String
    Dim h As Integer
    For h = 0 To 255
        hexLookup(h) = Right$("00" & Hex$(h), 2)
    Next h

    Dim rowIdx   As Long
    Dim colByte  As Long
    Dim rowAddr  As Long
    Dim curAddr  As Long
    Dim b        As Long
    Dim asciiStr As String
    Dim ch       As Integer

    rowAddr = memStart

    For rowIdx = 1 To totalRows

        tbl(rowIdx, COL_ADDR) = rowAddr   ' decimal address

        asciiStr = ""

        For colByte = 0 To BYTES_PER_ROW - 1
            curAddr = rowAddr + colByte
            If curAddr <= memEnd Then
                b = CLng(gMemory.addr(curAddr)) And &HFF&
                tbl(rowIdx, COL_BYTE0 + colByte) = hexLookup(b)

                ch = b
                If ch >= &H20 And ch <= &H7E Then
                    asciiStr = asciiStr & Chr$(ch)
                Else
                    asciiStr = asciiStr & "."
                End If
            Else
                tbl(rowIdx, COL_BYTE0 + colByte) = ""   ' past end of memory
                asciiStr = asciiStr & " "
            End If
        Next colByte

        tbl(rowIdx, COL_ASCII) = asciiStr

        rowAddr = rowAddr + BYTES_PER_ROW
    Next rowIdx

    ' ── 4. Single bulk write to sheet ─────────────────────────────────────────
    Application.ScreenUpdating = False

    Dim writeRange As Range
    Set writeRange = addrAnchor.Resize(totalRows, TOTAL_COLS)
    writeRange.Value = tbl

    Application.ScreenUpdating = True

End Sub
