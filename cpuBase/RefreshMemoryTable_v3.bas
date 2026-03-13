'================================================================================
' Sub:          RefreshMemoryTable
' Purpose:      Reads gMemory and writes the memory display table on the CPU
'               worksheet using bulk array assignments for maximum speed.
'
' Strategy:
'   - Builds two separate arrays (address col, byte cols) and writes each
'     in a single Range.Value call.  The ASCII column is left completely
'     untouched so any formula there continues to work.
'   - The write range is anchored at the first DATA row of MemoryTable /
'     MemoryTableAddress, so the header row above is never touched.
'
' Named ranges consumed:
'   MemoryTableAddress  - first DATA cell of the address column (row 1 = data)
'   MemoryTable         - first DATA cell of the 8-column hex byte area
'   MemStart            - hex string first address  (e.g. "0100")
'   MemEnd              - hex string last  address  (e.g. "01FF")
'                         Falls back to MemStart + MemSize - 1 if MemEnd absent.
'================================================================================
Public Sub RefreshMemoryTable()

    Const BYTES_PER_ROW As Long = 8

    ' ── 1. Worksheet & anchors ────────────────────────────────────────────────
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CPU")

    ' These named ranges must point to the FIRST DATA ROW (below any header)
    Dim addrAnchor As Range   ' first data cell of the address column
    Dim memAnchor  As Range   ' first data cell of MemoryTable (byte col 0)
    Set addrAnchor = ws.Range("MemoryTableAddress")
    Set memAnchor  = ws.Range("MemoryTable")

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
    Dim totalRows As Long
    totalRows = ((memEnd - memStart) \ BYTES_PER_ROW) + 1
    If totalRows > memAnchor.Rows.Count Then totalRows = memAnchor.Rows.Count

    ' ── 3. Pre-build hex lookup: 0-255 -> "00".."FF" ─────────────────────────
    Dim hexLookup(0 To 255) As String
    Dim h As Integer
    For h = 0 To 255
        hexLookup(h) = Right$("00" & Hex$(h), 2)
    Next h

    ' ── 4. Allocate arrays ────────────────────────────────────────────────────
    Dim addrArr() As Variant
    Dim byteArr() As Variant
    ReDim addrArr(1 To totalRows, 1 To 1)
    ReDim byteArr(1 To totalRows, 1 To BYTES_PER_ROW)

    ' ── 5. Fill arrays in RAM ─────────────────────────────────────────────────
    Dim rowIdx  As Long
    Dim colByte As Long
    Dim rowAddr As Long
    Dim curAddr As Long
    Dim b       As Long

    rowAddr = memStart

    For rowIdx = 1 To totalRows

        addrArr(rowIdx, 1) = rowAddr

        For colByte = 0 To BYTES_PER_ROW - 1
            curAddr = rowAddr + colByte
            If curAddr <= memEnd Then
                b = CLng(gMemory.addr(curAddr)) And &HFF&
                byteArr(rowIdx, colByte + 1) = hexLookup(b)
            Else
                byteArr(rowIdx, colByte + 1) = ""
            End If
        Next colByte

        rowAddr = rowAddr + BYTES_PER_ROW
    Next rowIdx

    ' ── 6. Two bulk writes — address col, then byte cols ─────────────────────
    Application.ScreenUpdating = False

    addrAnchor.Resize(totalRows, 1).Value = addrArr
    memAnchor.Resize(totalRows, BYTES_PER_ROW).Value = byteArr

    Application.ScreenUpdating = True

End Sub
