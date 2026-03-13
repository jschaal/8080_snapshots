'================================================================================
' Sub:          RefreshMemoryTable
' Purpose:      Reads the global gMemory object and writes its contents into
'               the CPU worksheet's memory display table.
'
' Table layout (anchored at named range "MemoryTable"):
'   Col offset  0  = Hex bytes 0-7  (8 cells, one per byte)
'   Col offset -2  = Address (decimal) written into "MemoryTableAddress" column
'   Col offset -1  = ASCII  preview  (8 chars: printable or ".")
'
' Named ranges used:
'   MemoryTable        - top-left byte cell of the 8-column hex byte area
'   MemoryTableAddress - column that holds the row base address (decimal)
'   MemStart           - hex string of first memory address (e.g. "0100")
'   MemEnd             - hex string of last  memory address (e.g. "FFFF")
'                        OR MemSize (number of bytes) is used if MemEnd absent
'
' The ASCII column sits between MemoryTableAddress and MemoryTable byte columns.
' Bytes 0x20-0x7E are displayed as their character; everything else becomes ".".
'
' Call after Assemble8080_ToMachine or after the emulator modifies gMemory so
' the display reflects the live memory state.
'================================================================================
Public Sub RefreshMemoryTable()

    Const BYTES_PER_ROW As Long = 8

    Dim ws          As Worksheet
    Dim rngMem      As Range    ' top-left byte cell (MemoryTable named range)
    Dim addrCol     As Long     ' column index of the address  column (MemoryTableAddress)
    Dim asciiCol    As Long     ' column index of the ASCII    column (addrCol + 1)
    Dim byteCol0    As Long     ' column index of first byte   column (rngMem column)

    Dim memStart    As Long
    Dim memEnd      As Long
    Dim totalRows   As Long

    Dim rowOffset   As Long
    Dim rowAddr     As Long
    Dim colByte     As Long
    Dim b           As Long     ' byte value (0-255)
    Dim asciiStr    As String
    Dim ch          As Integer

    ' ── 1. Setup ──────────────────────────────────────────────────────────────
    Set ws      = ThisWorkbook.Worksheets("CPU")
    Set rngMem  = ws.Range("MemoryTable")         ' 8-column hex byte area

    addrCol  = ws.Range("MemoryTableAddress").Column
    asciiCol = addrCol + 1                         ' ASCII sits right of address
    byteCol0 = rngMem.Column                       ' first byte column

    memStart = usrHexToDec(CStr(ws.Range("MemStart").Value))

    ' Support both MemEnd (hex last addr) and MemSize (hex byte count)
    On Error Resume Next
    Dim memEndRng As Range
    Set memEndRng = ws.Range("MemEnd")
    On Error GoTo 0
    If Not memEndRng Is Nothing Then
        memEnd = usrHexToDec(CStr(memEndRng.Value))
    Else
        memEnd = memStart + usrHexToDec(CStr(ws.Range("MemSize").Value)) - 1
    End If

    ' Number of 8-byte rows needed, capped at the table's physical row count
    totalRows = ((memEnd - memStart) \ BYTES_PER_ROW) + 1
    If totalRows > rngMem.Rows.Count Then totalRows = rngMem.Rows.Count

    ' ── 2. Performance ────────────────────────────────────────────────────────
    Application.ScreenUpdating = False

    ' ── 3. Clear old content ─────────────────────────────────────────────────
    ' Address column
    ws.Cells(rngMem.Row, addrCol).Resize(totalRows, 1).ClearContents
    ' ASCII column
    ws.Cells(rngMem.Row, asciiCol).Resize(totalRows, 1).ClearContents
    ' Byte columns (all 8)
    rngMem.Resize(totalRows, BYTES_PER_ROW).ClearContents

    ' ── 4. Write new content ──────────────────────────────────────────────────
    rowOffset = 0
    rowAddr   = memStart

    Do While rowAddr <= memEnd And rowOffset < totalRows

        ' --- Address column ---
        ws.Cells(rngMem.Row + rowOffset, addrCol).Value = rowAddr

        ' --- Byte columns & ASCII preview ---
        asciiStr = ""
        For colByte = 0 To BYTES_PER_ROW - 1

            Dim curAddr As Long
            curAddr = rowAddr + colByte

            If curAddr <= memEnd Then
                b = CLng(gMemory.addr(curAddr)) And &HFF&

                ' Hex byte cell
                rngMem.Cells(rowOffset + 1, colByte + 1).Value = _
                    Right$("00" & Hex$(b), 2)

                ' ASCII: printable 0x20-0x7E → character, else "."
                ch = b
                If ch >= &H20 And ch <= &H7E Then
                    asciiStr = asciiStr & Chr$(ch)
                Else
                    asciiStr = asciiStr & "."
                End If
            Else
                ' Past end of memory – leave byte blank, pad ASCII with space
                asciiStr = asciiStr & " "
            End If

        Next colByte

        ' --- ASCII column ---
        ws.Cells(rngMem.Row + rowOffset, asciiCol).Value = asciiStr

        rowOffset = rowOffset + 1
        rowAddr   = rowAddr + BYTES_PER_ROW
    Loop

    ' ── 5. Restore ────────────────────────────────────────────────────────────
    Application.ScreenUpdating = True

End Sub
