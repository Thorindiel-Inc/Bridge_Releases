Attribute VB_Name = "KILIF"
' KILIF - strand pictures of a span's 4?_KILIF sheet: one text cell per strand row and section,
' one symbol per strand (bonded / sheathed), for the sections 0.02L ... 0.40L.
'
' Reads the strand table (name M_KILIF_HALAT of the sheet, 12 rows x 14 columns: two header rows, then
' strand rows 1-10; columns = row number, height, sheathed strands at 11 sections, strands in the row).
' Writes the picture block (name M_KILIF_CIZIM of the sheet, 40 rows x 33 columns): three bands of 10
' rows, strand row 1 at the bottom of each band, four pictures side by side, and the two legend cells.
' The mid-span and support pictures of the third band are formulas of the sheet and are not touched.
'
' Placement in a row (owner's strand layout drawing, 2026-11-19): every second strand from the row
' centre outwards, symmetric about the girder axis.
'   odd row, odd count  : centre strand, then 2, 4, 6 ... places to each side
'   odd row, even count : 1, 3, 5 ... places to each side of the centre strand
'   even row, even count: 1, 3, 5 ... places outside the two centre strands
'   even row, odd count : cannot be symmetric, one strand stays undrawn and is reported
'   every strand of the row sheathed: all of them, whatever the count
' PlaceRow is the only routine that knows this rule.
'
' Rework of the old sheet macro (audit V1-01 ... V1-16): runs on a ..._KILIF sheet only, checks the table
' before it clears anything, cells by name, one drawing routine, a fully sheathed row is drawn, the
' 0.40L picture follows the table, counts that cannot be drawn are reported.
Option Explicit

Private Const SCRIPT_ID As String = "kilif"
Private Const SCRIPT_VERSION As String = "2026-11-19a"
Private Const SCRIPT_CHANGELOG As String = "New managed module (was an old sheet macro): runs on a _KILIF sheet only, checks the strand table before it clears the pictures, cells found by M_ names; a fully sheathed row is drawn, the 0.40L picture follows the table, rows with an even strand count are drawn symmetric, counts that cannot be drawn are reported"
Private Const SCRIPT_NEEDS As String = "midas-core"   ' UPDATE adds this module when the workbook lacks it

Private Const N_ROWS As Long = 10        ' strand rows of the table
Private Const N_SEC As Long = 11         ' section columns of the table (x = 0 ... 0.5L)
Private Const FIRST_DRAWN As Long = 2    ' first section with a picture (0.02L)
Private Const LAST_DRAWN As Long = 10    ' last section with a picture (0.40L)
Private Const BAND_ROWS As Long = 15     ' rows from one band of pictures to the next
Private Const PIC_COLS As Long = 9       ' columns from one picture to the next
Private Const PICS_PER_BAND As Long = 4

' Button on the _KILIF sheets (OnAction KILIF.KILIF). Silent when everything was drawn.
Public Sub KILIF()
    Dim msg As String
    msg = KILIF_RUN()
    If msg <> "" Then MsgBox msg, vbExclamation, "KILIF"
End Sub

' Scripted entry (tests): sheetName "" = active sheet. Returns "" when every sheathed strand was drawn.
' "STOPPED: ..." = nothing was changed. Any other text = the pictures were drawn, with the listed remarks.
Public Function KILIF_RUN(Optional ByVal sheetName As String = "") As String
    Dim sh As Object, ws As Worksheet, blk As Range, drw As Range
    Dim tot(1 To N_ROWS) As Long, deb(1 To N_SEC, 1 To N_ROWS) As Long
    Dim i As Long, j As Long, v As Variant, why As String
    Dim pic() As String, leftOver As Long, notes As String, nNotes As Long
    Dim band As Long, col As Long

    On Error GoTo Fail
    If sheetName = "" Then
        Set sh = ActiveSheet
    Else
        On Error Resume Next
        Set sh = ActiveWorkbook.Worksheets(sheetName)
        On Error GoTo Fail
        If sh Is Nothing Then KILIF_RUN = "STOPPED: no sheet '" & sheetName & "' in this workbook.": Exit Function
    End If
    If TypeName(sh) <> "Worksheet" Then KILIF_RUN = "STOPPED: KILIF runs on a worksheet.": Exit Function
    Set ws = sh
    If UCase$(Right$(ws.Name, 6)) <> "_KILIF" Then
        KILIF_RUN = "STOPPED: KILIF runs on a ..._KILIF sheet (for example 4A_KILIF). The active sheet is '" & ws.Name & "'. Nothing was changed."
        Exit Function
    End If

    ' by the names of this sheet, created at the original places when missing
    Set blk = McAnchor(ws, "M_KILIF_HALAT", "B3:O14", True)
    Set drw = McAnchor(ws, "M_KILIF_CIZIM", "S6:AY45", True)
    If blk.Rows.Count <> N_ROWS + 2 Or blk.Columns.Count <> N_SEC + 3 Then
        KILIF_RUN = "STOPPED: the name M_KILIF_HALAT on " & ws.Name & " covers " & blk.Address(False, False) & _
            ", expected " & (N_ROWS + 2) & " rows x " & (N_SEC + 3) & " columns (strand table). Nothing was changed."
        Exit Function
    End If
    If drw.Rows.Count <> 2 * BAND_ROWS + N_ROWS Or drw.Columns.Count <> (PICS_PER_BAND - 1) * PIC_COLS + 6 Then
        KILIF_RUN = "STOPPED: the name M_KILIF_CIZIM on " & ws.Name & " covers " & drw.Address(False, False) & _
            ", expected " & (2 * BAND_ROWS + N_ROWS) & " rows x " & ((PICS_PER_BAND - 1) * PIC_COLS + 6) & " columns (picture block). Nothing was changed."
        Exit Function
    End If

    ' read and check the whole table before anything is cleared
    For i = 1 To N_ROWS
        v = blk.Cells(2 + i, N_SEC + 3).Value
        If Not ReadCount(v, tot(i)) Then
            KILIF_RUN = "STOPPED: " & ws.Name & "!" & blk.Cells(2 + i, N_SEC + 3).Address(False, False) & _
                " (strands in row " & i & ") is not a whole number >= 0. Nothing was changed."
            Exit Function
        End If
        For j = 1 To N_SEC
            v = blk.Cells(2 + i, 2 + j).Value
            If Not ReadCount(v, deb(j, i)) Then
                KILIF_RUN = "STOPPED: " & ws.Name & "!" & blk.Cells(2 + i, 2 + j).Address(False, False) & _
                    " (sheathed strands, row " & i & ") is not a whole number >= 0. Nothing was changed."
                Exit Function
            End If
            If deb(j, i) > tot(i) Then
                KILIF_RUN = "STOPPED: " & ws.Name & "!" & blk.Cells(2 + i, 2 + j).Address(False, False) & " has " & deb(j, i) & _
                    " sheathed strands in row " & i & ", which holds " & tot(i) & ". Nothing was changed."
                Exit Function
            End If
        Next j
    Next i

    ' all pictures in memory first
    ReDim pic(FIRST_DRAWN To LAST_DRAWN, 1 To N_ROWS)
    For j = FIRST_DRAWN To LAST_DRAWN
        For i = 1 To N_ROWS
            pic(j, i) = RowPicture(tot(i), deb(j, i), leftOver)
            If leftOver > 0 Then
                nNotes = nNotes + 1
                If nNotes <= 12 Then
                    If tot(i) Mod 2 = 0 And deb(j, i) Mod 2 = 1 Then
                        why = "an odd count cannot be symmetric in a row of " & tot(i)
                    Else
                        why = "every second strand of a row of " & tot(i) & " gives fewer places"
                    End If
                    notes = notes & vbLf & "  row " & i & ", section " & SectionText(blk.Cells(1, 2 + j).Value) & ": " & _
                        leftOver & " of " & deb(j, i) & " sheathed strand(s) not drawn (" & why & ")"
                End If
            End If
        Next i
    Next j

    ' clear as the old macro did: bands 1 and 2 whole, band 3 only its first picture (the others are formulas)
    drw.Rows(1).Resize(N_ROWS).ClearContents
    drw.Rows(1 + BAND_ROWS).Resize(N_ROWS).ClearContents
    drw.Cells(1 + 2 * BAND_ROWS, 1).Resize(N_ROWS, 4).ClearContents

    ' legend
    drw.Cells(29, 17).Value = ChrW$(&H25CE)
    drw.Cells(29, 21).Value = ChrW$(&H25C9)

    For j = FIRST_DRAWN To LAST_DRAWN
        band = (j - FIRST_DRAWN) \ PICS_PER_BAND
        col = 1 + PIC_COLS * ((j - FIRST_DRAWN) Mod PICS_PER_BAND)
        For i = 1 To N_ROWS
            If pic(j, i) <> "" Then drw.Cells(band * BAND_ROWS + N_ROWS + 1 - i, col).Value = pic(j, i)
        Next i
    Next j

    If nNotes > 0 Then
        KILIF_RUN = "KILIF drew the pictures on " & ws.Name & ", but the pictures show fewer sheathed strands than the table in " & _
            nNotes & " place(s):" & notes
        If nNotes > 12 Then KILIF_RUN = KILIF_RUN & vbLf & "  ..."
    End If
    Exit Function
Fail:
    KILIF_RUN = "STOPPED: KILIF failed (" & Err.Description & "). Check the pictures on the sheet and press the button again."
End Function

' A count cell: empty = 0; True when it is a whole number >= 0.
Private Function ReadCount(ByVal v As Variant, ByRef n As Long) As Boolean
    n = 0
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then ReadCount = True: Exit Function
    If VarType(v) = vbString Then
        If Trim$(v) = "" Then ReadCount = True
        Exit Function
    End If
    If Not IsNumeric(v) Then Exit Function
    If v < 0 Or v > 1000 Then Exit Function
    If Abs(v - Round(v, 0)) > 0.000001 Then Exit Function
    n = CLng(Round(v, 0))
    ReadCount = True
End Function

' Picture text of one strand row; leftOver = sheathed strands that found no place.
Private Function RowPicture(ByVal total As Long, ByVal nDeb As Long, ByRef leftOver As Long) As String
    Dim d() As Boolean, t As Long, s As String
    leftOver = 0
    If total <= 0 Then Exit Function
    leftOver = PlaceRow(total, nDeb, d)
    For t = 1 To total
        If d(t) Then s = s & ChrW$(&H25CE) Else s = s & ChrW$(&H25C9)
    Next t
    RowPicture = s
End Function

' The placement rule (see the top of the module): marks the sheathed places 1..total of one row and
' returns how many of nDeb found no place.
Private Function PlaceRow(ByVal total As Long, ByVal nDeb As Long, ByRef d() As Boolean) As Long
    Dim c As Long, k As Long, rest As Long
    ReDim d(1 To total)
    rest = nDeb
    If nDeb >= total Then                   ' the whole row is sheathed: nothing to place
        For k = 1 To total: d(k) = True: Next k
        PlaceRow = nDeb - total
        Exit Function
    End If
    If total Mod 2 = 1 Then
        c = (total + 1) \ 2                 ' centre strand
        If rest Mod 2 = 1 Then
            d(c) = True: rest = rest - 1: k = 2
        Else
            k = 1
        End If
        Do While rest >= 2 And c - k >= 1
            d(c - k) = True: d(c + k) = True
            rest = rest - 2: k = k + 2
        Loop
    Else
        c = total \ 2                       ' the axis lies between strands c and c + 1
        k = 1
        Do While rest >= 2 And c - k >= 1
            d(c - k) = True: d(c + 1 + k) = True
            rest = rest - 2: k = k + 2
        Loop
    End If
    PlaceRow = rest
End Function

' Heading of a section column as text (x / L), for the remarks.
Private Function SectionText(ByVal v As Variant) As String
    If IsError(v) Then SectionText = "?": Exit Function
    If IsNumeric(v) And Not IsEmpty(v) Then
        SectionText = Format$(v, "0.00") & "L"
    Else
        SectionText = CStr(v)
    End If
End Function

' Test entry: the placement of one row as text of 0 / 1 and the count left over, e.g. "010101010 | 0".
Public Function KILIF_PLACE(ByVal total As Long, ByVal nDeb As Long) As String
    Dim d() As Boolean, t As Long, s As String, rest As Long
    If total <= 0 Then KILIF_PLACE = " | " & nDeb: Exit Function
    rest = PlaceRow(total, nDeb, d)
    For t = 1 To total
        s = s & IIf(d(t), "1", "0")
    Next t
    KILIF_PLACE = s & " | " & rest
End Function
