Attribute VB_Name = "RAPOR_LISTE"
Option Explicit
' Report workbooks (1-2 INT-LDS, 3A DYN, ...) - 2026-10-30b
' One button per list sheet:
'   ICINDEKILER_AYARLA (sheet ICINDEKILER): headings -> page numbers -> hyperlinks
'   FIG_TAB_AYARLA     (sheet FIG_TAB):     page numbers -> hyperlinks
' Replaces the INT modules ICINDEKILER, SAYFA ("Sayfa Ayarla"), LINKLER / LINKLER_IC
' ("Kopru Ayarla"), which knew one report sheet (1_RAPOR).
' Report sheets = every sheet whose name starts with a number and "_" (1_GIRIS, 2_YUKLER,
' 1_RAPOR, ...), in tab order; Z_YATAY and the list sheets do not match.
'
' List sheet columns (row = one entry, only rows marked "x" in BC are used):
'   Z  heading code (ToC, filled here)     AZ page number (filled here)
'   BA row of the entry in its report sheet (formula)
'   BD report sheet of the entry, e.g. 1_GIRIS / 2_YUKLER / 1_RAPOR (formula)
'   BG optional link target 'Sheet'!Cell (formula; empty = the entry's row in BD), used for
'      figures placed on a landscape sheet (Z_YATAY) whose page comes from a 1_RAPOR row
' Report sheets: heading codes in column AD, first page number in W45.

Private Const SCRIPT_ID As String = "report-lists"
Private Const SCRIPT_VERSION As String = "2026-10-30b"
Private Const SCRIPT_CHANGELOG As String = "Report sheets found by name (1_RAPOR, 2_YUKLER...); optional link target in column BG"

Private Const COL_CODE As Long = 26     ' Z
Private Const COL_HEAD As Long = 30     ' AD on the report sheets
Private Const COL_PAGE As Long = 52     ' AZ
Private Const COL_ROW As Long = 53      ' BA
Private Const COL_MARK As Long = 55     ' BC
Private Const COL_SHEET As Long = 56    ' BD
Private Const COL_LINK As Long = 59     ' BG
Private Const LAST_ROW As Long = 507

Public Sub ICINDEKILER_AYARLA()
    Dim ws As Worksheet, msg As String
    On Error GoTo Fail
    BeginWork
    Set ws = ThisWorkbook.Worksheets("ICINDEKILER")
    msg = FillHeadings(ws)
    Application.Calculate                       ' BA / BD / AB depend on the new codes
    msg = msg & SetPages(ws)
    msg = msg & SetLinks(ws, "F", "C")
    EndWork
    Report "ICINDEKILER", msg
    Exit Sub
Fail:
    EndWork
    MsgBox "ICINDEKILER_AYARLA stopped: " & Err.Description, vbCritical, "ICINDEKILER"
End Sub

Public Sub FIG_TAB_AYARLA()
    Dim ws As Worksheet, msg As String
    On Error GoTo Fail
    BeginWork
    Set ws = ThisWorkbook.Worksheets("FIG_TAB")
    Application.Calculate
    msg = SetPages(ws)
    msg = msg & SetLinks(ws, "C", "B")
    EndWork
    Report "FIG_TAB", msg
    Exit Sub
Fail:
    EndWork
    MsgBox "FIG_TAB_AYARLA stopped: " & Err.Description, vbCritical, "FIG_TAB"
End Sub

' Heading codes of the report sheets, in tab order, into the x rows of column Z (from row 9).
Private Function FillHeadings(ws As Worksheet) As String
    Dim codes As New Collection, rep As Worksheet, v As Variant
    Dim i As Long, k As Long, lastRow As Long, repList As Collection

    Set repList = ReportSheets()
    If repList.Count = 0 Then Err.Raise vbObjectError + 1, , "no report sheet (name like 1_RAPOR) in this workbook"
    ws.Range(ws.Cells(9, COL_CODE), ws.Cells(1000, COL_CODE + 1)).ClearContents   ' Z9:AA1000, as before
    For Each rep In repList
        lastRow = rep.Cells(rep.Rows.Count, COL_HEAD).End(xlUp).Row
        For i = 1 To lastRow
            v = rep.Cells(i, COL_HEAD).Value
            If Not IsError(v) Then
                If Len(CStr(v)) > 0 And CStr(v) <> "0" Then codes.Add v
            End If
        Next i
    Next rep

    i = 1
    For k = 9 To 1000
        If IsX(ws.Cells(k, COL_MARK).Value) Then
            If i > codes.Count Then Exit For
            ws.Cells(k, COL_CODE).Value = codes(i)
            i = i + 1
        End If
    Next k
    If i <= codes.Count Then FillHeadings = (codes.Count - i + 1) & " headings did not fit: mark more rows with x in column BC." & vbLf
End Function

' Page of each listed entry = print page of its row in its report sheet + that sheet's W45 - 1.
Private Function SetPages(ws As Worksheet) As String
    Dim r As Long, m As Long, p As Long, nMiss As Long, rep As Worksheet

    ws.Range(ws.Cells(8, COL_PAGE), ws.Cells(1000, COL_PAGE)).ClearContents       ' AZ8:AZ1000
    For r = 8 To LAST_ROW
        If IsX(ws.Cells(r, COL_MARK).Value) Then
            m = RowOf(ws.Cells(r, COL_ROW).Value)
            Set rep = SheetOf(TextOf(ws.Cells(r, COL_SHEET).Value))
            p = 0
            If m > 0 And Not rep Is Nothing Then p = PrintPage(rep, m)
            If p > 0 Then
                ws.Cells(r, COL_PAGE).Value = p + RowOf(rep.Range("W45").Value) - 1
            ElseIf m > 0 Then
                nMiss = nMiss + 1
            End If
        End If
    Next r
    If nMiss > 0 Then SetPages = nMiss & " entries are outside the print area of their sheet (no page number)." & vbLf
End Function

' Hyperlink in column linkCol of each listed entry to column targetCol of its report row.
' The cell keeps its own formula and text (no TextToDisplay).
Private Function SetLinks(ws As Worksheet, linkCol As String, targetCol As String) As String
    Dim r As Long, m As Long, shName As String, nMiss As Long, target As String

    ws.Range("A1:BE" & LAST_ROW).ClearHyperlinks
    For r = 1 To LAST_ROW
        If IsX(ws.Cells(r, COL_MARK).Value) Then
            m = RowOf(ws.Cells(r, COL_ROW).Value)
            shName = TextOf(ws.Cells(r, COL_SHEET).Value)
            If m > 0 And Not SheetOf(shName) Is Nothing Then
                target = TextOf(ws.Cells(r, COL_LINK).Value)
                If target = "" Then target = "'" & shName & "'!" & targetCol & m
                ws.Hyperlinks.Add Anchor:=ws.Range(linkCol & r), Address:="", SubAddress:=target
                With ws.Range(linkCol & r).Font
                    .Name = "Cambria"
                    .Color = vbBlack
                    .Size = 11
                    .Underline = xlUnderlineStyleNone
                    .Italic = True
                End With
            ElseIf Len(TextOf(ws.Cells(r, COL_CODE).Value)) > 0 Then
                nMiss = nMiss + 1                  ' listed (code present) but its row was not found
            End If
        End If
    Next r
    If nMiss > 0 Then SetLinks = nMiss & " listed entries were not found in the report sheets (no link; check column BA)." & vbLf
End Function

' Print page (1, 2, ...) holding row m of rep, from its print area; 0 = not printed.
Private Function PrintPage(rep As Worksheet, m As Long) As Long
    Dim parts As Variant, a As String, i As Long
    parts = Split(Replace(rep.PageSetup.PrintArea, ";", ","), ",")
    For i = LBound(parts) To UBound(parts)
        a = Trim$(parts(i))
        If InStr(a, "!") > 0 Then a = Mid$(a, InStrRev(a, "!") + 1)
        If a <> "" Then
            If Not Intersect(rep.Range("B" & m), rep.Range(a)) Is Nothing Then
                PrintPage = i + 1
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub BeginWork()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
End Sub

Private Sub EndWork()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub Report(title As String, msg As String)
    If msg <> "" Then
        MsgBox msg, vbExclamation, title
    Else
        Application.StatusBar = title & " updated " & Format$(Now, "hh:nn")
    End If
End Sub

Private Function IsX(v As Variant) As Boolean
    If Not IsError(v) Then IsX = (CStr(v) = "x")
End Function

Private Function RowOf(v As Variant) As Long
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then
        If Len(CStr(v)) > 0 Then RowOf = CLng(v)
    End If
End Function

Private Function TextOf(v As Variant) As String
    If Not IsError(v) Then TextOf = CStr(v)
End Function

' Sheets named <number>_<text> (1_GIRIS, 2_YUKLER, 1_RAPOR, ...), in tab order.
Private Function ReportSheets() As Collection
    Dim ws As Worksheet, p As Long
    Set ReportSheets = New Collection
    For Each ws In ThisWorkbook.Worksheets
        p = InStr(ws.Name, "_")
        If p > 1 Then
            If IsNumeric(Left$(ws.Name, p - 1)) Then ReportSheets.Add ws
        End If
    Next ws
End Function

Private Function SheetOf(shName As String) As Worksheet
    If shName = "" Then Exit Function
    On Error Resume Next
    Set SheetOf = ThisWorkbook.Worksheets(shName)
    On Error GoTo 0
End Function
