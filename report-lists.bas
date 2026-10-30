Attribute VB_Name = "RAPOR_LISTE"
Option Explicit
' Report workbooks (1-2 INT-LDS, 3A DYN, ...) - 2026-10-30d
' One button per list sheet, the macro builds the whole list (no helper columns):
'   ICINDEKILER_AYARLA (sheet ICINDEKILER): table of contents
'   FIG_TAB_AYARLA     (sheet FIG_TAB):     list of tables, then list of figures
' Report sheets = every sheet whose name starts with a number and "_" (1_GIRIS, 2_YUKLER,
' 1_RAPOR, ...), in tab order. On a report sheet:
'   heading  : code in AD (3A, 3A1, 3A11 ...), text in AE
'   caption  : code in Z ("T-12" table / "S-12" figure, S = S-cedilla), text in AG
'   page     : print page of the row (one print-area piece = one page) + W45 - 1
' A caption on a one-row print-area piece is a placeholder (the figure itself is on another
' sheet, e.g. Z_YATAY): the link goes to the row of the other sheet that holds the same code.
' A code found twice (e.g. "Burada: Tablo 20") is listed once, at its first row.
'
' List sheet layout = page 1: rows 2-46 printed (B:X), row 47 empty, list rows 9-43 (ToC) or
' 8-43 (FIG_TAB), title in C6. More entries -> the macro copies page 1 (A2:X47) below as new
' pages; fewer -> it deletes its extra pages. Rows are only inserted/deleted when the number
' of pages changes. Do not put anything else below page 1 in columns A:X of the list sheets.

Private Const SCRIPT_ID As String = "report-lists"
Private Const SCRIPT_VERSION As String = "2026-10-30d"
Private Const SCRIPT_CHANGELOG As String = "Faster (arrays, no full recalc, rows touched only when the page count changes); print area set through PageSetup and repaired if broken; calculation always back to automatic"

Private Const COL_CODE As Long = 26      ' Z  caption code on the report sheets
Private Const COL_HEAD As Long = 30      ' AD heading code
Private Const COL_HTEXT As Long = 31     ' AE heading text
Private Const COL_CTEXT As Long = 33     ' AG caption text
Private Const PAGE_STEP As Long = 46     ' page k = rows 2+46(k-1) .. 46+46(k-1)
Private Const PAGE_TOP As Long = 2
Private Const PAGE_LAST As Long = 46
Private Const TITLE_ROW As Long = 6
Private Const LIST_END As Long = 43

Private Type ListEntry
    Code As String       ' ToC: heading code; FIG_TAB: "T-12"
    Num As Long          ' FIG_TAB: caption number (sort key)
    Txt As String
    Page As Long
    Target As String     ' 'Sheet'!C12
End Type

Public Sub ICINDEKILER_AYARLA()
    Dim ws As Worksheet, msg As String
    Dim heads() As ListEntry, nHead As Long
    Dim noTab() As ListEntry, nNoTab As Long, noFig() As ListEntry, nNoFig As Long
    On Error GoTo Fail
    BeginWork
    Set ws = ThisWorkbook.Worksheets("ICINDEKILER")
    ScanReports heads, nHead, noTab, nNoTab, noFig, nNoFig, msg
    LayoutPages ws, 9, Array(Pages(nHead, 9)), Array(vbNullString)
    WriteToc ws, heads, nHead
    EndWork
    Report "ICINDEKILER", msg
    Exit Sub
Fail:
    EndWork
    MsgBox "ICINDEKILER_AYARLA stopped: " & Err.Description, vbCritical, "ICINDEKILER"
End Sub

Public Sub FIG_TAB_AYARLA()
    Dim ws As Worksheet, msg As String
    Dim tabs() As ListEntry, nTab As Long, figs() As ListEntry, nFig As Long
    Dim noHead() As ListEntry, nNo As Long, pTab As Long
    On Error GoTo Fail
    BeginWork
    Set ws = ThisWorkbook.Worksheets("FIG_TAB")
    ScanReports noHead, nNo, tabs, nTab, figs, nFig, msg
    SortByNum tabs, nTab
    SortByNum figs, nFig
    pTab = Pages(nTab, 8)
    LayoutPages ws, 8, Array(pTab, Pages(nFig, 8)), Array("Tablolar", ChrW(350) & "ekiller")
    WriteCaptions ws, tabs, nTab, 1, "Tablo "
    WriteCaptions ws, figs, nFig, pTab + 1, ChrW(350) & "ekil "
    EndWork
    Report "FIG_TAB", msg
    Exit Sub
Fail:
    EndWork
    MsgBox "FIG_TAB_AYARLA stopped: " & Err.Description, vbCritical, "FIG_TAB"
End Sub

' Headings (AD) and captions (Z) of all report sheets, in tab order and row order.
' Each report sheet is read once as arrays (columns C, Z, AD, AE, AG).
Private Sub ScanReports(heads() As ListEntry, nHead As Long, tabs() As ListEntry, nTab As Long, _
                        figs() As ListEntry, nFig As Long, msg As String)
    Dim rep As Worksheet, repList As Collection, seen As Object
    Dim r As Long, lastRow As Long, s As String, e As ListEntry, nOut As Long
    Dim nErr As Long, nDup As Long, dupList As String
    Dim vC As Variant, vZ As Variant, vAD As Variant, vAE As Variant, vAG As Variant
    Dim pTop() As Long, pBot() As Long, nPiece As Long, firstPage As Long, k As Long

    Set repList = ReportSheets()
    If repList.Count = 0 Then Err.Raise vbObjectError + 1, , "no report sheet (name like 1_RAPOR) in this workbook"
    Set seen = CreateObject("Scripting.Dictionary")
    ReDim heads(1 To 1): ReDim tabs(1 To 1): ReDim figs(1 To 1)
    nHead = 0: nTab = 0: nFig = 0

    For Each rep In repList
        lastRow = rep.UsedRange.Row + rep.UsedRange.Rows.Count - 1
        If lastRow < 2 Then lastRow = 2
        vC = rep.Range(rep.Cells(1, 3), rep.Cells(lastRow, 3)).Value
        vZ = rep.Range(rep.Cells(1, COL_CODE), rep.Cells(lastRow, COL_CODE)).Value
        vAD = rep.Range(rep.Cells(1, COL_HEAD), rep.Cells(lastRow, COL_HEAD)).Value
        vAE = rep.Range(rep.Cells(1, COL_HTEXT), rep.Cells(lastRow, COL_HTEXT)).Value
        vAG = rep.Range(rep.Cells(1, COL_CTEXT), rep.Cells(lastRow, COL_CTEXT)).Value
        ReadPieces rep, pTop, pBot, nPiece
        firstPage = RowOf(rep.Range("W45").Value)

        For r = 1 To lastRow
            If IsError(vAD(r, 1)) Then nErr = nErr + 1
            If IsError(vZ(r, 1)) Then nErr = nErr + 1

            s = TextOf(vAD(r, 1))
            If Len(s) > 0 And s <> "0" And Not seen.Exists("H|" & s) Then
                seen.Add "H|" & s, True
                e.Code = s: e.Num = 0
                e.Txt = TextOf(vAE(r, 1))
                k = PieceOf(r, pTop, pBot, nPiece)
                e.Page = 0
                If k > 0 Then e.Page = k + firstPage - 1 Else nOut = nOut + 1
                e.Target = "'" & rep.Name & "'!C" & r
                AddEntry heads, nHead, e
            End If

            s = TextOf(vZ(r, 1))
            If IsCaption(s) Then
                If seen.Exists("C|" & s) Then
                    If IsCaptionText(TextOf(vC(r, 1))) Then   ' a second real caption, not a "Burada:" mention
                        nDup = nDup + 1
                        dupList = dupList & " " & rep.Name & "!" & r & " (" & s & ")"
                    End If
                Else
                    seen.Add "C|" & s, True
                    e.Code = s
                    e.Num = CLng(Mid$(s, 3))
                    e.Txt = TextOf(vAG(r, 1))
                    k = PieceOf(r, pTop, pBot, nPiece)
                    e.Page = 0
                    If k > 0 Then e.Page = k + firstPage - 1 Else nOut = nOut + 1
                    e.Target = "'" & rep.Name & "'!B" & r
                    If k > 0 Then
                        If pTop(k) = pBot(k) Then e.Target = ElsewhereOr(s, e.Target)   ' one-row page = placeholder
                    End If
                    If Left$(s, 1) = "T" Then AddEntry tabs, nTab, e Else AddEntry figs, nFig, e
                End If
            End If
        Next r
    Next rep
    If nOut > 0 Then msg = msg & nOut & " entries are outside the print area of their sheet (no page number)." & vbLf
    If nErr > 0 Then msg = msg & nErr & " heading/caption code cells (columns AD / Z) show an error - those entries are missing." & vbLf
    If nDup > 0 Then msg = msg & nDup & " captions use a number that is already taken (not listed):" & dupList & vbLf
End Sub

' Print-area pieces of a sheet (first/last row of each), read once.
Private Sub ReadPieces(ws As Worksheet, pTop() As Long, pBot() As Long, n As Long)
    Dim parts As Variant, a As String, i As Long, rg As Range
    parts = Split(Replace(ws.PageSetup.PrintArea, ";", ","), ",")
    n = 0
    ReDim pTop(1 To UBound(parts) - LBound(parts) + 2)
    ReDim pBot(1 To UBound(parts) - LBound(parts) + 2)
    For i = LBound(parts) To UBound(parts)
        a = Trim$(parts(i))
        If InStr(a, "!") > 0 Then a = Mid$(a, InStrRev(a, "!") + 1)
        If a <> "" Then
            Set rg = ws.Range(a)
            n = n + 1
            pTop(n) = rg.Row
            pBot(n) = rg.Row + rg.Rows.Count - 1
        End If
    Next i
End Sub

' Piece number (1, 2, ...) holding row r; 0 = not printed.
Private Function PieceOf(r As Long, pTop() As Long, pBot() As Long, n As Long) As Long
    Dim i As Long
    For i = 1 To n
        If r >= pTop(i) And r <= pBot(i) Then
            PieceOf = i
            Exit Function
        End If
    Next i
End Function

' Caption text shown in column C: "Tablo ..." or "<S-cedilla>ekil ..."
Private Function IsCaptionText(s As String) As Boolean
    IsCaptionText = (Left$(s, 6) = "Tablo " Or Left$(s, 6) = ChrW(350) & "ekil ")
End Function

' "T-12" or "<S-cedilla>-12"
Private Function IsCaption(s As String) As Boolean
    Dim k As String
    If Len(s) < 3 Then Exit Function
    k = Left$(s, 1)
    If (k = "T" Or k = ChrW(350)) And Mid$(s, 2, 1) = "-" Then
        IsCaption = IsAllDigits(Mid$(s, 3))
    End If
End Function

Private Function IsAllDigits(s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Or Len(s) > 6 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function

' Row on a non-report, non-list sheet that holds the caption code (landscape figure pages).
Private Function ElsewhereOr(code As String, fallback As String) As String
    Dim ws As Worksheet, hit As Range
    ElsewhereOr = fallback
    For Each ws In ThisWorkbook.Worksheets
        If Not IsReportName(ws.Name) And ws.Name <> "ICINDEKILER" And ws.Name <> "FIG_TAB" _
           And ws.Name <> "INPUT" Then
            Set hit = ws.UsedRange.Find(What:=code, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
            If Not hit Is Nothing Then
                ElsewhereOr = "'" & ws.Name & "'!B" & hit.Row
                Exit Function
            End If
        End If
    Next ws
End Function

Private Sub AddEntry(arr() As ListEntry, n As Long, e As ListEntry)
    n = n + 1
    If n > UBound(arr) Then ReDim Preserve arr(1 To n * 2)
    arr(n) = e
End Sub

Private Sub SortByNum(arr() As ListEntry, n As Long)
    Dim i As Long, j As Long, t As ListEntry
    For i = 2 To n
        t = arr(i)
        j = i - 1
        Do While j >= 1
            If arr(j).Num <= t.Num Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = t
    Next i
End Sub

' Pages needed for n entries (at least one page per section).
Private Function Pages(n As Long, firstRow As Long) As Long
    Dim cap As Long
    cap = LIST_END - firstRow + 1
    Pages = (n + cap - 1) \ cap
    If Pages < 1 Then Pages = 1
End Function

Private Function PageRow(k As Long) As Long
    PageRow = PAGE_TOP + PAGE_STEP * (k - 1)
End Function

' Make the sheet hold exactly sum(sectionPages) pages with empty lists; section titles on the
' first page of each section (empty title = keep the page 1 title as it is).
Private Sub LayoutPages(ws As Worksheet, firstRow As Long, sectionPages As Variant, titles As Variant)
    Dim nOld As Long, nNew As Long, k As Long, s As Long, r As Long, top As Long

    Dim repaired As Boolean
    RemoveStrayPrintName ws
    nOld = CheckedPageCount(ws, repaired)
    If ws.ProtectContents Then Err.Raise vbObjectError + 4, , ws.Name & " is protected - unprotect it first"
    For s = LBound(sectionPages) To UBound(sectionPages)
        nNew = nNew + sectionPages(s)
    Next s

    If nNew <> nOld Then
        ' page count changes: remove pages 2..nOld, tidy page 1, copy it nNew-1 times
        If nOld > 1 Then ws.Rows(PageRow(2) & ":" & PageRow(nOld) + PAGE_STEP - 1).Delete
        ClearList ws, 1, firstRow
        NormalizeRows ws, PAGE_TOP, firstRow
        For k = 2 To nNew
            top = PageRow(k)
            ws.Rows(top & ":" & top + PAGE_STEP - 1).Insert Shift:=xlDown
            ws.Range(ws.Cells(PAGE_TOP, 1), ws.Cells(PAGE_TOP + PAGE_STEP - 1, 24)).Copy ws.Cells(top, 1)
            For r = 0 To PAGE_STEP - 1
                ws.Rows(top + r).RowHeight = ws.Rows(PAGE_TOP + r).RowHeight
            Next r
            ws.Cells(top + TITLE_ROW - PAGE_TOP, 3).MergeArea.ClearContents
        Next k
        Application.CutCopyMode = False
        SetPrintArea ws, nNew
    Else
        ' same pages: only empty the lists (one-row entries everywhere)
        For k = 1 To nNew
            ClearList ws, k, firstRow
            NormalizeRows ws, PageRow(k), firstRow
        Next k
        If repaired Then SetPrintArea ws, nNew
    End If

    k = 1
    For s = LBound(sectionPages) To UBound(sectionPages)
        If Len(titles(s)) > 0 Then ws.Cells(PageRow(k) + TITLE_ROW - PAGE_TOP, 3).MergeArea.Cells(1, 1).Value = titles(s)
        k = k + sectionPages(s)
    Next s
End Sub

Private Sub ClearList(ws As Worksheet, k As Long, firstRow As Long)
    Dim rg As Range
    Set rg = ws.Range(ws.Cells(PageRow(k) + firstRow - PAGE_TOP, 2), ws.Cells(PageRow(k) + LIST_END - PAGE_TOP, 24))
    rg.ClearHyperlinks
    rg.ClearContents
End Sub

' List rows of the page starting at row top: every row merged like the first list row
' (rows that already have that shape are left alone - usually all of them).
Private Sub NormalizeRows(ws As Worksheet, top As Long, firstRow As Long)
    Dim r As Long, r0 As Long, c As Long, want() As String, nW As Long, ok As Boolean
    r0 = top + firstRow - PAGE_TOP
    ReDim want(1 To 24)
    For c = 2 To 24
        If ws.Cells(r0, c).MergeCells Then
            If ws.Cells(r0, c).MergeArea.Column = c Then
                nW = nW + 1
                want(nW) = c & ":" & ws.Cells(r0, c).MergeArea.Columns.Count
            End If
        End If
    Next c
    For r = r0 + 1 To top + LIST_END - PAGE_TOP
        ok = True
        For c = 1 To nW
            If Not SameMerge(ws, r, want(c)) Then
                ok = False
                Exit For
            End If
        Next c
        If Not ok Then
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 24)).UnMerge
            ws.Range(ws.Cells(r0, 2), ws.Cells(r0, 24)).Copy ws.Cells(r, 2)
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 24)).ClearContents
            ws.Rows(r).RowHeight = ws.Rows(r0).RowHeight
        End If
    Next r
    Application.CutCopyMode = False
End Sub

' "col:width" merge expected at row r
Private Function SameMerge(ws As Worksheet, r As Long, spec As String) As Boolean
    Dim c As Long, w As Long, ma As Range
    c = CLng(Split(spec, ":")(0)): w = CLng(Split(spec, ":")(1))
    Set ma = ws.Cells(r, c).MergeArea
    SameMerge = (ma.Row = r And ma.Column = c And ma.Rows.Count = 1 And ma.Columns.Count = w)
End Function

' Version 2026-10-30c used Names.Add "Print_Area"; in a Turkish Excel that made an ordinary
' name next to Excel's own print area (shown as Yazdirma_Alani). Delete that ordinary name -
' only when Excel's own one exists too (its local name differs), so English Excel is untouched.
Private Sub RemoveStrayPrintName(ws As Worksheet)
    Dim i As Long, nm As Name, stray As Name, hasOwn As Boolean, sfx As String
    sfx = "!Print_Area"
    For i = ThisWorkbook.Names.Count To 1 Step -1
        Set nm = ThisWorkbook.Names(i)
        If Right$(nm.Name, Len(sfx)) = sfx And IsOnSheet(nm, ws) Then
            If nm.NameLocal = nm.Name Then Set stray = nm Else hasOwn = True
        End If
    Next i
    If hasOwn And Not stray Is Nothing Then stray.Delete
End Sub

Private Function IsOnSheet(nm As Name, ws As Worksheet) As Boolean
    On Error Resume Next
    IsOnSheet = (nm.Parent.Name = ws.Name) And (TypeOf nm.Parent Is Worksheet)
    On Error GoTo 0
End Function

' Print area = one piece per page, through PageSetup (keeps Excel's own print-area name).
Private Sub SetPrintArea(ws As Worksheet, nPages As Long)
    Dim k As Long, rg As Range
    For k = 1 To nPages
        If rg Is Nothing Then
            Set rg = ws.Range(ws.Cells(PageRow(k), 2), ws.Cells(PageRow(k) + PAGE_LAST - PAGE_TOP, 24))
        Else
            Set rg = Union(rg, ws.Range(ws.Cells(PageRow(k), 2), ws.Cells(PageRow(k) + PAGE_LAST - PAGE_TOP, 24)))
        End If
    Next k
    ws.PageSetup.PrintArea = rg.Address
End Sub

' Pages now on the sheet = print-area pieces; each must be one of the standard pages.
' An empty / #REF! print area (left by version 2026-10-30c) is repaired: pages are counted
' from the footer formula in column B (row 45 of each page).
Private Function CheckedPageCount(ws As Worksheet, repaired As Boolean) As Long
    Dim parts As Variant, i As Long, a As String, want As String
    a = ws.PageSetup.PrintArea
    If Len(a) = 0 Or InStr(a, "#REF") > 0 Then
        Do While ws.Cells(PageRow(CheckedPageCount + 1) + 43, 2).HasFormula
            CheckedPageCount = CheckedPageCount + 1
        Loop
        If CheckedPageCount = 0 Then CheckedPageCount = 1
        repaired = True
        Exit Function
    End If
    parts = Split(Replace(a, ";", ","), ",")
    For i = LBound(parts) To UBound(parts)
        a = Trim$(parts(i))
        If InStr(a, "!") > 0 Then a = Mid$(a, InStrRev(a, "!") + 1)
        want = "$B$" & PageRow(i + 1) & ":$X$" & PageRow(i + 1) + PAGE_LAST - PAGE_TOP
        If UCase$(a) <> want Then Err.Raise vbObjectError + 2, , _
            "print area of " & ws.Name & " is not the standard page layout (" & a & ", expected " & want & ")"
    Next i
    CheckedPageCount = UBound(parts) - LBound(parts) + 1
End Function

' ToC: C = 3.A.1., F = text (link), V = page, written page by page as arrays.
' Level-1 headings are bold through the sheet's conditional format (LEN($C9)=2).
Private Sub WriteToc(ws As Worksheet, heads() As ListEntry, n As Long)
    Dim i As Long, r As Long
    Dim vC() As Variant, vF() As Variant, vV() As Variant
    If n = 0 Then Exit Sub
    ReDim vC(1 To n, 1 To 1): ReDim vF(1 To n, 1 To 1): ReDim vV(1 To n, 1 To 1)
    For i = 1 To n
        vC(i, 1) = "'" & Numbering(heads(i).Code)
        vF(i, 1) = "'" & heads(i).Txt
        If heads(i).Page > 0 Then vV(i, 1) = heads(i).Page Else vV(i, 1) = Empty
    Next i
    PutColumn ws, vC, n, 9, 1, 3
    PutColumn ws, vF, n, 9, 1, 6
    PutColumn ws, vV, n, 9, 1, 22
    For i = 1 To n
        r = ListRow(i, 9, 1)
        ws.Hyperlinks.Add Anchor:=ws.Cells(r, 6), Address:="", SubAddress:=heads(i).Target
    Next i
    LinkFont ws, 6, n, 9, 1
End Sub

Private Sub WriteCaptions(ws As Worksheet, caps() As ListEntry, n As Long, firstPage As Long, prefix As String)
    Dim i As Long, r As Long, vC() As Variant, vV() As Variant
    If n = 0 Then Exit Sub
    ReDim vC(1 To n, 1 To 1): ReDim vV(1 To n, 1 To 1)
    For i = 1 To n
        vC(i, 1) = "'" & prefix & caps(i).Num & ": " & caps(i).Txt
        If caps(i).Page > 0 Then vV(i, 1) = caps(i).Page Else vV(i, 1) = Empty
    Next i
    PutColumn ws, vC, n, 8, firstPage, 3
    PutColumn ws, vV, n, 8, firstPage, 22
    For i = 1 To n
        r = ListRow(i, 8, firstPage)
        ws.Hyperlinks.Add Anchor:=ws.Cells(r, 3), Address:="", SubAddress:=caps(i).Target
    Next i
    LinkFont ws, 3, n, 8, firstPage
End Sub

' Write v(1..n) into column col, one block per page.
Private Sub PutColumn(ws As Worksheet, v() As Variant, n As Long, firstRow As Long, firstPage As Long, col As Long)
    Dim cap As Long, i0 As Long, m As Long, i As Long, blk() As Variant, r As Long
    cap = LIST_END - firstRow + 1
    For i0 = 1 To n Step cap
        m = n - i0 + 1
        If m > cap Then m = cap
        ReDim blk(1 To m, 1 To 1)
        For i = 1 To m
            blk(i, 1) = v(i0 + i - 1, 1)
        Next i
        r = ListRow(i0, firstRow, firstPage)
        ws.Range(ws.Cells(r, col), ws.Cells(r + m - 1, col)).Value = blk
    Next i0
End Sub

' Hyperlinks take the Hyperlink style; put the list font back, one range per page.
Private Sub LinkFont(ws As Worksheet, col As Long, n As Long, firstRow As Long, firstPage As Long)
    Dim cap As Long, i0 As Long, m As Long, r As Long
    cap = LIST_END - firstRow + 1
    For i0 = 1 To n Step cap
        m = n - i0 + 1
        If m > cap Then m = cap
        r = ListRow(i0, firstRow, firstPage)
        With ws.Range(ws.Cells(r, col), ws.Cells(r + m - 1, col)).Font
            .Name = "Cambria"
            .Color = vbBlack
            .Size = 11
            .Underline = xlUnderlineStyleNone
            .Italic = True
        End With
    Next i0
End Sub

' Row of entry i (1-based) when the section starts on page firstPage.
Private Function ListRow(i As Long, firstRow As Long, firstPage As Long) As Long
    Dim cap As Long
    cap = LIST_END - firstRow + 1
    ListRow = PageRow(firstPage + (i - 1) \ cap) + (firstRow - PAGE_TOP) + (i - 1) Mod cap
End Function

' 3A11 -> 3.A.1.1. (first four characters, as the old sheet formula)
Private Function Numbering(code As String) As String
    Dim i As Long
    For i = 1 To Len(code)
        If i > 4 Then Exit For
        Numbering = Numbering & Mid$(code, i, 1) & "."
    Next i
End Function

' On automatic calculation the report values are already current; a workbook left on manual
' is recalculated first. Calculation is manual while writing and ALWAYS automatic afterwards
' (also after an error).
Private Sub BeginWork()
    Application.StatusBar = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    If Application.Calculation <> xlCalculationAutomatic Then Application.Calculate
    Application.Calculation = xlCalculationManual
End Sub

Private Sub EndWork()
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
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

Private Function RowOf(v As Variant) As Long
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then
        If Len(CStr(v)) > 0 Then RowOf = CLng(v)
    End If
End Function

Private Function TextOf(v As Variant) As String
    If Not IsError(v) Then TextOf = CStr(v)
End Function

Private Function IsReportName(shName As String) As Boolean
    Dim p As Long
    p = InStr(shName, "_")
    If p > 1 Then IsReportName = IsNumeric(Left$(shName, p - 1))
End Function

' Sheets named <number>_<text> (1_GIRIS, 2_YUKLER, 1_RAPOR, ...), in tab order.
Private Function ReportSheets() As Collection
    Dim ws As Worksheet
    Set ReportSheets = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If IsReportName(ws.Name) Then ReportSheets.Add ws
    Next ws
End Function
