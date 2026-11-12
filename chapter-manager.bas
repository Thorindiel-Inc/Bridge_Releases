Attribute VB_Name = "BOLUM_YONETICI"
Option Explicit
' ============================================================================
'  BOLUM_YONETICI - per-span report chapters of the 1-2 workbook
'
'  Button 'BOLUMLERI GUNCELLE' on INPUT (OnAction BOLUM_YONETICI.BOLUMLERI_GUNCELLE).
'  For every chapter type in REGISTRY one set of sheets per FILLED span of INPUT
'  G_HESAP_ACIKLIK (B11:B16): sheet name = <code><span>_<suffix>, e.g. 3A1_DYN +
'  3A1_DYN_Y for span 1, 3A2_DYN + 3A2_DYN_Y for span 2 (owner 2026-11-03; companion renamed _Y by the owner).
'    - missing spans: the lowest existing set (the template) is copied as one set
'      (references between its sheets follow), renamed, span index CF2 and the
'      heading code cell set; typed values are copied as they are (owner)
'    - spans no longer filled: their sheets are deleted after a confirmation
'      (the last remaining set of a chapter is never deleted)
'    - tab order: managed sheets right after the last fixed report sheet
'      (2_YUKLER), chapter by chapter in REGISTRY order, span by span
'    - numbering chain: on every main sheet CG2 / CH2 / CI2 = last table / figure /
'      page before it (formulas over the previous report sheets); INPUT M2:O2
'      (G_SON_NUMARA) = the end of the last report sheet
'    - whole-bridge chapters (WHOLE_SHEETS): one set of sheets, not per span: never copied or
'      deleted, placed after the span chapters and chained like them; WHOLE_END = its last page cell
'    - exports: e.g. Phi3 of span n -> INPUT row 92 column B+n-1 (G_PHI3)
'    - then the ToC and figure lists (RAPOR_LISTE) are rebuilt
'  Report sheet = name starts with a digit and does not end with _Y / _YATAY
'  (landscape companions are link targets of the main sheet's placeholders).
' ============================================================================

Private Const SCRIPT_ID As String = "chapter-manager"
Private Const SCRIPT_VERSION As String = "2026-11-12b"
Private Const SCRIPT_CHANGELOG As String = "Its macro-window row is declared in this module ('@TOOL line), so a changed or new tool needs no PANEL release; Cells read and written found by M_ names (MIDAS_CORE McAnchor), created at the old address when missing; RTI_SONUC stays right after SPECTRUM; inputs and messages through the macro window; One span: headings 3, 4, 5, 6, 7 without a letter; more spans: 3A, 3B ...; chapters 3 and 4 per span with a letter; KILIF sheets stay next to SPECTRUM; several heading cells per chapter, support sheets are not report sheets"
Private Const SCRIPT_NEEDS As String = "midas-core"   ' UPDATE adds this module when the workbook lacks it
' macro-window row(s) of this module (PANEL reads them): order|group|title|macro|sheet|text|fields[|modules]
'@TOOL 40|Girdiler|A\u00E7\u0131kl\u0131k b\u00F6l\u00FCmlerini g\u00FCncelle|BOLUMLERI_GUNCELLE||Dolu a\u00E7\u0131kl\u0131k say\u0131s\u0131na g\u00F6re 3. ve 4. b\u00F6l\u00FCm sayfalar\u0131n\u0131 ekler veya siler; sayfa s\u0131ras\u0131n\u0131, numaralar\u0131, ba\u015Fl\u0131klar\u0131 ve listeleri g\u00FCnceller.|

Private Const MSG_TITLE As String = "BOLUM_YONETICI"
' chapter types: code | suffixes (main first, ';') | heading code cells ('cell' or 'cell>code', ';') | export: cell>INPUT row (column B = span 1)
' WHOLE = the place of the whole-bridge sheets (WHOLE_SHEETS) in the chapter order
' a one-character code gets a letter per span (3A, 3B ...; 4A, 4B ...), see Tag
' order of the entries = order of the chapters in the workbook
Private Const REGISTRY As String = "3|DYN;DYN_Y|Z6|Q122>92,WHOLE,4|KDME;KDME_Y;KILIF|Z6;Z904>5;Z1134>6;Z1318>7|"
Private Const WHOLE_SHEETS As String = "3B_RTI;3B_RTI_Y"
Private Const WHOLE_END As String = "3B_RTI>3B_RTI_Y!$AI$95"   ' last print piece of 3B_RTI is a placeholder row: its last page is on the landscape sheet
Private Const SPAN_NAME As String = "G_HESAP_ACIKLIK"
Private Const MAX_SPANS As Long = 6
Private Const H_IDX As String = "CF2"      ' helper cells on every main chapter sheet
Private Const H_TAB As String = "CG2"
Private Const H_FIG As String = "CH2"
Private Const H_PAGE As String = "CI2"

Private scripted As Boolean
Private msgLog As String

' ============================================================================
Public Sub BOLUMLERI_GUNCELLE()
    scripted = False
    RunSync ""
End Sub

' scripted test entry: no dialogs (deletions without asking), report to a text file
Public Sub BOLUMLERI_GUNCELLE_RUN(ByVal reportFile As String)
    scripted = True
    RunSync reportFile
End Sub

Private Sub RunSync(ByVal reportFile As String)
    Dim wb As Workbook, spans As Collection, ok As Boolean
    On Error GoTo Fail
    Set wb = ThisWorkbook
    msgLog = ""
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Set spans = FilledSpans(wb)
    PurgeBrokenNames wb
    OrderSheets wb                       ' tabs in canonical order before any copy
    If spans.Count = 0 Then Err.Raise vbObjectError + 1, , "No span filled in " & SPAN_NAME & "."
    ok = SyncChapters(wb, spans)
    If ok Then
        OrderSheets wb
        WriteHeadingCodes wb, spans.Count
        Application.Calculation = xlCalculationAutomatic
        WriteChains wb
        WriteExports wb
        Application.Calculation = xlCalculationAutomatic
        Application.CalculateFull
        RunLists wb
    End If
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Done reportFile, IIf(ok, "Done: " & spans.Count & " span(s)." & vbCrLf, "Stopped.") & msgLog
    Exit Sub
Fail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Done reportFile, "FAIL: " & Err.Description & vbCrLf & msgLog
End Sub

Private Sub Done(ByVal reportFile As String, ByVal msg As String)
    Dim f As Integer
    If reportFile <> "" Then
        f = FreeFile
        Open reportFile For Output As #f
        Print #f, msg
        Close #f
    Else
        UiMsg msg & vbCrLf & SCRIPT_ID & " " & SCRIPT_VERSION, IIf(Left$(msg, 4) = "Done", vbInformation, vbExclamation), MSG_TITLE
    End If
End Sub

' filled span positions (1..6) of G_HESAP_ACIKLIK
Private Function FilledSpans(wb As Workbook) As Collection
    Dim r As Range, i As Long, v As Variant
    Set FilledSpans = New Collection
    Set r = wb.Names(SPAN_NAME).RefersToRange
    For i = 1 To r.Cells.Count
        If i > MAX_SPANS Then Exit For
        v = r.Cells(i).Value
        If Not IsError(v) And Not IsEmpty(v) Then
            If IsNumeric(v) Then
                If CDbl(v) > 0 Then FilledSpans.Add i
            End If
        End If
    Next i
End Function

' ---- create / delete ----
Private Function SyncChapters(wb As Workbook, spans As Collection) As Boolean
    Dim reg As Variant, k As Long, f() As String, code As String, sfx() As String
    Dim have As Collection, idx As Variant, tmpl As Long, delList As String, toDel As Collection, s As Variant, j As Long
    reg = Split(REGISTRY, ",")
    Set toDel = New Collection
    For k = LBound(reg) To UBound(reg)
        If reg(k) = "WHOLE" Then GoTo NextChapter
        f = Split(reg(k), "|"): code = f(0): sfx = Split(f(1), ";")
        Set have = ExistingSpans(wb, code, sfx(0))
        If have.Count = 0 Then Err.Raise vbObjectError + 2, , "Chapter " & code & ": no template sheet " & Tag(code, 1) & "_" & sfx(0) & " in the workbook."
        tmpl = have(1)
        For Each idx In spans
            If Not InColl(have, CLng(idx)) Then CopySet wb, code, sfx, tmpl, CLng(idx), f(2)
        Next idx
        For Each idx In have
            If Not InColl(spans, CLng(idx)) Then
                For j = LBound(sfx) To UBound(sfx)
                    If SheetExists(wb, Tag(code, idx) & "_" & sfx(j)) Then toDel.Add Tag(code, idx) & "_" & sfx(j)
                Next j
            End If
        Next idx
NextChapter:
    Next k
    If toDel.Count > 0 Then
        For Each s In toDel: delList = delList & "  " & s & vbCrLf: Next s
        If Not scripted Then
            Application.ScreenUpdating = True
            If UiMsg("These sheets belong to spans that are no longer filled and will be DELETED:" & vbCrLf & delList & vbCrLf & "Continue?", _
                      vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) <> vbYes Then
                msgLog = msgLog & "Nothing deleted (cancelled)." & vbCrLf
                SyncChapters = False
                Exit Function
            End If
            Application.ScreenUpdating = False
        End If
        Application.DisplayAlerts = False
        For Each s In toDel: wb.Worksheets(CStr(s)).Delete: Next s
        Application.DisplayAlerts = True
        msgLog = msgLog & "Deleted:" & vbCrLf & delList
    End If
    SyncChapters = True
End Function

Private Sub CopySet(wb As Workbook, ByVal code As String, sfx() As String, ByVal tmpl As Long, ByVal idx As Long, ByVal codeCell As String)
    Dim names() As Variant, j As Long, n As Long, before As Long, i As Long, ws As Worksheet, part As Variant, q As Variant
    ReDim names(0 To UBound(sfx))
    For j = 0 To UBound(sfx)
        If SheetExists(wb, Tag(code, tmpl) & "_" & sfx(j)) Then names(n) = Tag(code, tmpl) & "_" & sfx(j): n = n + 1
    Next j
    ReDim Preserve names(0 To n - 1)
    before = wb.Worksheets.Count
    ' Excel asks about every name of the copied sheets that also exists in the workbook (e.g. old #REF!
    ' names); with alerts off it keeps the existing names - the same as answering Yes (owner 2026-11-03)
    Application.DisplayAlerts = False
    wb.Worksheets(names).Copy After:=wb.Worksheets(wb.Worksheets.Count)
    Application.DisplayAlerts = True
    ' the copies are the last n sheets, named '<source> (2)' by Excel and placed in the SOURCES' TAB
    ' order (not the array order) - so each copy is named after its source (owner's file had _Y first)
    Dim src As String, p As Long
    For i = before + 1 To wb.Worksheets.Count
        Set ws = wb.Worksheets(i)
        src = ws.Name
        p = InStrRev(src, " (")
        If p > 0 Then src = Left$(src, p - 1)
        For j = 0 To n - 1
            If StrComp(src, CStr(names(j)), vbTextCompare) = 0 Then Exit For
        Next j
        If j > n - 1 Then Err.Raise vbObjectError + 5, , "Copy of '" & ws.Name & "' does not match a template sheet."
        ws.Name = Replace(CStr(names(j)), Tag(code, tmpl) & "_", Tag(code, idx) & "_", 1, 1)
        If j = 0 Then
            Chain(ws).Cells(2, 1).Value = idx
            For Each part In Split(codeCell, ";")
                If part <> "" Then
                    q = Split(part, ">")
                    ws.Range(q(0)).Value = Tag(IIf(UBound(q) > 0, q(UBound(q)), code), idx)
                End If
            Next part
        End If
    Next i
    msgLog = msgLog & "Created span " & idx & " of " & code & " (copy of span " & tmpl & ")." & vbCrLf
End Sub

' span numbers of a chapter that have their main sheet, ascending
Private Function ExistingSpans(wb As Workbook, ByVal code As String, ByVal mainSfx As String) As Collection
    Dim i As Long
    Set ExistingSpans = New Collection
    For i = 1 To MAX_SPANS
        If SheetExists(wb, Tag(code, i) & "_" & mainSfx) Then ExistingSpans.Add i
    Next i
End Function

' ---- tab order ----
Private Sub OrderSheets(wb As Workbook)
    Dim reg As Variant, k As Long, f() As String, sfx() As String, i As Long, j As Long
    Dim anchor As Worksheet, ws As Worksheet, nm As String, side As Worksheet
    If SheetExists(wb, "SPECTRUM") Then Set side = wb.Worksheets("SPECTRUM")
    ' the RTI result sheet sits right after SPECTRUM, the calculation sheets follow it (owner 2026-11-09)
    If Not side Is Nothing And SheetExists(wb, "RTI_SONUC") Then
        If wb.Worksheets("RTI_SONUC").Index <> side.Index + 1 Then wb.Worksheets("RTI_SONUC").Move After:=side
        Set side = wb.Worksheets("RTI_SONUC")
    End If
    ' anchor = last report sheet that is not managed
    For Each ws In wb.Worksheets
        If IsReportName(ws.Name) And ManagedCode(ws.Name) = "" Then Set anchor = ws
    Next ws
    If anchor Is Nothing Then Err.Raise vbObjectError + 3, , "No fixed report sheet (e.g. 2_YUKLER) to place the chapters after."
    reg = Split(REGISTRY, ",")
    Dim w As Variant
    For k = LBound(reg) To UBound(reg)
        If reg(k) = "WHOLE" Then
            For Each w In Split(WHOLE_SHEETS, ";")
                If SheetExists(wb, CStr(w)) Then
                    If wb.Worksheets(CStr(w)).Index <> anchor.Index + 1 Then wb.Worksheets(CStr(w)).Move After:=anchor
                    Set anchor = wb.Worksheets(CStr(w))
                End If
            Next w
            GoTo NextChapter
        End If
        f = Split(reg(k), "|"): sfx = Split(f(1), ";")
        For i = 1 To MAX_SPANS
            For j = 0 To UBound(sfx)
                nm = Tag(f(0), i) & "_" & sfx(j)
                If SheetExists(wb, nm) Then
                    If IsSupportName(nm) Then
                        ' calculation sheets stay beside SPECTRUM (owner 2026-11-06)
                        If Not side Is Nothing Then
                            If wb.Worksheets(nm).Index <> side.Index + 1 Then wb.Worksheets(nm).Move After:=side
                            Set side = wb.Worksheets(nm)
                        End If
                    Else
                        If wb.Worksheets(nm).Index <> anchor.Index + 1 Then wb.Worksheets(nm).Move After:=anchor
                        Set anchor = wb.Worksheets(nm)
                    End If
                End If
            Next j
        Next i
NextChapter:
    Next k
End Sub

' heading codes: one filled span = plain chapter number (3, 4, 5 ...), more spans = letter per span
' (3A, 3B ...; 4A, 4B ...; 5A, 5B ...) (owner 2026-11-06)
Private Sub WriteHeadingCodes(wb As Workbook, ByVal nSpans As Long)
    Dim reg As Variant, k As Long, f() As String, sfx() As String, i As Long, part As Variant, q As Variant, base As String, nm As String
    reg = Split(REGISTRY, ",")
    For k = LBound(reg) To UBound(reg)
        If reg(k) = "WHOLE" Then GoTo NextChapter
        f = Split(reg(k), "|"): sfx = Split(f(1), ";")
        For i = 1 To MAX_SPANS
            nm = Tag(f(0), i) & "_" & sfx(0)
            If SheetExists(wb, nm) Then
                For Each part In Split(f(2), ";")
                    If part <> "" Then
                        q = Split(part, ">")
                        base = IIf(UBound(q) > 0, q(UBound(q)), f(0))
                        wb.Worksheets(nm).Range(q(0)).Value = IIf(nSpans = 1, base, Tag(base, i))
                    End If
                Next part
            End If
        Next i
NextChapter:
    Next k
End Sub

' ---- numbering chain ----
Private Sub WriteChains(wb As Workbook)
    Dim ws As Worksheet, prev As Collection, last As Worksheet, ft As String, inp As Worksheet
    Set prev = New Collection
    For Each ws In wb.Worksheets
        If IsReportName(ws.Name) Then
            If ManagedCode(ws.Name) <> "" And Not last Is Nothing Then
                Chain(ws).Rows(1).Value = Array("Span", "Prev. last table", "Prev. last figure", "Prev. last page")
                Chain(ws).Cells(2, 2).Formula = MaxNum(prev, "T")
                Chain(ws).Cells(2, 3).Formula = MaxNum(prev, ChrW(350))
                ft = LastFooter(last)
                Chain(ws).Cells(2, 4).Formula = "=" & ft
            End If
            prev.Add ws
            Set last = ws
        End If
    Next ws
    If last Is Nothing Then Exit Sub
    Set inp = wb.Worksheets("INPUT")
    With McAnchor(inp, "G_SON_NUMARA", "M2:O2")    ' the numbers the next workbook (8_GEN) continues from
        .Cells(1, 1).Formula = MaxNum(prev, "T")
        .Cells(1, 2).Formula = MaxNum(prev, ChrW(350))
        .Cells(1, 3).Formula = "=" & LastFooter(last)
    End With
End Sub

' the chain cells of a chapter sheet: name M_ZINCIR of that sheet on the value row (CF2:CI2 of the
' original layout); returned with the label row above it: row 1 labels, row 2 span index /
' last table / last figure / last page of the chapters before
Private Function Chain(ws As Worksheet) As Range
    Set Chain = McAnchor(ws, "M_ZINCIR", H_IDX & ":" & H_PAGE, True).Offset(-1, 0).Resize(2, 4)
End Function

' =MAX(IFERROR(AGGREGATE(14,6,'s'!$AF$1:$AF$5000/('s'!$AE$1:$AE$5000="T"),1),0), ...)
Private Function MaxNum(sheets As Collection, ByVal kind As String) As String
    Dim ws As Variant, s As String, q As String
    For Each ws In sheets
        q = "'" & Replace(ws.Name, "'", "''") & "'!"
        s = s & IIf(s = "", "", ",") & "IFERROR(AGGREGATE(14,6," & q & "$AF$1:$AF$5000/(" & q & "$AE$1:$AE$5000=""" & kind & """),1),0)"
    Next ws
    MaxNum = "=MAX(" & IIf(s = "", "0", s) & ")"
End Function

' footer cell of the last print piece: column W, one row above the piece's bottom row
Private Function LastFooter(ws As Worksheet) As String
    Dim pa As String, p As Variant, v As Long, bot As Long, e() As String, c() As String
    e = Split(WHOLE_END, ">")
    If ws.Name = e(0) Then
        ' the whole-bridge chapter ends on its landscape sheet: name M_RTI_SON_SAYFA (AI95 of the original layout)
        c = Split(e(1), "!")
        LastFooter = "'" & c(0) & "'!" & McAnchor(ws.Parent.Worksheets(c(0)), "M_RTI_SON_SAYFA", c(1)).Address
        Exit Function
    End If
    pa = Replace(ws.PageSetup.PrintArea, ";", ",")
    If pa = "" Then Err.Raise vbObjectError + 4, , ws.Name & ": no print area."
    ' piece by piece: Range() refuses a text over 255 characters (many print pieces)
    For Each p In Split(pa, ",")
        v = CLng(Val(Mid$(p, InStrRev(p, "$") + 1)))
        If v > bot Then bot = v
    Next p
    LastFooter = "'" & Replace(ws.Name, "'", "''") & "'!$W$" & (bot - 1)
End Function

' ---- exports (e.g. Phi3 per span) ----
Private Sub WriteExports(wb As Workbook)
    Dim reg As Variant, k As Long, f() As String, sfx() As String, ex() As String, rowNo As Long, i As Long, inp As Worksheet, nm As String
    Set inp = wb.Worksheets("INPUT")
    reg = Split(REGISTRY, ",")
    For k = LBound(reg) To UBound(reg)
        If reg(k) = "WHOLE" Then GoTo NextChapter
        f = Split(reg(k), "|"): sfx = Split(f(1), ";")
        If UBound(f) >= 3 Then
            If f(3) <> "" Then
                ex = Split(f(3), ">"): rowNo = CLng(ex(1))
                If Trim$(CStr(inp.Cells(rowNo, 1).Value)) = "" Then inp.Cells(rowNo, 1).Value = ChrW(934) & "3"
                For i = 1 To MAX_SPANS
                    nm = Tag(f(0), i) & "_" & sfx(0)
                    If SheetExists(wb, nm) Then
                        inp.Cells(rowNo, 1 + i).Formula = "='" & nm & "'!$" & ColOf(ex(0)) & "$" & RowOf(ex(0))
                    Else
                        inp.Cells(rowNo, 1 + i).ClearContents
                    End If
                Next i
                On Error Resume Next
                wb.Names("G_PHI3").Delete
                On Error GoTo 0
                wb.Names.Add Name:="G_PHI3", RefersTo:="=INPUT!$B$" & rowNo & ":$G$" & rowNo
            End If
        End If
NextChapter:
    Next k
End Sub

' ---- ToC / figure lists ----
Private Sub RunLists(wb As Workbook)
    Dim r As Variant
    On Error Resume Next
    r = Application.Run("'" & wb.Name & "'!RAPOR_LISTE.LISTELERI_AYARLA_SESSIZ")
    If Err.Number <> 0 Then
        msgLog = msgLog & "ToC / figure lists not rebuilt: " & Err.Description & vbCrLf: Err.Clear
    ElseIf CStr(r) <> "" Then
        msgLog = msgLog & "ToC / figure lists: " & CStr(r)
    Else
        msgLog = msgLog & "ToC and figure lists rebuilt." & vbCrLf
    End If
    On Error GoTo 0
End Sub

' Names that point to #REF! are dead (old matrix add-in names, stale _FilterDatabase of 3A ...). They
' are copied with every chapter sheet and make Excel ask about name conflicts, so they are removed
' (owner 2026-11-03: replicated names). Excel's own hidden _xlfn names (=#NAME?) are not touched.
Private Sub PurgeBrokenNames(wb As Workbook)
    Dim i As Long, nm As Name, n As Long, rt As String
    For i = wb.Names.Count To 1 Step -1
        Set nm = wb.Names(i)
        rt = ""
        On Error Resume Next
        rt = nm.RefersTo
        On Error GoTo 0
        If InStr(1, rt, "#REF!", vbBinaryCompare) > 0 And Left$(nm.Name, 6) <> "_xlfn." Then
            On Error Resume Next
            nm.Delete
            If Err.Number = 0 Then n = n + 1
            Err.Clear
            On Error GoTo 0
        End If
    Next i
    If n > 0 Then msgLog = msgLog & "Removed " & n & " broken (#REF!) name(s)." & vbCrLf
End Sub

' ---- helpers ----
Private Function IsReportName(ByVal nm As String) As Boolean
    If Len(nm) = 0 Then Exit Function
    If Not (Left$(nm, 1) Like "#") Then Exit Function
    If IsCompanionName(nm) Then Exit Function
    If IsSupportName(nm) Then Exit Function
    IsReportName = True
End Function

' chapter tag of span idx: 3 -> 3A, 3B ...; 4 -> 4A, 4B ... (owner 2026-11-06)
Private Function Tag(ByVal code As String, ByVal idx As Long) As String
    If Len(code) = 1 Then
        Tag = code & Chr$(64 + idx)
    Else
        Tag = code & idx
    End If
End Function

' calculation sheets of a chapter: they carry no report headings or captions
Private Function IsSupportName(ByVal nm As String) As Boolean
    Dim u As String: u = UCase$(nm)
    IsSupportName = (Right$(u, 6) = "_KILIF")
End Function

' landscape companion sheets: name ends with _Y or _YATAY (owner 2026-11-03)
Private Function IsCompanionName(ByVal nm As String) As Boolean
    IsCompanionName = (UCase$(Right$(nm, 2)) = "_Y" Or UCase$(Right$(nm, 6)) = "_YATAY")
End Function

' chapter code if the sheet is a managed chapter sheet, else ""
Private Function ManagedCode(ByVal nm As String) As String
    Dim reg As Variant, k As Long, f() As String, sfx() As String, i As Long, j As Long
    If InStr(";" & WHOLE_SHEETS & ";", ";" & nm & ";") > 0 Then ManagedCode = "WHOLE": Exit Function
    reg = Split(REGISTRY, ",")
    For k = LBound(reg) To UBound(reg)
        If reg(k) = "WHOLE" Then GoTo NextChapter
        f = Split(reg(k), "|"): sfx = Split(f(1), ";")
        For i = 1 To MAX_SPANS
            For j = 0 To UBound(sfx)
                If nm = Tag(f(0), i) & "_" & sfx(j) Then ManagedCode = f(0): Exit Function
            Next j
        Next i
NextChapter:
    Next k
End Function

Private Function SheetExists(wb As Workbook, ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Private Function InColl(c As Collection, ByVal v As Long) As Boolean
    Dim x As Variant
    For Each x In c
        If CLng(x) = v Then InColl = True: Exit Function
    Next x
End Function

Private Function ColOf(ByVal a As String) As String
    Dim i As Long
    For i = 1 To Len(a)
        If Mid$(a, i, 1) Like "#" Then ColOf = Left$(a, i - 1): Exit Function
    Next i
End Function

Private Function RowOf(ByVal a As String) As Long
    RowOf = CLng(Mid$(a, Len(ColOf(a)) + 1))
End Function
