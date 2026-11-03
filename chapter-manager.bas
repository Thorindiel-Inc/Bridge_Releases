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
'    - exports: e.g. Phi3 of span n -> INPUT row 92 column B+n-1 (G_PHI3)
'    - then the ToC and figure lists (RAPOR_LISTE) are rebuilt
'  Report sheet = name starts with a digit and does not end with _Y / _YATAY
'  (landscape companions are link targets of the main sheet's placeholders).
' ============================================================================

Private Const SCRIPT_ID As String = "chapter-manager"
Private Const SCRIPT_VERSION As String = "2026-11-03d"
Private Const SCRIPT_CHANGELOG As String = "Copies named after their source sheet (fixes swapped names when _Y stood before the main sheet); removes broken #REF! names (no replicated junk names, no conflict prompts); landscape companion suffix _Y (owner rename; copies and deletions include it); no name-conflict prompts while copying; creates / deletes per-span chapter sheets (3A<n>_DYN + 3A<n>_DYN_YATAY) from the filled spans, orders them, chains table / figure / page numbers, Phi3 row, rebuilds ToC and figure lists"

Private Const MSG_TITLE As String = "BOLUM_YONETICI"
' chapter types: code | suffixes (main first, ';') | heading code cell | export: cell>INPUT row (column B = span 1)
' order of the entries = order of the chapters in the workbook
Private Const REGISTRY As String = "3A|DYN;DYN_Y|Z6|Q122>92"
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
        MsgBox msg & vbCrLf & SCRIPT_ID & " " & SCRIPT_VERSION, IIf(Left$(msg, 4) = "Done", vbInformation, vbExclamation), MSG_TITLE
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
        f = Split(reg(k), "|"): code = f(0): sfx = Split(f(1), ";")
        Set have = ExistingSpans(wb, code, sfx(0))
        If have.Count = 0 Then Err.Raise vbObjectError + 2, , "Chapter " & code & ": no template sheet " & code & "1_" & sfx(0) & " in the workbook."
        tmpl = have(1)
        For Each idx In spans
            If Not InColl(have, CLng(idx)) Then CopySet wb, code, sfx, tmpl, CLng(idx), f(2)
        Next idx
        For Each idx In have
            If Not InColl(spans, CLng(idx)) Then
                For j = LBound(sfx) To UBound(sfx)
                    If SheetExists(wb, code & idx & "_" & sfx(j)) Then toDel.Add code & idx & "_" & sfx(j)
                Next j
            End If
        Next idx
    Next k
    If toDel.Count > 0 Then
        For Each s In toDel: delList = delList & "  " & s & vbCrLf: Next s
        If Not scripted Then
            Application.ScreenUpdating = True
            If MsgBox("These sheets belong to spans that are no longer filled and will be DELETED:" & vbCrLf & delList & vbCrLf & "Continue?", _
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
    Dim names() As Variant, j As Long, n As Long, before As Long, i As Long, ws As Worksheet
    ReDim names(0 To UBound(sfx))
    For j = 0 To UBound(sfx)
        If SheetExists(wb, code & tmpl & "_" & sfx(j)) Then names(n) = code & tmpl & "_" & sfx(j): n = n + 1
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
        ws.Name = Replace(CStr(names(j)), code & tmpl & "_", code & idx & "_", 1, 1)
        If j = 0 Then
            ws.Range(H_IDX).Value = idx
            If codeCell <> "" Then ws.Range(codeCell).Value = code & idx
        End If
    Next i
    msgLog = msgLog & "Created span " & idx & " of " & code & " (copy of span " & tmpl & ")." & vbCrLf
End Sub

' span numbers of a chapter that have their main sheet, ascending
Private Function ExistingSpans(wb As Workbook, ByVal code As String, ByVal mainSfx As String) As Collection
    Dim i As Long
    Set ExistingSpans = New Collection
    For i = 1 To MAX_SPANS
        If SheetExists(wb, code & i & "_" & mainSfx) Then ExistingSpans.Add i
    Next i
End Function

' ---- tab order ----
Private Sub OrderSheets(wb As Workbook)
    Dim reg As Variant, k As Long, f() As String, sfx() As String, i As Long, j As Long
    Dim anchor As Worksheet, ws As Worksheet, nm As String
    ' anchor = last report sheet that is not managed
    For Each ws In wb.Worksheets
        If IsReportName(ws.Name) And ManagedCode(ws.Name) = "" Then Set anchor = ws
    Next ws
    If anchor Is Nothing Then Err.Raise vbObjectError + 3, , "No fixed report sheet (e.g. 2_YUKLER) to place the chapters after."
    reg = Split(REGISTRY, ",")
    For k = LBound(reg) To UBound(reg)
        f = Split(reg(k), "|"): sfx = Split(f(1), ";")
        For i = 1 To MAX_SPANS
            For j = 0 To UBound(sfx)
                nm = f(0) & i & "_" & sfx(j)
                If SheetExists(wb, nm) Then
                    If wb.Worksheets(nm).Index <> anchor.Index + 1 Then wb.Worksheets(nm).Move After:=anchor
                    Set anchor = wb.Worksheets(nm)
                End If
            Next j
        Next i
    Next k
End Sub

' ---- numbering chain ----
Private Sub WriteChains(wb As Workbook)
    Dim ws As Worksheet, prev As Collection, last As Worksheet, ft As String, inp As Worksheet
    Set prev = New Collection
    For Each ws In wb.Worksheets
        If IsReportName(ws.Name) Then
            If ManagedCode(ws.Name) <> "" And Not last Is Nothing Then
                ws.Range("CF1:CI1").Value = Array("Span", "Prev. last table", "Prev. last figure", "Prev. last page")
                ws.Range(H_TAB).Formula = MaxNum(prev, "T")
                ws.Range(H_FIG).Formula = MaxNum(prev, ChrW(350))
                ft = LastFooter(last)
                ws.Range(H_PAGE).Formula = "=" & ft
            End If
            prev.Add ws
            Set last = ws
        End If
    Next ws
    If last Is Nothing Then Exit Sub
    Set inp = wb.Worksheets("INPUT")
    inp.Range("M2").Formula = MaxNum(prev, "T")
    inp.Range("N2").Formula = MaxNum(prev, ChrW(350))
    inp.Range("O2").Formula = "=" & LastFooter(last)
End Sub

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
    Dim pa As String, r As Range, a As Range, bot As Long
    pa = Replace(ws.PageSetup.PrintArea, ";", ",")
    If pa = "" Then Err.Raise vbObjectError + 4, , ws.Name & ": no print area."
    Set r = ws.Range(pa)
    For Each a In r.Areas
        If a.Row + a.Rows.Count - 1 > bot Then bot = a.Row + a.Rows.Count - 1
    Next a
    LastFooter = "'" & Replace(ws.Name, "'", "''") & "'!$W$" & (bot - 1)
End Function

' ---- exports (e.g. Phi3 per span) ----
Private Sub WriteExports(wb As Workbook)
    Dim reg As Variant, k As Long, f() As String, sfx() As String, ex() As String, rowNo As Long, i As Long, inp As Worksheet, nm As String
    Set inp = wb.Worksheets("INPUT")
    reg = Split(REGISTRY, ",")
    For k = LBound(reg) To UBound(reg)
        f = Split(reg(k), "|"): sfx = Split(f(1), ";")
        If UBound(f) >= 3 Then
            If f(3) <> "" Then
                ex = Split(f(3), ">"): rowNo = CLng(ex(1))
                If Trim$(CStr(inp.Cells(rowNo, 1).Value)) = "" Then inp.Cells(rowNo, 1).Value = ChrW(934) & "3"
                For i = 1 To MAX_SPANS
                    nm = f(0) & i & "_" & sfx(0)
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
    IsReportName = True
End Function

' landscape companion sheets: name ends with _Y or _YATAY (owner 2026-11-03)
Private Function IsCompanionName(ByVal nm As String) As Boolean
    IsCompanionName = (UCase$(Right$(nm, 2)) = "_Y" Or UCase$(Right$(nm, 6)) = "_YATAY")
End Function

' chapter code if the sheet is a managed chapter sheet, else ""
Private Function ManagedCode(ByVal nm As String) As String
    Dim reg As Variant, k As Long, f() As String, sfx() As String, i As Long, j As Long
    reg = Split(REGISTRY, ",")
    For k = LBound(reg) To UBound(reg)
        f = Split(reg(k), "|"): sfx = Split(f(1), ";")
        For i = 1 To MAX_SPANS
            For j = 0 To UBound(sfx)
                If nm = f(0) & i & "_" & sfx(j) Then ManagedCode = f(0): Exit Function
            Next j
        Next i
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
