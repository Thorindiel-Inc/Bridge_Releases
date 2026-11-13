Attribute VB_Name = "ORTA_AYAK"
Option Explicit
' Pier report workbook (8'7_...) - 2026-11-13a
' GRUP_SAYFALARI: one report sheet per pier group, named <group number>_AYAK (1_AYAK, 2_AYAK ...).
'   groups        = the typed Grup No column of the pier table (OA_GRUP_NO beside G_AYAK_TABLOSU);
'                   a pier without a group number is not in the report
'   missing sheet = copy of the very hidden template SABLON_AYAK, group number set in its chain cell;
'                   an existing group sheet is never copied again: the design values typed on it stay
'   removed group = its sheet is deleted after a confirmation
'   one group     = one section: type, En / Cap, Boy, Kalinlik and the column count of all its piers
'                   must be equal (heights may differ); the tool stops and names the piers otherwise
'   pages         = rows whose gate (M_AYAK_KAPI) is 0 are hidden, the print area is the open pages
'                   (M_AYAK_SAYFA = page index of every row, M_AYAK_BASKI = the printed columns)
'   numbers       = M_AYAK_ZINCIR (2 rows x 4): row 1 = group, last table / figure / page before the
'                   sheet, row 2 = the sheet's own last table / figure / page; the first group continues
'                   from G_ONCEKI_NUMARA, each next one from the sheet before it, G_SON_NUMARA = the last
'   then the contents and the figure / table lists are rebuilt (RAPOR_LISTE).
' GRUP_SAYFALARI_SESSIZ: the same without a message box (deletes without asking); returns the log.
' GRUP_SAYFALARI_RUN file: the same, log written to a text file (scripted tests).

Private Const SCRIPT_ID As String = "pier-sheets"
Private Const SCRIPT_VERSION As String = "2026-11-13a"
Private Const SCRIPT_CHANGELOG As String = "New: one report sheet per pier group from the template SABLON_AYAK; the pages of the group's section type are shown and printed, the others hidden; table, figure and page numbers run from group to group; contents and lists rebuilt"
' macro-window row(s) of this module (PANEL reads them): order|group|title|macro|sheet|text|fields[|modules]
'@TOOL 30|Girdiler|Ayak grubu sayfalar\u0131|ORTA_AYAK.GRUP_SAYFALARI||Ayak tablosundaki Grup No s\u00FCtununa g\u00F6re her ayak grubu i\u00E7in bir rapor sayfas\u0131 a\u00E7ar (var olan sayfalara dokunmaz), kesit tipine g\u00F6re sayfalar\u0131 g\u00F6sterir / gizler, yazd\u0131rma alan\u0131n\u0131, numaralar\u0131 ve listeleri yeniler.|

Private Const TEMPLATE As String = "SABLON_AYAK"
Private Const SHEET_SFX As String = "_AYAK"
Private Const AFTER_SHEET As String = "AYRAC"          ' the first group sheet goes right after this sheet
Private Const N_TABLE As String = "G_AYAK_TABLOSU"
Private Const N_COLS As String = "G_AYAK_KOLON"        ' optional
Private Const N_GROUP As String = "OA_GRUP_NO"
Private Const N_GTAB As String = "OA_GRUP_TABLO"
Private Const N_PREV As String = "G_ONCEKI_NUMARA"
Private Const N_LAST As String = "G_SON_NUMARA"
Private Const M_CHAIN As String = "M_AYAK_ZINCIR"      ' sheet-level names of the template, copied with it
Private Const M_GATE As String = "M_AYAK_KAPI"
Private Const M_PAGE As String = "M_AYAK_SAYFA"
Private Const M_PRINT As String = "M_AYAK_BASKI"
Private Const SEC_TYPES As String = "|KUTU|DAIRESEL|DIKDORTGEN|OVAL|2 KUTU|ALTIGEN|"

Private scripted As Boolean
Private lg As String

Public Sub GRUP_SAYFALARI()
    Dim s As String
    scripted = False
    s = RunSync()
    If Left$(s, 5) = "FAIL:" Then
        ShowMsg Mid$(s, 6), vbExclamation
    Else
        ShowMsg s, vbInformation
    End If
End Sub

Public Function GRUP_SAYFALARI_SESSIZ() As String
    scripted = True
    GRUP_SAYFALARI_SESSIZ = RunSync()
End Function

Public Sub GRUP_SAYFALARI_RUN(ByVal reportFile As String)
    Dim s As String, f As Integer
    scripted = True
    s = RunSync()
    f = FreeFile
    Open reportFile For Output As #f
    Print #f, s
    Close #f
End Sub

Private Function RunSync() As String
    Dim wb As Workbook, groups As Collection, k As Variant, ws As Worksheet, i As Long
    Dim toDel As Collection, delList As String, s As Variant, made As String, n As Long
    On Error GoTo Fail
    Set wb = ThisWorkbook
    lg = ""
    Application.ScreenUpdating = False
    Set groups = ReadGroups(wb)

    ' sheets of groups that no longer exist
    Set toDel = New Collection
    For Each ws In wb.Worksheets
        n = GroupOfSheet(ws.Name)
        If n > 0 Then
            If Not InColl(groups, n) Then toDel.Add ws.Name: delList = delList & "  " & ws.Name & vbLf
        End If
    Next ws
    If toDel.Count > 0 Then
        If Not scripted Then
            Application.ScreenUpdating = True
            If Ask(U("Bu sayfalar\u0131n grubu art\u0131k ayak tablosunda yok. Sayfalar ve \u00FCzerlerine yaz\u0131lm\u0131\u015F tasar\u0131m de\u011Ferleri S\u0130L\u0130NECEK:") & vbLf & delList & vbLf & U("Devam edilsin mi?")) <> vbYes Then
                RunSync = "FAIL:" & U("\u0130\u015Flem durduruldu, hi\u00E7bir \u015Fey de\u011Fi\u015Fmedi.")
                Exit Function
            End If
            Application.ScreenUpdating = False
        End If
    End If

    ' missing sheets first, so the workbook always keeps a visible sheet
    For Each k In groups
        If Not SheetExists(wb, CLng(k) & SHEET_SFX) Then
            MakeSheet wb, CLng(k)
            made = made & " " & CLng(k) & SHEET_SFX
        End If
    Next k
    If Len(made) > 0 Then lg = lg & U("Yeni sayfa:") & made & vbLf
    If toDel.Count > 0 Then
        Application.DisplayAlerts = False
        For Each s In toDel: wb.Worksheets(CStr(s)).Delete: Next s
        Application.DisplayAlerts = True
        lg = lg & U("Silinen sayfa:") & vbLf & delList
    End If

    OrderSheets wb, groups
    WriteChains wb, groups
    If Application.Calculation <> xlCalculationAutomatic Then Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    For Each k In groups
        Set ws = wb.Worksheets(CLng(k) & SHEET_SFX)
        n = ApplyGate(ws)
        lg = lg & ws.Name & ": " & GroupLabel(wb, CLng(k)) & ", " & n & U(" sayfa") & vbLf
    Next k
    Application.Calculate
    RunLists wb
    Application.ScreenUpdating = True
    RunSync = groups.Count & U(" ayak grubu.") & vbLf & lg
    Exit Function
Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    RunSync = "FAIL:" & Err.Description & IIf(Len(lg) > 0, vbLf & lg, "")
End Function

' group numbers in use, ascending; stops when a group number or a section type cannot be read or the
' piers of one group differ in section
Private Function ReadGroups(wb As Workbook) As Collection
    Dim t As Variant, g As Variant, kc As Variant, rc As Range, hasCols As Boolean
    Dim nG As Long, i As Long, k As Long, v As Variant, ty As String, d As String
    Dim first() As Long, bad As String, noGrp As String, pier As String, gap As Boolean
    t = NameRange(wb, N_TABLE).Value
    g = NameRange(wb, N_GROUP).Value
    If UBound(g, 1) <> UBound(t, 1) Then Err.Raise vbObjectError + 20, , N_GROUP & U(" ile ") & N_TABLE & U(" ayn\u0131 sat\u0131r say\u0131s\u0131nda de\u011Fil.")
    nG = NameRange(wb, N_GTAB).Rows.Count
    On Error Resume Next
    Set rc = wb.Names(N_COLS).RefersToRange
    On Error GoTo 0
    If Not rc Is Nothing Then
        kc = rc.Value
        hasCols = (UBound(kc, 1) = UBound(t, 1))
    End If
    ReDim first(1 To nG)
    For i = 1 To UBound(t, 1)
        v = g(i, 1)
        pier = CellText(t(i, 1))
        ty = NormType(t(i, 3))
        If IsError(v) Then Err.Raise vbObjectError + 21, , pier & U(": Grup No hata de\u011Feri.")
        If Len(Trim$(CStr(v))) > 0 Then
            k = 0
            If IsNumeric(v) Then
                If CDbl(v) = Int(CDbl(v)) And CDbl(v) >= 1 And CDbl(v) <= nG Then k = CLng(v)
            End If
            If k = 0 Then Err.Raise vbObjectError + 21, , pier & U(": Grup No 1 ile ") & nG & U(" aras\u0131nda bir tam say\u0131 olmal\u0131 ('") & CStr(v) & "')."
            If InStr(SEC_TYPES, "|" & ty & "|") = 0 Then Err.Raise vbObjectError + 22, , pier & U(": kesit tipi '") & CellText(t(i, 3)) & U("' tan\u0131nm\u0131yor (KUTU, DA\u0130RESEL, D\u0130KD\u00D6RTGEN, OVAL, 2 KUTU, ALTIGEN).")
            If first(k) = 0 Then
                first(k) = i
            Else
                d = SectionDiff(t, kc, hasCols, first(k), i)
                If Len(d) > 0 Then bad = bad & "  Grup " & k & ": " & CellText(t(first(k), 1)) & " / " & pier & " (" & d & ")" & vbLf
            End If
        ElseIf InStr(SEC_TYPES, "|" & ty & "|") > 0 Then
            noGrp = noGrp & " " & pier
        End If
    Next i
    If Len(bad) > 0 Then Err.Raise vbObjectError + 23, , U("Bir grubun ayaklar\u0131 ayn\u0131 kesitte olmal\u0131 (y\u00FCkseklik farkl\u0131 olabilir). Farkl\u0131 olanlar:") & vbLf & bad & U("Grup numaralar\u0131n\u0131 veya ayak tablosunu d\u00FCzeltin.")
    Set ReadGroups = New Collection
    For k = 1 To nG
        If first(k) > 0 Then
            If k > ReadGroups.Count + 1 Then gap = True
            ReadGroups.Add k
        End If
    Next k
    If ReadGroups.Count = 0 Then Err.Raise vbObjectError + 24, , U("Hi\u00E7bir aya\u011Fa grup numaras\u0131 verilmemi\u015F (INPUT, ayak tablosu, Grup No s\u00FCtunu).")
    If Len(noGrp) > 0 Then lg = lg & U("Uyar\u0131: grup numaras\u0131 olmayan ayaklar raporda yok:") & noGrp & vbLf
    If gap Then lg = lg & U("Uyar\u0131: grup numaralar\u0131 ard\u0131\u015F\u0131k de\u011Fil; ba\u015Fl\u0131k numaralar\u0131 grup numaras\u0131n\u0131 izler.") & vbLf
End Function

' what differs between two pier rows of one group ("" = same section)
Private Function SectionDiff(t As Variant, kc As Variant, ByVal hasCols As Boolean, ByVal a As Long, ByVal b As Long) As String
    Dim s As String
    If NormType(t(a, 3)) <> NormType(t(b, 3)) Then s = s & ", Tip"
    If Not SameVal(t(a, 4), t(b, 4)) Then s = s & U(", En / \u00C7ap")
    If Not SameVal(t(a, 5), t(b, 5)) Then s = s & ", Boy"
    If Not SameVal(t(a, 6), t(b, 6)) Then s = s & U(", Kal\u0131nl\u0131k")
    If hasCols Then
        If ColCount(kc(a, 1)) <> ColCount(kc(b, 1)) Then s = s & U(", Kolon Say\u0131s\u0131")
    End If
    If Len(s) > 0 Then SectionDiff = Mid$(s, 3)
End Function

Private Function SameVal(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsError(a) Or IsError(b) Then Exit Function
    If IsNumeric(a) And IsNumeric(b) And Not IsEmpty(a) And Not IsEmpty(b) Then
        SameVal = (Abs(CDbl(a) - CDbl(b)) <= 0.000000001 * (1 + Abs(CDbl(a))))
    Else
        SameVal = (Trim$(CStr(a)) = Trim$(CStr(b)))
    End If
End Function

Private Function ColCount(ByVal v As Variant) As Long
    ColCount = 1
    If IsError(v) Then Exit Function
    If IsNumeric(v) And Not IsEmpty(v) Then
        If CDbl(v) >= 2 Then ColCount = CLng(v)
    End If
End Function

Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Then CellText = "#" Else CellText = Trim$(CStr(v))
End Function

' section type in capitals without Turkish dots
Private Function NormType(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then Exit Function
    s = UCase$(Trim$(CStr(v)))
    s = Replace(s, ChrW(304), "I"): s = Replace(s, ChrW(305), "I")
    s = Replace(s, ChrW(214), "O"): s = Replace(s, ChrW(246), "O")
    s = Replace(s, ChrW(220), "U"): s = Replace(s, ChrW(252), "U")
    s = Replace(s, ChrW(286), "G"): s = Replace(s, ChrW(287), "G")
    NormType = s
End Function

Private Function GroupLabel(wb As Workbook, ByVal k As Long) As String
    Dim v As Variant
    v = NameRange(wb, N_GTAB).Cells(k, 2).Value
    GroupLabel = CellText(v) & " (" & CellText(NameRange(wb, N_GTAB).Cells(k, 3).Value) & ")"
End Function

' "<n>_AYAK" -> n, anything else -> 0
Private Function GroupOfSheet(ByVal nm As String) As Long
    Dim p As String
    If Len(nm) <= Len(SHEET_SFX) Then Exit Function
    If UCase$(Right$(nm, Len(SHEET_SFX))) <> SHEET_SFX Then Exit Function
    p = Left$(nm, Len(nm) - Len(SHEET_SFX))
    If Len(p) > 3 Then Exit Function
    If p Like String$(Len(p), "#") Then GroupOfSheet = CLng(p)
End Function

Private Sub MakeSheet(wb As Workbook, ByVal k As Long)
    Dim tpl As Worksheet, ws As Worksheet, vis As Long, e As Long, d As String, n0 As Long
    If Not SheetExists(wb, TEMPLATE) Then Err.Raise vbObjectError + 25, , U("\u015Eablon sayfa yok: ") & TEMPLATE
    Set tpl = wb.Worksheets(TEMPLATE)
    vis = tpl.Visible
    n0 = wb.Worksheets.Count
    On Error Resume Next
    tpl.Visible = xlSheetVisible
    Application.DisplayAlerts = False
    tpl.Copy After:=wb.Worksheets(n0)
    e = Err.Number: d = Err.Description
    If e = 0 And wb.Worksheets.Count = n0 + 1 Then
        Set ws = wb.Worksheets(n0 + 1)
        ws.Name = k & SHEET_SFX
        ws.Names(M_CHAIN).RefersToRange.Cells(1, 1).Value = k
        e = Err.Number: d = Err.Description
    ElseIf e = 0 Then
        e = 1: d = "copy"
    End If
    Err.Clear
    tpl.Visible = vis
    Application.DisplayAlerts = True
    On Error GoTo 0
    If e <> 0 Then Err.Raise vbObjectError + 26, , U("Grup sayfas\u0131 a\u00E7\u0131lamad\u0131 (") & k & SHEET_SFX & "): " & d
End Sub

Private Sub OrderSheets(wb As Workbook, groups As Collection)
    Dim k As Variant, ws As Worksheet, before As Worksheet
    On Error Resume Next
    Set before = wb.Worksheets(AFTER_SHEET)
    On Error GoTo 0
    For Each k In groups
        Set ws = wb.Worksheets(CLng(k) & SHEET_SFX)
        If Not before Is Nothing Then
            If ws.Index <> before.Index + 1 Then ws.Move After:=before
        End If
        Set before = ws
    Next k
End Sub

Private Sub WriteChains(wb As Workbook, groups As Collection)
    Dim k As Variant, ws As Worksheet, c As Range, lastC As Range, j As Long, f As String, r As Range
    For Each k In groups
        Set ws = wb.Worksheets(CLng(k) & SHEET_SFX)
        Set c = SheetRange(ws, M_CHAIN)
        If CStr(c.Cells(1, 1).Value) <> CStr(CLng(k)) Then c.Cells(1, 1).Value = CLng(k)
        For j = 2 To 4
            If lastC Is Nothing Then
                f = "=INDEX(" & N_PREV & "," & (j - 1) & ")"
            Else
                f = "='" & lastC.Worksheet.Name & "'!" & lastC.Cells(2, j).Address
            End If
            If c.Cells(1, j).Formula <> f Then c.Cells(1, j).Formula = f
        Next j
        Set lastC = c
    Next k
    Set r = NameRange(wb, N_LAST)
    For j = 1 To 3
        f = "='" & lastC.Worksheet.Name & "'!" & lastC.Cells(2, j + 1).Address
        If r.Cells(j).Formula <> f Then r.Cells(j).Formula = f
    Next j
End Sub

' hides the rows whose gate is 0, print area = the pages whose first row is open; returns the page count
Private Function ApplyGate(ws As Worksheet) As Long
    Dim rg As Range, rp As Range, rb As Range, g As Variant, p As Variant
    Dim n As Long, r0 As Long, i As Long, j As Long, o As Boolean, pg As Long, cur As Long, a As Long, addr As String
    Set rg = SheetRange(ws, M_GATE)
    Set rp = SheetRange(ws, M_PAGE)
    Set rb = SheetRange(ws, M_PRINT)
    g = rg.Value: p = rp.Value
    n = UBound(g, 1): r0 = rg.Row
    If UBound(p, 1) <> n Or rp.Row <> r0 Then Err.Raise vbObjectError + 27, , ws.Name & ": " & M_GATE & " / " & M_PAGE
    For i = 1 To n
        If IsError(g(i, 1)) Then Err.Raise vbObjectError + 28, , ws.Name & U(": sayfa kap\u0131s\u0131 hesaplanam\u0131yor (grup tablosunda bu grubun kesit tipi okunam\u0131yor).")
    Next i
    i = 1
    Do While i <= n
        o = (CDbl(g(i, 1)) <> 0)
        j = i
        Do While j < n
            If (CDbl(g(j + 1, 1)) <> 0) <> o Then Exit Do
            j = j + 1
        Loop
        ws.Rows(r0 + i - 1).Resize(j - i + 1).Hidden = Not o
        i = j + 1
    Loop
    For i = 1 To n + 1
        cur = 0
        If i <= n Then
            If IsNumeric(p(i, 1)) And Not IsEmpty(p(i, 1)) Then cur = CLng(p(i, 1))
        End If
        If cur <> pg Then
            If pg > 0 Then
                If CDbl(g(a, 1)) <> 0 Then
                    addr = addr & "," & Intersect(rb.EntireColumn, ws.Rows(r0 + a - 1).Resize(i - a)).Address
                    ApplyGate = ApplyGate + 1
                End If
            End If
            pg = cur: a = i
        End If
    Next i
    If ApplyGate = 0 Then Err.Raise vbObjectError + 29, , ws.Name & U(": a\u00E7\u0131k sayfa yok.")
    ws.PageSetup.PrintArea = Mid$(addr, 2)
End Function

Private Sub RunLists(wb As Workbook)
    Dim r As Variant
    On Error Resume Next
    r = Application.Run("'" & Replace(wb.Name, "'", "''") & "'!RAPOR_LISTE.LISTELERI_AYARLA_SESSIZ")
    If Err.Number <> 0 Then
        lg = lg & U("Listeler yenilenemedi: ") & Err.Description & vbLf
        Err.Clear
    Else
        lg = lg & U("\u0130\u00E7indekiler ve \u015Fekil / tablo listeleri yenilendi.") & vbLf
        If Len(CStr(r)) > 0 Then lg = lg & CStr(r)
    End If
    On Error GoTo 0
End Sub

Private Function NameRange(wb As Workbook, ByVal nm As String) As Range
    On Error Resume Next
    Set NameRange = wb.Names(nm).RefersToRange
    On Error GoTo 0
    If NameRange Is Nothing Then Err.Raise vbObjectError + 30, , U("Ad yok veya h\u00FCcre g\u00F6stermiyor: ") & nm
End Function

Private Function SheetRange(ws As Worksheet, ByVal nm As String) As Range
    On Error Resume Next
    Set SheetRange = ws.Names(nm).RefersToRange
    On Error GoTo 0
    If SheetRange Is Nothing Then Err.Raise vbObjectError + 31, , ws.Name & U(": sayfa ad\u0131 yok: ") & nm
    If SheetRange.Worksheet.Name <> ws.Name Then Err.Raise vbObjectError + 31, , ws.Name & ": " & nm & U(" ba\u015Fka sayfay\u0131 g\u00F6steriyor.")
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

' The macro window (module PANEL) is optional: without it, a plain message box.
Private Sub ShowMsg(ByVal msg As String, ByVal icon As Long)
    Dim title As String
    title = U("Ayak grubu sayfalar\u0131")
    On Error Resume Next
    Application.Run "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PANEL.UiMsg", msg, icon, title
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox msg, icon, title
    End If
End Sub

Private Function Ask(ByVal msg As String) As Long
    Dim title As String, r As Variant
    title = U("Ayak grubu sayfalar\u0131")
    On Error Resume Next
    r = Application.Run("'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PANEL.UiMsg", msg, vbYesNo + vbExclamation + vbDefaultButton2, title)
    If Err.Number <> 0 Then
        Err.Clear
        r = MsgBox(msg, vbYesNo + vbExclamation + vbDefaultButton2, title)
    End If
    Ask = CLng(r)
End Function

Private Function U(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, "\u")
    Do While p > 0
        s = Left$(s, p - 1) & ChrW(CLng("&H" & Mid$(s, p + 2, 4))) & Mid$(s, p + 6)
        p = InStr(p + 1, s, "\u")
    Loop
    U = s
End Function
