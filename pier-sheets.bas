Attribute VB_Name = "ORTA_AYAK"
Option Explicit
' Pier report workbook (8'7_...) - 2026-11-26a
' GRUP_SAYFALARI: one report sheet per pier group, named <group number>_<the group's pier runs> without spaces
'   (1_P1-P3,P16-P17; owner 2026-11-16), at most 31 characters (longer names are cut and end with an ellipsis),
'   <group number>_AYAK when the group has no name. A group sheet is found by the group number in the first
'   cell of its sheet-level name M_AYAK_ZINCIR, not by its name, so it is renamed whenever its piers change.
'   groups        = the typed Grup No column of the pier table (OA_GRUP_NO beside G_AYAK_TABLOSU);
'                   a pier without a group number is not in the report
'   group name    = column 2 of OA_GRUP_TABLO (OA_GRUP_ADI, runs of next rows of the pier table)
'   missing sheet = copy of the very hidden template SABLON_AYAK, group number set in its chain cell;
'                   an existing group sheet is never copied again: the design values typed on it stay
'   removed group = its sheet is deleted after a confirmation
'   renames       = two passes through free temporary names, so two groups that swap names never clash;
'                   the tool stops when a sheet that is not a group sheet already has the wanted name, or
'                   when two sheets hold the same group number
'   one group     = one section: type, En / Cap, Boy, Kalinlik and the column count of all its piers
'                   must be equal (heights may differ); the tool stops and names the piers otherwise
'   pages         = rows whose gate (M_AYAK_KAPI) is 0 are hidden, the print area is the open pages;
'                   pictures, text boxes and groups anchored on a closed row are hidden too (an object set to
'                   'move but don't size' would otherwise slide onto the next open page), shown again when it opens
'                   (M_AYAK_SAYFA = page index of every row, M_AYAK_BASKI = the printed columns)
'   numbers       = M_AYAK_ZINCIR (2 rows x 4): row 1 = group, last table / figure / page before the
'                   sheet, row 2 = the sheet's own last table / figure / page; the first group continues
'                   from G_ONCEKI_NUMARA, each next one from the sheet before it, G_SON_NUMARA = the last
'   data sheet    = one per group (owner 2026-11-13 / 2026-11-16), a copy of the very hidden SABLON_MIDAS (the
'                   former MIDAS_CIKTISI_KOLON), named MIDAS_<the group's pier runs without spaces> (MIDAS_P1,
'                   MIDAS_P2-P3,P16-P17; cut like the group sheets), placed in group order right before INPUT (they are
'                   inputs too; owner 2026-11-16), found by its group number in the sheet-level name M_TESIR_NO (a
'                   number, no cell) and renamed when the piers change; an older data sheet named MIDAS_<number>
'                   without that name is taken as the sheet of that group. The group sheet's formulas that read
'                   SABLON_MIDAS are pointed at its own data sheet (Tablo 63 'Ayak Tesirleri', Tablo 64, shear and
'                   torsion picks); a removed group's data sheet is deleted with its sheet.
'                   MIDAS_AYAK writes the column forces there; when the piers it read (M_TESIR_GRUP) are not the
'                   group's piers any more, the log says so. Without SABLON_MIDAS the data sheets are skipped.
'   Tablo 63      = 'Ayak Tesirleri' grows with its data (owner 2026-11-16): its last page (sheet-level name
'                   M_T63_SAYFA, the pattern page) is copied below itself as often as the rows of the group's data
'                   sheet need, and copies no longer needed are deleted; each copy continues the table rows (the
'                   R1C1 row offset of its data references), the page number chain (W) and the page index (CO);
'                   the page after the table gets its page number from the last copy; the number of copies is the
'                   sheet-level name M_T63_EK. Pages without rows are closed by their CF tag formula as before.
'   then the contents and the figure / table lists are rebuilt (RAPOR_LISTE; it lists only sheets named
'   <number>_<text> that do not end in _Y, _YATAY or _KILIF).
' SAYFALARI_YENILE_SESSIZ: gates, print areas, page numbers and lists of the existing group sheets again, nothing
'   made, renamed or deleted (MIDAS_AYAK calls it after writing forces, which can open a table page).
' GRUP_SAYFALARI_SESSIZ: the same without a message box (deletes without asking); returns the log.
' GRUP_SAYFALARI_RUN file: the same, log written to a text file (scripted tests).
' SABLON_GOSTER: shows the very hidden template SABLON_AYAK for editing with the rows and pictures of every section
'   type open; the next press closes them by the template's own gate and hides the sheet again.
' SABLON_UYGULA (owner 2026-11-16): a fix made once on SABLON_AYAK is carried to every group sheet. The very hidden
'   sheet SABLON_AYAK_IZ holds the template as it was at the last apply (as text); per cell:
'     template = that baseline                  -> nothing (a value typed on a group sheet stays)
'     template changed, group = baseline        -> the group cell takes the template cell, its references to
'                                                  SABLON_AYAK / SABLON_MIDAS pointed at the group's own sheets
'     template changed, group = template        -> nothing
'     template changed, group = something else  -> conflict: listed, written only when the owner says yes
'   Group formulas are compared in template terms. The Tablo 63 copies are taken out first (FitTablePages with 0
'   copies, so every reference is on the template rows again) and made again at the end with the gates, print
'   areas, numbers and lists (SAYFALARI_YENILE_SESSIZ). Not compared: M_AYAK_ZINCIR, the page index M_AYAK_SAYFA,
'   array formulas. A written cell also takes the template's number format; formats and row heights are not
'   touched. Pictures and drawings (16h, owner 2026-11-16: place and size only): one moved or resized on the
'   template since the last apply (list in SABLON_AYAK_IZ_R) puts the group sheet's own picture of the same id at
'   the same cell, offset and size; its image is never copied, so a group's own graph stays. New or deleted
'   pictures are listed, not carried. A run with picture changes only does not rebuild the Tablo 63 pages.
'   Without the baseline sheet the tool makes it from the template and changes nothing. After an
'   apply the baseline = the template, so a conflict that was kept is not asked again.
' SABLON_UYGULA_SESSIZ [doApply, force]: the same without questions (doApply False = list only); returns the log.
' VERI_SABLON_GOSTER / VERI_SABLON_UYGULA (16i, owner 2026-11-16): the same for the data sheets. SABLON_MIDAS is shown
'   for editing (the first show writes the baseline SABLON_MIDAS_IZ, so what is edited after it is carried); the
'   apply compares it with the baseline and writes each changed cell to every MIDAS_ data sheet whose cell still
'   equals the baseline; a cell the owner changed there (pasted GSD output) is listed and kept unless the owner says
'   yes. Formats and merged cells of the rectangle around the changed cells are pasted from the template (the new
'   layout), the MIDAS_AYAK cells (M_TESIR_*) are never compared. Addresses do not move, no page is rebuilt.
'   Pictures are not copied with the cells (17a): every data sheet keeps its own pictures where they are.
' VERI_SABLON_UYGULA_SESSIZ [doApply, force]: the same without questions; returns the log.

Private Const SCRIPT_ID As String = "pier-sheets"
Private Const SCRIPT_VERSION As String = "2026-11-26a"
Private Const SCRIPT_CHANGELOG As String = "A new group or data sheet is copied after the last visible sheet and found by its new name (with a very hidden last sheet the old copy took and renamed the wrong sheet); the data-template apply no longer copies the template's pictures onto the data sheets (each sheet keeps its own pictures; cells, formats and merged cells only); data-sheet template tools: show / hide SABLON_MIDAS and carry its changes (cells, formats, merged cells) to every MIDAS_ data sheet, three-way against the last apply; the template apply also carries the place and size of pictures moved or resized on SABLON_AYAK (the group sheet's own image stays); the shown template has every row and picture of every section type open (closed again as before when it is hidden); template tools: show / hide SABLON_AYAK and carry its changes to every group sheet (three-way against the last apply, cells changed by hand on a group sheet are listed and kept)"
' macro-window row(s) of this module (PANEL reads them): order|group|title|macro|sheet|text|fields[|modules]
'@TOOL 30|Girdiler|Ayak grubu sayfalar\u0131|ORTA_AYAK.GRUP_SAYFALARI||Ayak tablosundaki Grup No s\u00FCtununa g\u00F6re her ayak grubu i\u00E7in bir rapor sayfas\u0131 a\u00E7ar. Sayfa ad\u0131 grup numaras\u0131 ve grubun ayaklar\u0131d\u0131r (1_P1-P3,P16-P17); ayaklar de\u011Fi\u015Fince ad da de\u011Fi\u015Fir, sayfadaki de\u011Ferlere dokunulmaz. Kesit tipine g\u00F6re sayfalar\u0131 g\u00F6sterir / gizler, yazd\u0131rma alan\u0131n\u0131, numaralar\u0131 ve listeleri yeniler.|
'@TOOL 31|Girdiler|\u015Eablonu g\u00F6ster / gizle|ORTA_AYAK.SABLON_GOSTER||Gizli \u015Fablon sayfas\u0131n\u0131 (SABLON_AYAK) d\u00FCzenlemek i\u00E7in a\u00E7ar; tekrar bas\u0131l\u0131nca gizler. B\u00FCt\u00FCn grup sayfalar\u0131na gidecek bir d\u00FCzeltmeyi \u015Fablonda bir kez yap\u0131n, sonra '\u015Eablonu grup sayfalar\u0131na uygula'.|
'@TOOL 32|Girdiler|\u015Eablonu grup sayfalar\u0131na uygula|ORTA_AYAK.SABLON_UYGULA||\u015Eablonda son uygulamadan beri de\u011Fi\u015Fen h\u00FCcreleri her grup sayfas\u0131na yazar; form\u00FCller grubun kendi sayfas\u0131na ve veri sayfas\u0131na ba\u011Flan\u0131r. Grup sayfas\u0131nda elle de\u011Fi\u015Ftirdi\u011Finiz h\u00FCcreler listelenir ve yaln\u0131z siz evet derseniz yaz\u0131l\u0131r. Yazmadan \u00F6nce sayfa ba\u015F\u0131na ka\u00E7 h\u00FCcre de\u011Fi\u015Fece\u011Fini g\u00F6sterir.|
'@TOOL 33|Girdiler|Veri \u015Fablonunu g\u00F6ster / gizle|ORTA_AYAK.VERI_SABLON_GOSTER||Gizli veri \u015Fablonunu (SABLON_MIDAS: GSD girdileri ve MIDAS tesirleri sayfas\u0131) d\u00FCzenlemek i\u00E7in a\u00E7ar; tekrar bas\u0131l\u0131nca gizler. B\u00FCt\u00FCn MIDAS_ veri sayfalar\u0131na gidecek bir d\u00FCzen de\u011Fi\u015Fikli\u011Fini burada bir kez yap\u0131n, sonra 'Veri \u015Fablonunu veri sayfalar\u0131na uygula'.|
'@TOOL 34|Girdiler|Veri \u015Fablonunu veri sayfalar\u0131na uygula|ORTA_AYAK.VERI_SABLON_UYGULA||SABLON_MIDAS'ta son uygulamadan beri de\u011Fi\u015Fen h\u00FCcreleri, bi\u00E7imleri ve birle\u015Ftirilmi\u015F h\u00FCcreleri her MIDAS_ veri sayfas\u0131na ta\u015F\u0131r. Veri sayfas\u0131nda sizin de\u011Fi\u015Ftirdi\u011Finiz (yap\u0131\u015Ft\u0131rd\u0131\u011F\u0131n\u0131z) h\u00FCcreler listelenir ve yaln\u0131z siz evet derseniz yaz\u0131l\u0131r. MIDAS_AYAK tesir blo\u011Funa dokunulmaz. Yazmadan \u00F6nce sorar.|

Private Const TEMPLATE As String = "SABLON_AYAK"
Private Const SHEET_SFX As String = "_AYAK"            ' name of a group sheet whose group has no name
Private Const TMP_BASE As String = "~tmp_AYAK_"        ' temporary names during the renames
Private Const MAX_NAME As Long = 31                    ' Excel's limit for a sheet name
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
Private Const DATA_TPL As String = "SABLON_MIDAS"      ' very hidden template of a group's data sheet
Private Const DATA_PFX As String = "MIDAS_"            ' data sheet name = MIDAS_<pier runs>
Private Const M_DATA_NO As String = "M_TESIR_NO"       ' data sheet: sheet-level name holding its group number
Private Const DATA_BEFORE As String = "INPUT"          ' data sheets go right before this sheet
Private Const M_T63_PAGE As String = "M_T63_SAYFA"     ' group sheet: rows of the last Tablo 63 page (pattern)
Private Const M_T63_EXTRA As String = "M_T63_EK"       ' group sheet: number of pattern copies below it
Private Const M_BLOCK As String = "M_TESIR_BLOK"       ' data sheet: force block of MIDAS_AYAK (3rd column = names)
Private Const M_READ As String = "M_TESIR_GRUP"        ' data sheet: the group's piers when MIDAS_AYAK read them
Private Const BASE_SHAPES As String = "SABLON_AYAK_IZ_R"  ' very hidden: place and size of the template's pictures at the last apply
Private Const BASE_SHEET As String = "SABLON_AYAK_IZ"   ' very hidden: SABLON_AYAK at the last apply (as text)
Private Const DATA_BASE As String = "SABLON_MIDAS_IZ"  ' very hidden: SABLON_MIDAS at the last apply (as text)

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

' gates, print areas, page numbers and lists of the existing group sheets again; returns the log
Public Function SAYFALARI_YENILE_SESSIZ() As String
    Dim wb As Workbook, nG As Long, sh() As Worksheet, k As Long, n As Long, dsh() As Worksheet
    On Error GoTo Fail
    scripted = True
    lg = ""
    Set wb = ThisWorkbook
    Application.ScreenUpdating = False
    nG = NameRange(wb, N_GTAB).Rows.Count
    FindGroupSheets wb, nG, sh
    FindDataSheets wb, nG, dsh
    For k = 1 To nG
        If Not sh(k) Is Nothing And Not dsh(k) Is Nothing Then lg = lg & FitTablePages(sh(k), dsh(k))
    Next k
    If Application.Calculation <> xlCalculationAutomatic Then Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    For k = 1 To nG
        If Not sh(k) Is Nothing Then
            n = ApplyGate(sh(k))
            lg = lg & sh(k).Name & ": " & n & U(" sayfa") & vbLf
        End If
    Next k
    Application.Calculate
    RunLists wb
    Application.ScreenUpdating = True
    SAYFALARI_YENILE_SESSIZ = lg
    Exit Function
Fail:
    Application.ScreenUpdating = True
    SAYFALARI_YENILE_SESSIZ = "FAIL:" & Err.Description & IIf(Len(lg) > 0, vbLf & lg, "")
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
    Dim wb As Workbook, groups As Collection, k As Variant, ws As Worksheet, nG As Long, sh() As Worksheet
    Dim toDel As Collection, delList As String, s As Variant, made As Collection, n As Long, t As String
    Dim dsh() As Worksheet, extraD As Collection
    On Error GoTo Fail
    Set wb = ThisWorkbook
    lg = ""
    Application.ScreenUpdating = False
    Set groups = ReadGroups(wb)
    nG = NameRange(wb, N_GTAB).Rows.Count

    ' group sheets by their group number; sheets of groups that no longer exist
    Set toDel = FindGroupSheets(wb, nG, sh)
    For n = 1 To nG
        If Not sh(n) Is Nothing Then
            If Not InColl(groups, n) Then
                toDel.Add sh(n)
                Set sh(n) = Nothing
            End If
        End If
    Next n
    ' data sheets of groups that no longer exist go with them
    Set extraD = FindDataSheets(wb, nG, dsh)
    For n = 1 To nG
        If Not dsh(n) Is Nothing Then
            If Not InColl(groups, n) Then
                toDel.Add dsh(n)
                Set dsh(n) = Nothing
            End If
        End If
    Next n
    For Each s In extraD
        toDel.Add s
    Next s
    For Each s In toDel
        delList = delList & "  " & s.Name & vbLf
    Next s
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
    Set made = New Collection
    For Each k In groups
        If sh(CLng(k)) Is Nothing Then
            Set sh(CLng(k)) = MakeSheet(wb, CLng(k))
            made.Add CLng(k)
        End If
    Next k
    If toDel.Count > 0 Then
        Application.DisplayAlerts = False
        For Each s In toDel
            s.Delete
        Next s
        Application.DisplayAlerts = True
        lg = lg & U("Silinen sayfa:") & vbLf & delList
    End If

    t = RenameSheets(wb, groups, sh, False)
    If made.Count > 0 Then
        lg = lg & U("Yeni sayfa:") & vbLf
        For Each k In made
            lg = lg & "  " & sh(CLng(k)).Name & vbLf
        Next k
    End If
    If Len(t) > 0 Then lg = lg & U("Ad\u0131 de\u011Fi\u015Fen sayfa:") & vbLf & t

    OrderSheets wb, groups, sh
    EnsureDataSheets wb, groups, sh, dsh
    For Each k In groups
        If Not dsh(CLng(k)) Is Nothing Then lg = lg & FitTablePages(sh(CLng(k)), dsh(CLng(k)))
    Next k
    WriteChains wb, groups, sh
    If Application.Calculation <> xlCalculationAutomatic Then Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    For Each k In groups
        Set ws = sh(CLng(k))
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

' sh(1 To nG) = the sheet of each group, found by the group number in the first cell of its sheet-level
' M_AYAK_ZINCIR (the template is skipped); returns the sheets whose number is above nG (removed groups).
' Stops when two sheets hold the same number; a chain cell without a number is a warning, not a group sheet.
Private Function FindGroupSheets(wb As Workbook, ByVal nG As Long, sh() As Worksheet) As Collection
    Dim ws As Worksheet, r As Range, v As Variant, k As Long
    ReDim sh(1 To nG)
    Set FindGroupSheets = New Collection
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, TEMPLATE, vbTextCompare) <> 0 Then
            Set r = Nothing
            On Error Resume Next
            Set r = ws.Names(M_CHAIN).RefersToRange
            On Error GoTo 0
            If Not r Is Nothing Then
                If r.Worksheet.Name = ws.Name Then
                    v = r.Cells(1, 1).Value
                    k = 0
                    If Not IsError(v) Then
                        If IsNumeric(v) And Not IsEmpty(v) Then
                            If CDbl(v) = Int(CDbl(v)) And CDbl(v) >= 1 And CDbl(v) <= 100000 Then k = CLng(v)
                        End If
                    End If
                    If k = 0 Then
                        lg = lg & U("Uyar\u0131: '") & ws.Name & U("' sayfas\u0131n\u0131n zincir h\u00FCcresinde (") & M_CHAIN & U(") grup numaras\u0131 yok; grup sayfas\u0131 say\u0131lmad\u0131.") & vbLf
                    ElseIf k > nG Then
                        FindGroupSheets.Add ws
                    ElseIf Not sh(k) Is Nothing Then
                        Err.Raise vbObjectError + 32, , U("\u0130ki sayfa ayn\u0131 grup numaras\u0131n\u0131 ta\u015F\u0131yor (grup ") & k & "): '" & sh(k).Name & "' / '" & ws.Name & U("'. Birini silin veya zincir h\u00FCcresindeki (") & M_CHAIN & U(") grup numaras\u0131n\u0131 d\u00FCzeltin.")
                    Else
                        Set sh(k) = ws
                    End If
                End If
            End If
        End If
    Next ws
End Function

' <group>_<pier runs> without spaces (1_P1-P3,P16-P17), characters a sheet name cannot hold removed,
' at most MAX_NAME characters (cut with an ellipsis); <group>_AYAK when the group has no name
' isData: MIDAS_<pier runs> (MIDAS_<group> when the group has no name), the name of the group's data sheet
Private Function SheetNameFor(wb As Workbook, ByVal k As Long, Optional ByVal isData As Boolean = False) As String
    Dim v As Variant, s As String, c As Variant
    v = NameRange(wb, N_GTAB).Cells(k, 2).Value
    If Not IsError(v) Then s = CStr(v)
    For Each c In Array(" ", ChrW(160), vbTab, vbCr, vbLf, "\", "/", "?", "*", "[", "]", ":")
        s = Replace(s, c, "")
    Next c
    If isData Then
        If Len(s) = 0 Then s = CStr(k)
        s = DATA_PFX & s
    Else
        If Len(s) = 0 Then s = Mid$(SHEET_SFX, 2)
        s = CStr(k) & "_" & s
    End If
    If Len(s) > MAX_NAME Then s = Left$(s, MAX_NAME - 1) & ChrW(8230)
    Do While Right$(s, 1) = "'"
        s = Left$(s, Len(s) - 1)
    Loop
    SheetNameFor = s
End Function

' gives every group sheet its name; two passes through free temporary names, so two groups that swap
' names never clash; returns one "old -> new" line per renamed sheet
Private Function RenameSheets(wb As Workbook, groups As Collection, sh() As Worksheet, ByVal isData As Boolean) As String
    Dim k As Variant, j As Variant, want() As String, old() As String, o As Object, own As Boolean
    Dim e As Long, d As String, u2 As String
    ReDim want(LBound(sh) To UBound(sh)): ReDim old(LBound(sh) To UBound(sh))
    For Each k In groups
        want(k) = SheetNameFor(wb, CLng(k), isData)
        Set o = Nothing
        On Error Resume Next
        Set o = wb.Sheets(want(k))
        On Error GoTo 0
        If Not o Is Nothing Then
            own = False
            For Each j In groups
                If StrComp(o.Name, sh(CLng(j)).Name, vbTextCompare) = 0 Then own = True: Exit For
            Next j
            If Not own Then Err.Raise vbObjectError + 33, , "'" & o.Name & "' " & IIf(isData, U("ad\u0131nda veri sayfas\u0131 olmayan bir sayfa var; grup ") & k & U(" veri sayfas\u0131"), U("ad\u0131nda grup sayfas\u0131 olmayan bir sayfa var; grup ") & k & U(" sayfas\u0131")) & U(" bu ad\u0131 alam\u0131yor. O sayfan\u0131n ad\u0131n\u0131 de\u011Fi\u015Ftirin.")
        End If
        u2 = UCase$(want(k))
        If isData Then u2 = ""
        If Right$(u2, 2) = "_Y" Or Right$(u2, 6) = "_YATAY" Or Right$(u2, 6) = "_KILIF" Then
            lg = lg & U("Uyar\u0131: '") & want(k) & U("' ad\u0131 _Y, _YATAY veya _KILIF ile bitiyor; i\u00E7indekiler ve listeler bu sayfay\u0131 almaz.") & vbLf
        End If
    Next k
    For Each k In groups
        If StrComp(sh(CLng(k)).Name, want(k), vbBinaryCompare) <> 0 Then
            old(k) = sh(CLng(k)).Name
            sh(CLng(k)).Name = FreeName(wb, TMP_BASE)
        End If
    Next k
    For Each k In groups
        If Len(old(k)) > 0 Then
            On Error Resume Next
            sh(CLng(k)).Name = want(k)
            e = Err.Number: d = Err.Description
            On Error GoTo 0
            If e <> 0 Then Err.Raise vbObjectError + 34, , U("Sayfa ad\u0131 verilemedi (") & want(k) & "): " & d
            If Left$(old(k), Len(TMP_BASE)) <> TMP_BASE Then RenameSheets = RenameSheets & "  " & old(k) & " -> " & want(k) & vbLf
        End If
    Next k
End Function

' a sheet name that is not in use: base & 1, base & 2 ...
Private Function FreeName(wb As Workbook, ByVal base As String) As String
    Dim i As Long
    Do
        i = i + 1
        If Not SheetExists(wb, base & i) Then FreeName = base & i: Exit Function
    Loop
End Function

' copy of a template sheet placed after the last visible sheet (the callers order the sheets afterwards); the new
' sheet is the one whose name was not there before. Measured 2026-11-26 (TEMEL_GRUP): a copy placed after a very
' hidden last sheet lands right after its template, so the sheet at Worksheets.Count is another one (here the
' baseline SABLON_MIDAS_IZ would have been renamed). e / d = error number and text, Nothing when e <> 0.
Private Function CopyTemplate(wb As Workbook, tpl As Worksheet, ByRef e As Long, ByRef d As String) As Worksheet
    Dim seen As Object, w As Worksheet, anchor As Worksheet, ws As Worksheet, nNew As Long
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For Each w In wb.Worksheets
        seen(w.Name) = True
        If w.Visible = xlSheetVisible Then Set anchor = w
    Next w
    On Error Resume Next
    If anchor Is Nothing Then tpl.Copy After:=tpl Else tpl.Copy After:=anchor
    e = Err.Number: d = Err.Description
    Err.Clear
    On Error GoTo 0
    If e <> 0 Then Exit Function
    For Each w In wb.Worksheets
        If Not seen.Exists(w.Name) Then Set ws = w: nNew = nNew + 1
    Next w
    If nNew = 1 Then
        Set CopyTemplate = ws
    Else
        e = 1: d = "copy (" & nNew & U(" yeni sayfa)")
    End If
End Function

' copy of the template with a temporary name (RenameSheets gives the real one), group number in its chain cell
Private Function MakeSheet(wb As Workbook, ByVal k As Long) As Worksheet
    Dim tpl As Worksheet, ws As Worksheet, vis As Long, e As Long, d As String, nm As String
    If Not SheetExists(wb, TEMPLATE) Then Err.Raise vbObjectError + 25, , U("\u015Eablon sayfa yok: ") & TEMPLATE
    Set tpl = wb.Worksheets(TEMPLATE)
    vis = tpl.Visible
    nm = FreeName(wb, TMP_BASE)
    On Error Resume Next
    tpl.Visible = xlSheetVisible
    Application.DisplayAlerts = False
    Set ws = CopyTemplate(wb, tpl, e, d)
    If e = 0 Then
        ws.Name = nm
        ws.Names(M_CHAIN).RefersToRange.Cells(1, 1).Value = k
        e = Err.Number: d = Err.Description
    End If
    Err.Clear
    tpl.Visible = vis
    Application.DisplayAlerts = True
    On Error GoTo 0
    If e <> 0 Then Err.Raise vbObjectError + 26, , U("Grup sayfas\u0131 a\u00E7\u0131lamad\u0131 (grup ") & k & "): " & d
    Set MakeSheet = ws
End Function

' ---- Tablo 63 pages ----
' copies of the pattern page (M_T63_SAYFA) below it so that the table holds the rows of the data sheet; returns
' a log line when the number of copies changed. Without M_T63_SAYFA, or before MIDAS_AYAK made the data block,
' nothing is done.
' forceN >= 0: exactly that many copies, whatever the data (SABLON_UYGULA takes them out with 0, then fits again).
Private Function FitTablePages(ws As Worksheet, ds As Worksheet, Optional ByVal forceN As Long = -1) As String
    Dim pat As Range, p0 As Long, pr As Long, ex0 As Long, nOld As Long, nNew As Long, nRows As Long
    Dim cap0 As Long, dpp As Long, r As Long, i As Long, ofs As Long, wOff As Long, f As String, k As Long
    Dim top As Long, prevW As Long, c As Range, co As Range, v As Variant, cPat As Variant, delta As Long
    Dim cm As Long, eNo As Long, eTx As String, nextOff As Long
    Set pat = Nothing
    On Error Resume Next
    Set pat = ws.Names(M_T63_PAGE).RefersToRange
    On Error GoTo 0
    If pat Is Nothing Then Exit Function
    nRows = DataRows(ds)
    If nRows < 0 And forceN < 0 Then Exit Function
    p0 = pat.Row: pr = pat.Rows.Count: ex0 = p0 + pr
    nOld = NameNumber(ws, M_T63_EXTRA)
    If forceN >= 0 And nOld = forceN Then Exit Function
    ' rows of the table up to the end of the pattern page, and of the pattern page alone
    For r = 1 To ex0 - 1
        If IsTableRow(ws.Cells(r, 2), ds) Then
            cap0 = cap0 + 1
            If r >= p0 Then
                dpp = dpp + 1
                If ofs = 0 Then
                    f = ws.Cells(r, 2).FormulaR1C1
                    k = InStr(f, "!R[")
                    If k > 0 Then ofs = CLng(Val(Mid$(f, k + 3)))
                End If
            End If
        End If
    Next r
    If dpp = 0 Or ofs = 0 Then Err.Raise vbObjectError + 37, , ws.Name & U(": Tablo 63 \u00F6rnek sayfas\u0131nda (") & M_T63_PAGE & U(") tablo sat\u0131r\u0131 yok.")
    If nRows > cap0 Then nNew = -Int(-(nRows - cap0) / dpp)
    If forceN >= 0 Then nNew = forceN
    If nNew = nOld Then Exit Function
    wOff = -1
    For r = p0 To ex0 - 1
        If ws.Cells(r, 23).HasFormula Then
            If ws.Cells(r, 23).Formula Like "=$W$*+$CG*" Then wOff = r - p0
        End If
    Next r
    If wOff < 0 Then Err.Raise vbObjectError + 37, , ws.Name & U(": Tablo 63 \u00F6rnek sayfas\u0131nda sayfa numaras\u0131 (W) yok.")
    Set co = SheetRange(ws, M_PAGE)
    cPat = ws.Cells(p0, co.Column).Value
    ' the page after the table: its page-number cell, found before rows move (after a delete its formula
    ' would read #REF! and no longer look like a page number)
    nextOff = -1
    For r = ex0 + pr * nOld To ex0 + pr * nOld + 2 * pr
        If ws.Cells(r, 23).HasFormula Then
            If ws.Cells(r, 23).Formula Like "=$W$*+$CG*" Then nextOff = r - (ex0 + pr * nOld): Exit For
        End If
    Next r
    ' no recalculation while rows move and formulas are rewritten
    cm = Application.Calculation
    If cm <> xlCalculationManual Then Application.Calculation = xlCalculationManual
    On Error GoTo Bad
    ' make or delete the copies
    If nNew > nOld Then
        For i = nOld + 1 To nNew
            pat.EntireRow.Copy
            ws.Rows(ex0 + pr * (i - 1)).Insert Shift:=xlDown
        Next i
        Application.CutCopyMode = False
    Else
        ws.Rows((ex0 + pr * nNew) & ":" & (ex0 + pr * nOld - 1)).Delete
    End If
    ' each copy: rows continue the table, page number follows the page before it
    For i = 1 To nNew
        top = ex0 + pr * (i - 1)
        For Each c In ws.Range(ws.Cells(top, 2), ws.Cells(top + pr - 1, 22))
            If c.HasFormula Then
                f = c.FormulaR1C1
                If InStr(f, "R[" & ofs & "]") > 0 Then c.FormulaR1C1 = Replace(f, "R[" & ofs & "]", "R[" & (ofs + i * (dpp - pr)) & "]")
            End If
        Next c
        prevW = IIf(i = 1, p0 + wOff, top - pr + wOff)
        ws.Cells(top + wOff, 23).Formula = "=$W$" & prevW & "+$CG" & (top + wOff)
    Next i
    ' the page after the table numbers on from the last copy (or from the pattern page)
    prevW = IIf(nNew = 0, p0 + wOff, ex0 + pr * (nNew - 1) + wOff)
    If nextOff >= 0 Then
        r = ex0 + pr * nNew + nextOff
        ws.Cells(r, 23).Formula = "=$W$" & prevW & "+$CG" & r
    End If
    ' page index: the copies count on from the pattern page, the pages after them move by the difference
    Set co = SheetRange(ws, M_PAGE)
    delta = nNew - nOld
    v = co.Value
    For r = ex0 To co.Row + co.Rows.Count - 1
        i = r - co.Row + 1
        If Not IsEmpty(v(i, 1)) And IsNumeric(v(i, 1)) Then
            If r < ex0 + pr * nNew Then
                v(i, 1) = CDbl(cPat) + (r - ex0) \ pr + 1
            Else
                v(i, 1) = CDbl(v(i, 1)) + delta
            End If
        End If
    Next r
    co.Value = v
    SetNameNumber ws, M_T63_EXTRA, nNew
    If Application.Calculation <> cm Then Application.Calculation = cm
    FitTablePages = ws.Name & U(": Tablo 63 ") & nRows & U(" sat\u0131r, ek sayfa ") & nOld & " -> " & nNew & vbLf
    Exit Function
Bad:
    eNo = Err.Number: eTx = Err.Description
    Application.CutCopyMode = False
    If Application.Calculation <> cm Then Application.Calculation = cm
    Err.Raise eNo, , ws.Name & " (Tablo 63): " & eTx
End Function

' a table row of Tablo 63: column B reads column B of the group's data sheet
Private Function IsTableRow(c As Range, ds As Worksheet) As Boolean
    Dim f As String
    If Not c.HasFormula Then Exit Function
    f = c.Formula
    IsTableRow = (InStr(1, f, ds.Name & "!B", vbTextCompare) > 0 Or InStr(1, f, ds.Name & "'!B", vbTextCompare) > 0)
End Function

' rows of the force block of a data sheet (names in its 3rd column); -1 when MIDAS_AYAK has not made it yet
Private Function DataRows(ds As Worksheet) As Long
    Dim b As Range
    On Error Resume Next
    Set b = ds.Names(M_BLOCK).RefersToRange
    On Error GoTo 0
    If b Is Nothing Then DataRows = -1: Exit Function
    DataRows = Application.WorksheetFunction.CountA(b.Columns(3))
End Function

' number held by a sheet-level name (=n); 0 when the name is missing
Private Function NameNumber(ws As Worksheet, ByVal nm As String) As Long
    Dim n As Name
    On Error Resume Next
    Set n = ws.Names(nm)
    On Error GoTo 0
    If Not n Is Nothing Then NameNumber = CLng(Val(Mid$(n.RefersTo, 2)))
End Function

Private Sub SetNameNumber(ws As Worksheet, ByVal nm As String, ByVal v As Long)
    Dim n As Name
    On Error Resume Next
    Set n = ws.Names(nm)
    On Error GoTo 0
    If n Is Nothing Then
        ws.Names.Add Name:=nm, RefersTo:="=" & v
    ElseIf n.RefersTo <> "=" & v Then
        n.RefersTo = "=" & v
    End If
End Sub

' group number of a data sheet: its sheet-level name M_TESIR_NO (=k); an older data sheet named MIDAS_<k>
' without that name -> k; anything else -> 0
Private Function DataGroupOf(ws As Worksheet) As Long
    Dim nm As Name, v As String, p As String
    On Error Resume Next
    Set nm = ws.Names(M_DATA_NO)
    On Error GoTo 0
    If Not nm Is Nothing Then
        v = Mid$(nm.RefersTo, 2)
        If Len(v) > 0 And Len(v) <= 3 Then
            If v Like String$(Len(v), "#") Then DataGroupOf = CLng(v)
        End If
        Exit Function
    End If
    If StrComp(Left$(ws.Name, Len(DATA_PFX)), DATA_PFX, vbTextCompare) <> 0 Then Exit Function
    p = Mid$(ws.Name, Len(DATA_PFX) + 1)
    If Len(p) = 0 Or Len(p) > 3 Then Exit Function
    If p Like String$(Len(p), "#") Then DataGroupOf = CLng(p)
End Function

' dsh(1 To nG) = the data sheet of each group; returns the data sheets whose number is above nG;
' stops when two sheets hold the same group number
Private Function FindDataSheets(wb As Workbook, ByVal nG As Long, dsh() As Worksheet) As Collection
    Dim ws As Worksheet, k As Long
    ReDim dsh(1 To nG)
    Set FindDataSheets = New Collection
    For Each ws In wb.Worksheets
        k = DataGroupOf(ws)
        If k > nG Then
            FindDataSheets.Add ws
        ElseIf k > 0 Then
            If Not dsh(k) Is Nothing Then Err.Raise vbObjectError + 36, , "Grup " & k & U(" i\u00E7in iki veri sayfas\u0131 var: '") & dsh(k).Name & "' / '" & ws.Name & U("'. Birini silin.")
            Set dsh(k) = ws
        End If
    Next ws
End Function

' sheet-level name M_TESIR_NO = k on the data sheet (a number, no cell)
Private Sub SetDataNo(ws As Worksheet, ByVal k As Long)
    Dim nm As Name
    On Error Resume Next
    Set nm = ws.Names(M_DATA_NO)
    On Error GoTo 0
    If nm Is Nothing Then
        ws.Names.Add Name:=M_DATA_NO, RefersTo:="=" & k
    ElseIf nm.RefersTo <> "=" & k Then
        nm.RefersTo = "=" & k
    End If
End Sub

' the data sheet of each group (a copy of SABLON_MIDAS), its group number in M_TESIR_NO, named MIDAS_<pier runs>,
' in group order right before INPUT (after the last group sheet when there is no INPUT); the group sheet's formulas
' that read SABLON_MIDAS are pointed at its own data sheet; a data sheet read for other piers than the group's
' piers now gives a warning
Private Sub EnsureDataSheets(wb As Workbook, groups As Collection, sh() As Worksheet, dsh() As Worksheet)
    Dim k As Variant, ds As Worksheet, anchor As Worksheet, nm As String, made As Collection, repo As String
    Dim f As Range, rd As Range, wasP As String, nowP As String, t As String, i As Long
    If Not SheetExists(wb, DATA_TPL) Then
        lg = lg & U("Not: ") & DATA_TPL & U(" sayfas\u0131 yok; grup veri sayfalar\u0131 (") & DATA_PFX & U("<ayaklar>) a\u00E7\u0131lmad\u0131.") & vbLf
        Exit Sub
    End If
    Set made = New Collection
    For Each k In groups
        If dsh(CLng(k)) Is Nothing Then
            Set dsh(CLng(k)) = MakeDataSheet(wb, FreeName(wb, TMP_BASE))
            made.Add CLng(k)
        End If
        SetDataNo dsh(CLng(k)), CLng(k)
    Next k
    t = RenameSheets(wb, groups, dsh, True)
    If made.Count > 0 Then
        lg = lg & U("Yeni veri sayfas\u0131:") & vbLf
        For Each k In made
            lg = lg & "  " & dsh(CLng(k)).Name & vbLf
        Next k
    End If
    If Len(t) > 0 Then lg = lg & U("Ad\u0131 de\u011Fi\u015Fen veri sayfas\u0131:") & vbLf & t
    ' order: before INPUT, in group order (built from the last one backwards)
    Set anchor = Nothing
    On Error Resume Next
    Set anchor = wb.Worksheets(DATA_BEFORE)
    On Error GoTo 0
    If anchor Is Nothing Then
        Set anchor = sh(CLng(groups(groups.Count)))
        For Each k In groups
            Set ds = dsh(CLng(k))
            If ds.Index <> anchor.Index + 1 Then ds.Move After:=anchor
            Set anchor = ds
        Next k
    Else
        For i = groups.Count To 1 Step -1
            Set ds = dsh(CLng(groups(i)))
            If ds.Index <> anchor.Index - 1 Then ds.Move Before:=anchor
            Set anchor = ds
        Next i
    End If
    For Each k In groups
        Set ds = dsh(CLng(k))
        nm = ds.Name
        Set f = Nothing
        On Error Resume Next
        Set f = sh(CLng(k)).UsedRange.Find(What:=DATA_TPL & "!", LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
        On Error GoTo 0
        If Not f Is Nothing Then
            sh(CLng(k)).UsedRange.Replace What:=DATA_TPL & "!", Replacement:="'" & Replace(nm, "'", "''") & "'!", LookAt:=xlPart, MatchCase:=False
            repo = repo & "  " & sh(CLng(k)).Name & " -> " & nm & vbLf
        End If
        Set rd = Nothing
        On Error Resume Next
        Set rd = ds.Names(M_READ).RefersToRange
        On Error GoTo 0
        If Not rd Is Nothing Then
            wasP = Trim$(CellText(rd.Cells(1, 1).Value))
            nowP = Trim$(CellText(NameRange(wb, N_GTAB).Cells(CLng(k), 2).Value))
            If Len(wasP) > 0 And StrComp(wasP, nowP, vbTextCompare) <> 0 Then
                lg = lg & U("Uyar\u0131: ") & nm & U(" tesirleri ") & wasP & U(" i\u00E7in okunmu\u015F, grup \u015Fimdi ") & nowP & U(": MIDAS_AYAK ile yeniden okuyun.") & vbLf
            End If
        End If
    Next k
    If Len(repo) > 0 Then lg = lg & U("Kendi veri sayfas\u0131na ba\u011Flanan grup sayfas\u0131:") & vbLf & repo
End Sub

' copy of SABLON_MIDAS after the last visible sheet (EnsureDataSheets orders it), named nm, visible
Private Function MakeDataSheet(wb As Workbook, ByVal nm As String) As Worksheet
    Dim tpl As Worksheet, ws As Worksheet, vis As Long, e As Long, d As String
    Set tpl = wb.Worksheets(DATA_TPL)
    vis = tpl.Visible
    On Error Resume Next
    tpl.Visible = xlSheetVisible
    Application.DisplayAlerts = False
    Set ws = CopyTemplate(wb, tpl, e, d)
    If e = 0 Then
        ws.Name = nm
        ws.Visible = xlSheetVisible
        e = Err.Number: d = Err.Description
    End If
    Err.Clear
    tpl.Visible = vis
    Application.DisplayAlerts = True
    On Error GoTo 0
    If e <> 0 Then Err.Raise vbObjectError + 35, , U("Veri sayfas\u0131 a\u00E7\u0131lamad\u0131 (") & nm & "): " & d
    Set MakeDataSheet = ws
End Function

Private Sub OrderSheets(wb As Workbook, groups As Collection, sh() As Worksheet)
    Dim k As Variant, ws As Worksheet, before As Worksheet
    On Error Resume Next
    Set before = wb.Worksheets(AFTER_SHEET)
    On Error GoTo 0
    For Each k In groups
        Set ws = sh(CLng(k))
        If Not before Is Nothing Then
            If ws.Index <> before.Index + 1 Then ws.Move After:=before
        End If
        Set before = ws
    Next k
End Sub

Private Sub WriteChains(wb As Workbook, groups As Collection, sh() As Worksheet)
    Dim k As Variant, ws As Worksheet, c As Range, lastC As Range, j As Long, f As String, r As Range
    For Each k In groups
        Set ws = sh(CLng(k))
        Set c = SheetRange(ws, M_CHAIN)
        If CStr(c.Cells(1, 1).Value) <> CStr(CLng(k)) Then c.Cells(1, 1).Value = CLng(k)
        For j = 2 To 4
            If lastC Is Nothing Then
                f = "=INDEX(" & N_PREV & "," & (j - 1) & ")"
            Else
                f = "='" & Replace(lastC.Worksheet.Name, "'", "''") & "'!" & lastC.Cells(2, j).Address
            End If
            If c.Cells(1, j).Formula <> f Then c.Cells(1, j).Formula = f
        Next j
        Set lastC = c
    Next k
    Set r = NameRange(wb, N_LAST)
    For j = 1 To 3
        f = "='" & Replace(lastC.Worksheet.Name, "'", "''") & "'!" & lastC.Cells(2, j + 1).Address
        If r.Cells(j).Formula <> f Then r.Cells(j).Formula = f
    Next j
End Sub

' hides the rows whose gate is 0, print area = the pages whose first row is open; returns the page count
Private Function ApplyGate(ws As Worksheet) As Long
    Dim rg As Range, rp As Range, rb As Range, g As Variant, p As Variant
    Dim n As Long, r0 As Long, i As Long, j As Long, o As Boolean, pg As Long, cur As Long, a As Long, addr As String
    Dim part As Range, pa As Range
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
                    Set part = Intersect(rb.EntireColumn, ws.Rows(r0 + a - 1).Resize(i - a))
                    addr = addr & "," & part.Address
                    If pa Is Nothing Then Set pa = part Else Set pa = Union(pa, part)
                    ApplyGate = ApplyGate + 1
                End If
            End If
            pg = cur: a = i
        End If
    Next i
    If ApplyGate = 0 Then Err.Raise vbObjectError + 29, , ws.Name & U(": a\u00E7\u0131k sayfa yok.")
    SetPrintAreaLong ws, Mid$(addr, 2), pa
    GateShapes ws, g, r0
End Function

' print area of many pieces (one per page): PageSetup and a text formula take at most 255 characters, so a
' longer one is given to Excel's own print-area name of the sheet as the union Range itself (its local name
' differs from its name in a Turkish Excel, as RAPOR_LISTE finds it); measured 2026-11-16: 19 pages, PDF 19 pages
Private Sub SetPrintAreaLong(ws As Worksheet, ByVal addr As String, pa As Range)
    Dim nm As Name, own As Name
    If Len(addr) <= 255 Then ws.PageSetup.PrintArea = addr: Exit Sub
    ws.PageSetup.PrintArea = pa.Areas(1).Address
    For Each nm In ws.Parent.Names
        If Right$(nm.Name, 11) = "!Print_Area" Then
            If NameOnSheet(nm, ws) Then
                If own Is Nothing Or nm.NameLocal <> nm.Name Then Set own = nm
            End If
        End If
    Next nm
    If own Is Nothing Then Err.Raise vbObjectError + 38, , ws.Name & U(": yazd\u0131rma alan\u0131 ad\u0131 bulunamad\u0131.")
    own.RefersTo = pa
End Sub

Private Function NameOnSheet(nm As Name, ws As Worksheet) As Boolean
    On Error Resume Next
    NameOnSheet = (TypeOf nm.Parent Is Worksheet)
    If NameOnSheet Then NameOnSheet = (nm.Parent.Name = ws.Name)
    On Error GoTo 0
End Function

' drawing objects follow the gate of their top-left row: hidden on a closed row, shown on an open one;
' comments, form controls and ActiveX controls are left alone, objects outside the gated rows too
Private Sub GateShapes(ws As Worksheet, g As Variant, ByVal r0 As Long)
    Dim sh As Shape, r As Long, n As Long, want As Boolean
    n = UBound(g, 1)
    For Each sh In ws.Shapes
        Select Case sh.Type
            Case msoComment, msoFormControl, msoOLEControlObject
            Case Else
                r = 0
                On Error Resume Next
                r = sh.TopLeftCell.Row
                On Error GoTo 0
                If r >= r0 And r < r0 + n Then
                    want = (CDbl(g(r - r0 + 1, 1)) <> 0)
                    If (sh.Visible = msoTrue) <> want Then sh.Visible = IIf(want, msoTrue, msoFalse)
                End If
        End Select
    Next sh
End Sub

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

' any sheet (worksheet or chart) of that name, case ignored as Excel does
Private Function SheetExists(wb As Workbook, ByVal nm As String) As Boolean
    Dim o As Object
    On Error Resume Next
    Set o = wb.Sheets(nm)
    On Error GoTo 0
    SheetExists = Not o Is Nothing
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
    If scripted Then Exit Sub
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

' ---- template to group sheets (owner 2026-11-16) ----
' SABLON_GOSTER: shows the very hidden template for editing, hides it again on the next press.
Public Sub SABLON_GOSTER()
    Dim wb As Workbook, tpl As Worksheet, ws As Worksheet, rg As Range, g As Variant, i As Long
    Set wb = ThisWorkbook
    If Not SheetExists(wb, TEMPLATE) Then ShowMsg U("\u015Eablon sayfa yok: ") & TEMPLATE, vbExclamation: Exit Sub
    Set tpl = wb.Worksheets(TEMPLATE)
    If tpl.Visible = xlSheetVisible Then
        If ActiveSheet Is tpl Then
            For Each ws In wb.Worksheets
                If ws.Visible = xlSheetVisible And Not ws Is tpl Then ws.Activate: Exit For
            Next ws
        End If
        ' back to the pages of its own gate (rows, pictures, print area), as the group sheets are
        On Error Resume Next
        Application.ScreenUpdating = False
        ApplyGate tpl
        Application.ScreenUpdating = True
        On Error GoTo 0
        tpl.Visible = xlSheetVeryHidden
        ShowMsg U("\u015Eablon gizlendi. De\u011Fi\u015Fiklikleri grup sayfalar\u0131na ta\u015F\u0131mak i\u00E7in: '\u015Eablonu grup sayfalar\u0131na uygula'."), vbInformation
    Else
        tpl.Visible = xlSheetVisible
        ' every section type editable: all gated rows and their pictures open
        On Error Resume Next
        Application.ScreenUpdating = False
        Set rg = SheetRange(tpl, M_GATE)
        If Not rg Is Nothing Then
            rg.EntireRow.Hidden = False
            g = rg.Value
            For i = 1 To UBound(g, 1): g(i, 1) = 1: Next i
            GateShapes tpl, g, rg.Row
        End If
        Application.ScreenUpdating = True
        On Error GoTo 0
        tpl.Activate
        ShowMsg U("\u015Eablon a\u00E7\u0131k: ") & TEMPLATE & vbLf & _
            U("B\u00FCt\u00FCn kesit tiplerinin sat\u0131rlar\u0131 ve resimleri a\u00E7\u0131k (gizlenince eski h\u00E2line d\u00F6ner). ") & _
            U("Burada yaln\u0131z t\u00FCm gruplarda ge\u00E7erli olacak d\u00FCzeltmeleri yap\u0131n; gruba \u00F6zel tasar\u0131m de\u011Ferlerini grup sayfas\u0131na yaz\u0131n. Sonra '\u015Eablonu grup sayfalar\u0131na uygula'."), vbInformation
    End If
End Sub

' tests: the same without the message
Public Sub SABLON_GOSTER_SESSIZ()
    Dim was As Boolean
    was = scripted
    scripted = True
    SABLON_GOSTER
    scripted = was
End Sub

Public Sub SABLON_UYGULA()
    Dim s As String
    scripted = False
    s = RunTemplateApply(True, False)
    If Left$(s, 5) = "FAIL:" Then ShowMsg Mid$(s, 6), vbExclamation Else ShowMsg s, vbInformation
End Sub

' tests: doApply False = list only (nothing written; the Tablo 63 copies are made again), force True = conflicts too
Public Function SABLON_UYGULA_SESSIZ(Optional ByVal doApply As Boolean = True, Optional ByVal force As Boolean = False) As String
    scripted = True
    SABLON_UYGULA_SESSIZ = RunTemplateApply(doApply, force)
End Function

Private Function RunTemplateApply(ByVal doApply As Boolean, ByVal force As Boolean) As String
    Dim wb As Workbook, tpl As Worksheet, bs As Worksheet, nG As Long, sh() As Worksheet, dsh() As Worksheet
    Dim nR As Long, nC As Long, t As Variant, b As Variant, g As Variant, k As Long, r As Long, c As Long
    Dim tv As String, bv As String, gv As String, skipM() As Boolean, ws As Worksheet, ds As Worksheet
    Dim upd() As Collection, con() As Collection, nUpd As Long, nCon As Long, it As Variant, rep As String
    Dim conList As String, writeCon As Boolean, shrunk As Boolean, cm As Long, ha As Variant, keep As String
    Dim tail As String, eTx As String, wasScripted As Boolean, sft() As Long, pEnd() As Long, skipG() As Boolean
    Dim tplHid As Boolean, chg As Collection, picNote As String, picMiss As String, nMoved As Long, q As Variant, picLst As String
    Dim s1 As Long, e1 As Long
    On Error GoTo Fail
    wasScripted = scripted
    lg = ""
    Set wb = ThisWorkbook
    If Not SheetExists(wb, TEMPLATE) Then Err.Raise vbObjectError + 40, , U("\u015Eablon sayfa yok: ") & TEMPLATE
    Set tpl = wb.Worksheets(TEMPLATE)
    cm = Application.Calculation
    Application.ScreenUpdating = False
    nR = tpl.UsedRange.Row + tpl.UsedRange.Rows.Count - 1
    nC = tpl.UsedRange.Column + tpl.UsedRange.Columns.Count - 1
    If Not SheetExists(wb, BASE_SHEET) Then
        WriteBaseline wb, tpl, nR, nC
        WriteShapeBase wb, tpl
        Application.ScreenUpdating = True
        RunTemplateApply = U("\u0130z sayfas\u0131 (") & BASE_SHEET & U(") yoktu: \u015Fablonun bug\u00FCnk\u00FC h\u00E2li kaydedildi, hi\u00E7bir grup sayfas\u0131 de\u011Fi\u015Fmedi. Bundan sonra \u015Fablonda yapt\u0131\u011F\u0131n\u0131z de\u011Fi\u015Fiklikler bu tu\u015Fla grup sayfalar\u0131na ta\u015F\u0131n\u0131r.")
        Exit Function
    End If
    Set bs = wb.Worksheets(BASE_SHEET)
    With bs.UsedRange
        If .Row + .Rows.Count - 1 > nR Then nR = .Row + .Rows.Count - 1
        If .Column + .Columns.Count - 1 > nC Then nC = .Column + .Columns.Count - 1
    End With
    t = tpl.Range(tpl.Cells(1, 1), tpl.Cells(nR, nC)).Formula
    b = bs.Range(bs.Cells(1, 1), bs.Cells(nR, nC)).Value
    ' never compared: the chain cells (WriteChains), the page index (FitTablePages), array formulas
    ReDim skipM(1 To nR, 1 To nC)
    MarkSkip tpl, M_CHAIN, skipM, nR, nC
    MarkSkip tpl, M_PAGE, skipM, nR, nC
    ha = tpl.Range(tpl.Cells(1, 1), tpl.Cells(nR, nC)).HasArray
    If IsNull(ha) Then
        For r = 1 To nR
            For c = 1 To nC
                If Len(CellText(t(r, c))) > 0 Then
                    If tpl.Cells(r, c).HasArray Then skipM(r, c) = True
                End If
            Next c
        Next r
    ElseIf ha = True Then
        Err.Raise vbObjectError + 41, , U("\u015Eablonun tamam\u0131 dizi form\u00FCl\u00FC; uygulanamaz.")
    End If

    ' template = baseline everywhere: nothing to carry, the group sheets are not touched (the page rebuild takes minutes)
    nUpd = 0
    For r = 1 To nR
        For c = 1 To nC
            If Not skipM(r, c) Then
                If CellText(t(r, c)) <> CellText(b(r, c)) Then nUpd = nUpd + 1
            End If
        Next c
    Next r
    ' pictures and drawings: place and size on the template against the list of the last apply (rows open)
    tplHid = (tpl.Visible <> xlSheetVisible)
    Set chg = New Collection
    If SheetExists(wb, BASE_SHAPES) Then
        OpenRows tpl
        Set chg = ShapeChanges(wb, tpl, picNote)
        For Each q In chg
            picLst = picLst & "  " & CStr(q(1)) & " (" & tpl.Cells(CLng(q(2)), CLng(q(3))).Address(False, False) & ")" & vbLf
        Next q
        If chg.Count > 0 Then lg = lg & U("\u015Eablonda yeri / boyutu de\u011Fi\u015Fen resimler (grup sayfas\u0131ndaki resim ayn\u0131 yere ve boyuta getirilir, g\u00F6r\u00FCnt\u00FCs\u00FC de\u011Fi\u015Fmez):") & vbLf & picLst
        If Len(picNote) > 0 Then lg = lg & U("\u015Eablonda yeni (+) veya silinmi\u015F (-) resimler ta\u015F\u0131nmaz; grup sayfalar\u0131nda elle yap\u0131n:") & vbLf & picNote
    Else
        WriteShapeBase wb, tpl
        lg = lg & U("Resim izi (") & BASE_SHAPES & U(") yoktu: \u015Fablondaki resimlerin bug\u00FCnk\u00FC yeri kaydedildi; bundan sonra ta\u015F\u0131n\u0131r.") & vbLf
    End If
    If nUpd = 0 And chg.Count > 0 Then
        ' pictures only: no cell is compared, the Tablo 63 pages stay (rows below them are counted lower)
        If doApply And Not wasScripted Then
            Application.ScreenUpdating = True
            If Ask(lg & vbLf & chg.Count & U(" resim her grup sayfas\u0131nda \u015Fablondaki yerine / boyutuna getirilsin mi?")) <> vbYes Then doApply = False
            Application.ScreenUpdating = False
        End If
        If doApply Then
            nG = NameRange(wb, N_GTAB).Rows.Count
            FindGroupSheets wb, nG, sh
            For k = 1 To nG
                If Not sh(k) Is Nothing Then
                    T63Shift sh(k), s1, e1
                    OpenRows sh(k)
                    nMoved = MoveShapes(sh(k), chg, s1, e1, picMiss)
                    ApplyGate sh(k)
                    lg = lg & "  " & sh(k).Name & ": " & nMoved & U(" resim") & vbLf
                End If
            Next k
            If Len(picMiss) > 0 Then lg = lg & U("Bulunamayan (o sayfada silinmi\u015F veya de\u011Fi\u015Ftirilmi\u015F) resimler:") & vbLf & picMiss
            WriteShapeBase wb, tpl, True
        Else
            lg = lg & U("Hi\u00E7bir resim ta\u015F\u0131nmad\u0131.") & vbLf
        End If
        If tplHid Then CloseRows tpl
        Application.ScreenUpdating = True
        RunTemplateApply = lg
        Exit Function
    End If
    If nUpd = 0 Then
        If tplHid Then CloseRows tpl
        Application.ScreenUpdating = True
        RunTemplateApply = U("\u015Eablonda son uygulamadan beri de\u011Fi\u015Fen h\u00FCcre veya yeri de\u011Fi\u015Fen resim yok; grup sayfalar\u0131na dokunulmad\u0131.")
        If Len(lg) > 0 Then RunTemplateApply = lg & RunTemplateApply
        Exit Function
    End If
    nUpd = 0

    nG = NameRange(wb, N_GTAB).Rows.Count
    FindGroupSheets wb, nG, sh
    FindDataSheets wb, nG, dsh
    ReDim upd(1 To nG): ReDim con(1 To nG): ReDim sft(1 To nG): ReDim pEnd(1 To nG): ReDim skipG(1 To nG)
    ' rows the Tablo 63 copies add below the pattern page now (conflicts are listed at the group sheet's own rows)
    For k = 1 To nG
        If Not sh(k) Is Nothing Then
            T63Shift sh(k), sft(k), pEnd(k)
            If sft(k) > 0 And dsh(k) Is Nothing Then
                skipG(k) = True
                lg = lg & U("Uyar\u0131: ") & sh(k).Name & U(" sayfas\u0131nda Tablo 63 ek sayfalar\u0131 var ama veri sayfas\u0131 yok; bu sayfa kar\u015F\u0131la\u015Ft\u0131r\u0131lmad\u0131.") & vbLf
            End If
        End If
    Next k
    ' Tablo 63 copies out first: Excel moves every reference back to the template rows
    shrunk = True
    For k = 1 To nG
        If Not sh(k) Is Nothing And Not dsh(k) Is Nothing Then FitTablePages sh(k), dsh(k), 0
    Next k
    If Application.Calculation <> xlCalculationManual Then Application.Calculation = xlCalculationManual
    For k = 1 To nG
        Set upd(k) = New Collection: Set con(k) = New Collection
        If Not sh(k) Is Nothing And Not skipG(k) Then
            Set ws = sh(k): Set ds = dsh(k)
            g = ws.Range(ws.Cells(1, 1), ws.Cells(nR, nC)).Formula
            For r = 1 To nR
                For c = 1 To nC
                    If Not skipM(r, c) Then
                        tv = CellText(t(r, c)): bv = CellText(b(r, c))
                        If tv <> bv Then
                            gv = ToTemplate(CellText(g(r, c)), ws, ds)
                            If gv = bv Then
                                upd(k).Add Array(r, c)
                            ElseIf gv <> tv Then
                                con(k).Add Array(r, c, CellText(g(r, c)))
                            End If
                        End If
                    End If
                Next c
            Next r
            nUpd = nUpd + upd(k).Count: nCon = nCon + con(k).Count
            rep = rep & "  " & ws.Name & ": " & upd(k).Count & U(" h\u00FCcre") & IIf(con(k).Count > 0, ", " & con(k).Count & U(" \u00E7ak\u0131\u015Fma"), "") & vbLf
            For Each it In con(k)
                If Len(conList) < 3000 Then conList = conList & "  " & ws.Name & "!" & tpl.Cells(it(0) + IIf(it(0) > pEnd(k), sft(k), 0), it(1)).Address(False, False) & _
                    U(": grupta ") & Left$(it(2), 60) & U(" | \u015Fablonda ") & Left$(CellText(t(it(0), it(1))), 60) & vbLf
            Next it
        End If
    Next k
    lg = lg & U("\u015Eablonda son uygulamadan beri de\u011Fi\u015Fen h\u00FCcreler, grup sayfalar\u0131na:") & vbLf & rep
    If nCon > 0 Then lg = lg & U("Grup sayfas\u0131nda elle de\u011Fi\u015Ftirilmi\u015F (\u00E7ak\u0131\u015Fan) h\u00FCcreler:") & vbLf & conList

    If nUpd + nCon + chg.Count = 0 Then
        lg = lg & U("Yaz\u0131lacak bir \u015Fey yok.") & vbLf
        doApply = False
    ElseIf doApply And Not wasScripted Then
        Application.ScreenUpdating = True
        If Ask(lg & vbLf & nUpd & U(" h\u00FCcre \u015Fablona g\u00F6re yaz\u0131ls\u0131n m\u0131?")) <> vbYes Then doApply = False
        If doApply And nCon > 0 Then writeCon = (Ask(nCon & U(" \u00E7ak\u0131\u015Fan h\u00FCcre de \u015Fablona g\u00F6re yaz\u0131ls\u0131n m\u0131? (Hay\u0131r = grup sayfas\u0131ndaki de\u011Ferler kal\u0131r)")) = vbYes)
        Application.ScreenUpdating = False
    Else
        writeCon = force
    End If

    If doApply Then
        For k = 1 To nG
            If Not sh(k) Is Nothing Then
                For Each it In upd(k)
                    PutCell tpl, sh(k), dsh(k), it(0), it(1), CellText(t(it(0), it(1)))
                Next it
                If writeCon Then
                    For Each it In con(k)
                        PutCell tpl, sh(k), dsh(k), it(0), it(1), CellText(t(it(0), it(1)))
                    Next it
                End If
            End If
        Next k
        If chg.Count > 0 Then
            nMoved = 0
            For k = 1 To nG
                If Not sh(k) Is Nothing And Not skipG(k) Then
                    OpenRows sh(k)
                    nMoved = nMoved + MoveShapes(sh(k), chg, 0, 0, picMiss)
                End If
            Next k
            lg = lg & U("Yeri / boyutu d\u00FCzeltilen resim: ") & nMoved & vbLf
            If Len(picMiss) > 0 Then lg = lg & U("Bulunamayan (o sayfada silinmi\u015F veya de\u011Fi\u015Ftirilmi\u015F) resimler:") & vbLf & picMiss
            WriteShapeBase wb, tpl, True
        End If
        WriteBaseline wb, tpl, nR, nC
        lg = lg & U("Yaz\u0131ld\u0131: ") & (nUpd + IIf(writeCon, nCon, 0)) & U(" h\u00FCcre. \u0130z sayfas\u0131 \u015Fablonla e\u015Flendi.") & vbLf
    Else
        lg = lg & U("Hi\u00E7bir h\u00FCcre yaz\u0131lmad\u0131.") & vbLf
    End If

    ' Tablo 63 copies back, gates, print areas, numbers, lists
    If Application.Calculation <> cm Then Application.Calculation = cm
    shrunk = False
    lg = lg & U("Tablo 63 ek sayfalar\u0131 kar\u015F\u0131la\u015Ft\u0131rma i\u00E7in \u00E7\u0131kar\u0131l\u0131p yeniden kuruldu ('ek sayfa 0 ->' sat\u0131rlar\u0131):") & vbLf
    keep = lg
    tail = SAYFALARI_YENILE_SESSIZ()
    scripted = wasScripted
    If tplHid Then CloseRows tpl
    If Left$(tail, 5) = "FAIL:" Then Err.Raise vbObjectError + 42, , Mid$(tail, 6)
    lg = keep & tail
    Application.ScreenUpdating = True
    RunTemplateApply = lg
    Exit Function
Fail:
    eTx = Err.Description
    keep = lg
    On Error Resume Next
    If cm <> 0 Then
        If Application.Calculation <> cm Then Application.Calculation = cm
    End If
    If shrunk Then tail = SAYFALARI_YENILE_SESSIZ() Else tail = ""
    If tplHid Then CloseRows tpl
    scripted = wasScripted
    Application.ScreenUpdating = True
    RunTemplateApply = "FAIL:" & eTx & IIf(Len(keep) > 0, vbLf & keep, "") & IIf(Len(tail) > 0, vbLf & tail, "")
End Function

' rows the Tablo 63 copies add below the pattern page (M_T63_EK x pattern rows) and the pattern page's last row
Private Sub T63Shift(ws As Worksheet, ByRef nShift As Long, ByRef lastRow As Long)
    Dim pat As Range
    nShift = 0: lastRow = 0
    On Error Resume Next
    Set pat = ws.Names(M_T63_PAGE).RefersToRange
    On Error GoTo 0
    If pat Is Nothing Then Exit Sub
    lastRow = pat.Row + pat.Rows.Count - 1
    nShift = NameNumber(ws, M_T63_EXTRA) * pat.Rows.Count
End Sub

' a group sheet's formula in template terms: its own sheet -> SABLON_AYAK, its data sheet -> SABLON_MIDAS
Private Function ToTemplate(ByVal f As String, ws As Worksheet, ds As Worksheet) As String
    If Left$(f, 1) <> "=" Then ToTemplate = f: Exit Function
    f = ReplaceSheetRef(f, ws.Name, TEMPLATE)
    If Not ds Is Nothing Then f = ReplaceSheetRef(f, ds.Name, DATA_TPL)
    ToTemplate = f
End Function

' a template formula in the group's terms
Private Function ToGroup(ByVal f As String, ws As Worksheet, ds As Worksheet) As String
    If Left$(f, 1) <> "=" Then ToGroup = f: Exit Function
    f = ReplaceSheetRef(f, TEMPLATE, ws.Name)
    If Not ds Is Nothing Then f = ReplaceSheetRef(f, DATA_TPL, ds.Name)
    ToGroup = f
End Function

' fromName! or 'fromName'! -> toName in the form Excel writes it (quoted unless a plain name)
Private Function ReplaceSheetRef(ByVal f As String, ByVal fromName As String, ByVal toName As String) As String
    Dim tgt As String, plain As String, p As Long, pre As String
    If NeedsQuotes(toName) Then tgt = "'" & Replace(toName, "'", "''") & "'!" Else tgt = toName & "!"
    f = Replace(f, "'" & Replace(fromName, "'", "''") & "'!", tgt, , , vbTextCompare)
    If Not NeedsQuotes(fromName) Then
        plain = fromName & "!"
        p = InStr(1, f, plain, vbTextCompare)
        Do While p > 0
            pre = ""
            If p > 1 Then pre = Mid$(f, p - 1, 1)
            If pre Like "[A-Za-z0-9_.'!]" Then
                p = InStr(p + 1, f, plain, vbTextCompare)
            Else
                f = Left$(f, p - 1) & tgt & Mid$(f, p + Len(plain))
                p = InStr(p + Len(tgt), f, plain, vbTextCompare)
            End If
        Loop
    End If
    ReplaceSheetRef = f
End Function

' Excel writes a sheet name without quotes only when it holds ASCII letters, digits, _ and . , does not start
' with a digit and does not read as a cell address (letters then digits)
Private Function NeedsQuotes(ByVal nm As String) As Boolean
    Dim i As Long, ch As String, j As Long
    If Len(nm) = 0 Then NeedsQuotes = True: Exit Function
    If Left$(nm, 1) Like "[0-9.]" Then NeedsQuotes = True: Exit Function
    For i = 1 To Len(nm)
        ch = Mid$(nm, i, 1)
        If Not (ch Like "[A-Za-z0-9_.]") Then NeedsQuotes = True: Exit Function
    Next i
    j = 0
    Do While j < Len(nm)
        If Not (Mid$(nm, j + 1, 1) Like "[A-Za-z]") Then Exit Do
        j = j + 1
    Loop
    If j >= 1 And j <= 3 And j < Len(nm) Then
        If Mid$(nm, j + 1) Like String$(Len(nm) - j, "#") Then NeedsQuotes = True
    End If
End Function

' one group cell takes the template cell (number format too: a formula in a Text cell would stay text)
Private Sub PutCell(tpl As Worksheet, ws As Worksheet, ds As Worksheet, ByVal r As Long, ByVal c As Long, ByVal tv As String)
    Dim cl As Range
    Set cl = ws.Cells(r, c)
    If cl.NumberFormat <> tpl.Cells(r, c).NumberFormat Then cl.NumberFormat = tpl.Cells(r, c).NumberFormat
    If Len(tv) = 0 Then
        cl.ClearContents
    ElseIf Left$(tv, 1) = "=" Then
        cl.Formula = ToGroup(tv, ws, ds)
    Else
        cl.Value = tpl.Cells(r, c).Value
    End If
End Sub

' baseline = the template's formulas and constants as text (apostrophe prefix: nothing calculates there)
Private Sub WriteBaseline(wb As Workbook, tpl As Worksheet, ByVal nR As Long, ByVal nC As Long, Optional ByVal baseName As String = BASE_SHEET)
    Dim bs As Worksheet, t As Variant, r As Long, c As Long
    If SheetExists(wb, baseName) Then
        Set bs = wb.Worksheets(baseName)
        bs.Cells.ClearContents
    Else
        Set bs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        bs.Name = baseName
    End If
    t = tpl.Range(tpl.Cells(1, 1), tpl.Cells(nR, nC)).Formula
    For r = 1 To nR
        For c = 1 To nC
            If Len(CellText(t(r, c))) > 0 Then t(r, c) = "'" & CellText(t(r, c)) Else t(r, c) = Empty
        Next c
    Next r
    bs.Range(bs.Cells(1, 1), bs.Cells(nR, nC)).Value = t
    bs.Visible = xlSheetVeryHidden
End Sub

Private Sub MarkSkip(ws As Worksheet, ByVal nm As String, skipM() As Boolean, ByVal nR As Long, ByVal nC As Long)
    Dim rg As Range, cl As Range
    On Error Resume Next
    Set rg = ws.Names(nm).RefersToRange
    On Error GoTo 0
    If rg Is Nothing Then Exit Sub
    For Each cl In rg.Cells
        If cl.Row <= nR And cl.Column <= nC Then skipM(cl.Row, cl.Column) = True
    Next cl
End Sub

' ---- pictures and drawings: place and size only (16h, owner 2026-11-16) ----
' every gated row open, so that a picture's place and size are its real ones (hidden rows have no height)
Private Sub OpenRows(ws As Worksheet)
    Dim rg As Range
    Set rg = SheetRange(ws, M_GATE)
    If Not rg Is Nothing Then rg.EntireRow.Hidden = False
End Sub

Private Sub CloseRows(ws As Worksheet)
    On Error Resume Next
    ApplyGate ws
    On Error GoTo 0
End Sub

' "id|name|row|column|x in the cell|y in the cell|width|height" (points) of every picture / drawing; comments
' and controls are left out, as in GateShapes. The id stays the same on the copies of the template.
Private Function GeoList(ws As Worksheet) As Collection
    Dim c As New Collection, s As Shape, tl As Range, t0 As Double, l0 As Double, et As Double, el As Double
    For Each s In ws.Shapes
        Select Case s.Type
            Case msoComment, msoFormControl, msoOLEControlObject
            Case Else
                Set tl = Nothing
                On Error Resume Next
                Set tl = s.TopLeftCell
                On Error GoTo 0
                If Not tl Is Nothing Then
                    ' the cell's edge in the shape's own coordinates (see EdgeTop), then the shape back where it was
                    t0 = s.Top: l0 = s.Left
                    et = EdgeTop(s, tl.Row)
                    s.Top = t0
                    el = EdgeLeft(s, tl.Column)
                    s.Left = l0
                    c.Add s.ID & "|" & Replace(s.Name, "|", "/") & "|" & tl.Row & "|" & tl.Column & "|" & NumTx(l0 - el) & "|" & _
                        NumTx(t0 - et) & "|" & NumTx(s.Width) & "|" & NumTx(s.Height)
                End If
        End Select
    Next s
    Set GeoList = c
End Function

' Top at which the shape starts exactly at the top edge of row r. Measured 2026-11-16: on a sheet whose rows were
' just unhidden, Shape.Top and Range.Top do not agree (60-70 points apart near row 217, different on every sheet),
' so a place is never taken from Range.Top: the shape is moved until its TopLeftCell changes row. Rows must be open.
Private Function EdgeTop(s As Shape, ByVal r As Long) As Double
    Dim lo As Double, hi As Double, md As Double, span As Double, t0 As Double, i As Long
    If r <= 1 Then Exit Function
    t0 = s.Top: span = 64
    hi = t0 + span
    s.Top = hi
    Do While s.TopLeftCell.Row < r And span < 60000
        span = span * 2: hi = t0 + span: s.Top = hi
    Loop
    lo = t0 - span: If lo < 0 Then lo = 0
    s.Top = lo
    Do While s.TopLeftCell.Row >= r And lo > 0
        span = span * 2: lo = t0 - span: If lo < 0 Then lo = 0
        s.Top = lo
    Loop
    For i = 1 To 30
        If hi - lo < 0.3 Then Exit For
        md = (lo + hi) / 2
        s.Top = md
        If s.TopLeftCell.Row >= r Then hi = md Else lo = md
    Next i
    EdgeTop = hi
End Function

Private Function EdgeLeft(s As Shape, ByVal c As Long) As Double
    Dim lo As Double, hi As Double, md As Double, span As Double, l0 As Double, i As Long
    If c <= 1 Then Exit Function
    l0 = s.Left: span = 64
    hi = l0 + span
    s.Left = hi
    Do While s.TopLeftCell.Column < c And span < 60000
        span = span * 2: hi = l0 + span: s.Left = hi
    Loop
    lo = l0 - span: If lo < 0 Then lo = 0
    s.Left = lo
    Do While s.TopLeftCell.Column >= c And lo > 0
        span = span * 2: lo = l0 - span: If lo < 0 Then lo = 0
        s.Left = lo
    Loop
    For i = 1 To 30
        If hi - lo < 0.3 Then Exit For
        md = (lo + hi) / 2
        s.Left = md
        If s.TopLeftCell.Column >= c Then hi = md Else lo = md
    Next i
    EdgeLeft = hi
End Function

Private Function NumTx(ByVal v As Double) As String
    NumTx = Trim$(Str$(Round(v, 1)))
End Function

Private Function ReadShapeBase(wb As Workbook) As Collection
    Dim c As New Collection, ws As Worksheet, r As Long
    Set ws = wb.Worksheets(BASE_SHAPES)
    r = 1
    Do While Len(CellText(ws.Cells(r, 1).Value)) > 0
        c.Add CellText(ws.Cells(r, 1).Value)
        r = r + 1
    Loop
    Set ReadShapeBase = c
End Function

Private Sub WriteShapeList(wb As Workbook, lst As Collection)
    Dim bs As Worksheet, q As Variant, r As Long
    If SheetExists(wb, BASE_SHAPES) Then
        Set bs = wb.Worksheets(BASE_SHAPES)
        bs.Cells.ClearContents
    Else
        Set bs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        bs.Name = BASE_SHAPES
    End If
    For Each q In lst
        r = r + 1
        bs.Cells(r, 1).Value = "'" & CStr(q)
    Next q
    bs.Visible = xlSheetVeryHidden
End Sub

' the template's pictures as they are now into BASE_SHAPES; isOpen = its rows are open and stay so
Private Sub WriteShapeBase(wb As Workbook, tpl As Worksheet, Optional ByVal isOpen As Boolean = False)
    Dim lst As Collection
    If Not isOpen Then OpenRows tpl
    Set lst = GeoList(tpl)
    If Not isOpen Then
        If tpl.Visible <> xlSheetVisible Then CloseRows tpl
    End If
    WriteShapeList wb, lst
End Sub

' template pictures whose place or size changed since the last apply: Array(id, name, row, column, x, y, w, h);
' note = pictures that are new or gone on the template (not carried: the image itself is never copied)
Private Function ShapeChanges(wb As Workbook, tpl As Worksheet, ByRef note As String) As Collection
    Dim c As New Collection, cur As Collection, old As Collection, q As Variant, o As Variant
    Dim a() As String, b() As String, hit As Boolean, i As Long, dif As Boolean
    Set cur = GeoList(tpl)
    Set old = ReadShapeBase(wb)
    For Each q In cur
        a = Split(CStr(q), "|")
        hit = False
        For Each o In old
            b = Split(CStr(o), "|")
            If UBound(b) >= 7 Then
                If b(0) = a(0) Then
                    hit = True
                    dif = (b(2) <> a(2)) Or (b(3) <> a(3))
                    For i = 4 To 7
                        ' offsets 2 points; sizes 3 points + 1.5 % (a 'move and size' picture reads a few points wider or narrower from run to run)
                        If Abs(Val(b(i)) - Val(a(i))) > IIf(i <= 5, 2, 3 + 0.015 * Val(a(i))) Then dif = True
                    Next i
                    If dif Then c.Add Array(CLng(Val(a(0))), a(1), CLng(Val(a(2))), CLng(Val(a(3))), Val(a(4)), Val(a(5)), Val(a(6)), Val(a(7)))
                    Exit For
                End If
            End If
        Next o
        If Not hit Then note = note & "  + " & a(1) & " (" & tpl.Cells(CLng(Val(a(2))), CLng(Val(a(3)))).Address(False, False) & ")" & vbLf
    Next q
    For Each o In old
        b = Split(CStr(o), "|")
        If UBound(b) >= 7 Then
            hit = False
            For Each q In cur
                If Split(CStr(q), "|")(0) = b(0) Then hit = True: Exit For
            Next q
            If Not hit Then note = note & "  - " & b(1) & " (" & tpl.Cells(CLng(Val(b(2))), CLng(Val(b(3)))).Address(False, False) & ")" & vbLf
        End If
    Next o
    Set ShapeChanges = c
End Function

' the group sheet's own picture of the same id goes to the template's place and size (its image is not touched);
' rows below lastRow are shiftRows lower on this sheet. Both sheets have their rows open. Returns the number moved.
Private Function MoveShapes(ws As Worksheet, chg As Collection, ByVal shiftRows As Long, ByVal lastRow As Long, ByRef miss As String) As Long
    Dim q As Variant, s As Shape, hit As Shape, r As Long, c As Long, la As Long
    For Each q In chg
        Set hit = Nothing
        For Each s In ws.Shapes
            If s.ID = CLng(q(0)) Then
                If s.Name = CStr(q(1)) Then Set hit = s
                Exit For
            End If
        Next s
        If hit Is Nothing Then
            miss = miss & "  " & ws.Name & ": " & CStr(q(1)) & vbLf
        Else
            r = CLng(q(2)): c = CLng(q(3))
            If lastRow > 0 And r > lastRow Then r = r + shiftRows
            la = msoFalse
            On Error Resume Next
            la = hit.LockAspectRatio
            hit.LockAspectRatio = msoFalse
            On Error GoTo 0
            hit.Width = CDbl(q(6))
            hit.Height = CDbl(q(7))
            hit.Top = EdgeTop(hit, r) + CDbl(q(5))
            hit.Left = EdgeLeft(hit, c) + CDbl(q(4))
            On Error Resume Next
            hit.LockAspectRatio = la
            On Error GoTo 0
            MoveShapes = MoveShapes + 1
        End If
    Next q
End Function

' tests / install: the picture list of the last apply = the template of another file (a backup from before the
' pictures were moved), read without saving that file; no path = the template of this workbook as it is
Public Function SABLON_RESIM_IZ_SESSIZ(Optional ByVal path As String = "") As String
    Dim wb As Workbook, ob As Workbook, lst As Collection, ev As Boolean
    On Error GoTo Fail
    Set wb = ThisWorkbook
    Application.ScreenUpdating = False
    If Len(path) = 0 Then
        WriteShapeBase wb, wb.Worksheets(TEMPLATE)
    Else
        ev = Application.EnableEvents
        Application.EnableEvents = False
        Set ob = Workbooks.Open(path, 0, True)
        OpenRows ob.Worksheets(TEMPLATE)
        Set lst = GeoList(ob.Worksheets(TEMPLATE))
        ob.Close False
        Application.EnableEvents = ev
        WriteShapeList wb, lst
    End If
    Application.ScreenUpdating = True
    SABLON_RESIM_IZ_SESSIZ = BASE_SHAPES & ": " & ReadShapeBase(wb).Count
    Exit Function
Fail:
    Application.ScreenUpdating = True
    SABLON_RESIM_IZ_SESSIZ = "FAIL:" & Err.Description
End Function

' ---- data-sheet template to the data sheets (16i, owner 2026-11-16) ----
Public Sub VERI_SABLON_GOSTER()
    Dim wb As Workbook, tpl As Worksheet, ws As Worksheet, nR As Long, nC As Long
    Set wb = ThisWorkbook
    If Not SheetExists(wb, DATA_TPL) Then ShowMsg U("\u015Eablon sayfa yok: ") & DATA_TPL, vbExclamation: Exit Sub
    Set tpl = wb.Worksheets(DATA_TPL)
    If tpl.Visible = xlSheetVisible Then
        If ActiveSheet Is tpl Then
            For Each ws In wb.Worksheets
                If ws.Visible = xlSheetVisible And Not ws Is tpl Then ws.Activate: Exit For
            Next ws
        End If
        tpl.Visible = xlSheetVeryHidden
        ShowMsg U("Veri \u015Fablonu gizlendi. De\u011Fi\u015Fiklikleri veri sayfalar\u0131na ta\u015F\u0131mak i\u00E7in: 'Veri \u015Fablonunu veri sayfalar\u0131na uygula'."), vbInformation
    Else
        ' the first apply compares with the template as it is before the owner edits it
        If Not SheetExists(wb, DATA_BASE) Then
            Application.ScreenUpdating = False
            nR = tpl.UsedRange.Row + tpl.UsedRange.Rows.Count - 1
            nC = tpl.UsedRange.Column + tpl.UsedRange.Columns.Count - 1
            WriteBaseline wb, tpl, nR, nC, DATA_BASE
            Application.ScreenUpdating = True
        End If
        tpl.Visible = xlSheetVisible
        tpl.Activate
        ShowMsg U("Veri \u015Fablonu a\u00E7\u0131k: ") & DATA_TPL & vbLf & _
            U("Burada yaln\u0131z b\u00FCt\u00FCn gruplarda ge\u00E7erli olacak d\u00FCzeni (ba\u015Fl\u0131klar, h\u00FCcre yerleri, form\u00FCller, bi\u00E7im) de\u011Fi\u015Ftirin; gruba \u00F6zel GSD \u00E7\u0131kt\u0131s\u0131n\u0131 grubun MIDAS_ sayfas\u0131na yap\u0131\u015Ft\u0131r\u0131n. Sonra 'Veri \u015Fablonunu veri sayfalar\u0131na uygula'. ") & _
            U("Bir h\u00FCcrenin yeri de\u011Fi\u015Fince onu okuyan rapor form\u00FCl\u00FCn\u00FC SABLON_AYAK'ta d\u00FCzeltin."), vbInformation
    End If
End Sub

' tests: the same without the message
Public Sub VERI_SABLON_GOSTER_SESSIZ()
    Dim was As Boolean
    was = scripted
    scripted = True
    VERI_SABLON_GOSTER
    scripted = was
End Sub

Public Sub VERI_SABLON_UYGULA()
    Dim s As String
    scripted = False
    s = RunDataApply(True, False)
    If Left$(s, 5) = "FAIL:" Then ShowMsg Mid$(s, 6), vbExclamation Else ShowMsg s, vbInformation
End Sub

' tests: doApply False = list only (nothing written), force True = conflicts too
Public Function VERI_SABLON_UYGULA_SESSIZ(Optional ByVal doApply As Boolean = True, Optional ByVal force As Boolean = False) As String
    scripted = True
    VERI_SABLON_UYGULA_SESSIZ = RunDataApply(doApply, force)
End Function

Private Function RunDataApply(ByVal doApply As Boolean, ByVal force As Boolean) As String
    Dim wb As Workbook, tpl As Worksheet, bs As Worksheet, nG As Long, dsh() As Worksheet, ds As Worksheet
    Dim nR As Long, nC As Long, t As Variant, b As Variant, g As Variant, k As Long, r As Long, c As Long
    Dim tv As String, gv As String, bv As String, skipM() As Boolean, chg As Collection, upd() As Collection, con() As Collection
    Dim nUpd As Long, nCon As Long, it As Variant, rep As String, conList As String, writeCon As Boolean, cm As Long
    Dim ha As Variant, r1 As Long, r2 As Long, c1 As Long, c2 As Long, grown As Boolean, nm As Variant, eTx As String
    Dim wasScripted As Boolean, nSh As Long, rect As String, rg As Range, oldF As Variant, oldV As Variant, wr() As Boolean, nKept As Long
    Dim cow As Boolean, cowSet As Boolean
    On Error GoTo Fail
    wasScripted = scripted
    lg = ""
    Set wb = ThisWorkbook
    If Not SheetExists(wb, DATA_TPL) Then Err.Raise vbObjectError + 43, , U("\u015Eablon sayfa yok: ") & DATA_TPL
    Set tpl = wb.Worksheets(DATA_TPL)
    cm = Application.Calculation
    Application.ScreenUpdating = False
    nR = tpl.UsedRange.Row + tpl.UsedRange.Rows.Count - 1
    nC = tpl.UsedRange.Column + tpl.UsedRange.Columns.Count - 1
    If Not SheetExists(wb, DATA_BASE) Then
        WriteBaseline wb, tpl, nR, nC, DATA_BASE
        Application.ScreenUpdating = True
        RunDataApply = U("\u0130z sayfas\u0131 (") & DATA_BASE & U(") yoktu: veri \u015Fablonunun bug\u00FCnk\u00FC h\u00E2li kaydedildi, hi\u00E7bir veri sayfas\u0131 de\u011Fi\u015Fmedi. Bundan sonra \u015Fablonda yapt\u0131\u011F\u0131n\u0131z de\u011Fi\u015Fiklikler bu tu\u015Fla veri sayfalar\u0131na ta\u015F\u0131n\u0131r.")
        Exit Function
    End If
    Set bs = wb.Worksheets(DATA_BASE)
    With bs.UsedRange
        If .Row + .Rows.Count - 1 > nR Then nR = .Row + .Rows.Count - 1
        If .Column + .Columns.Count - 1 > nC Then nC = .Column + .Columns.Count - 1
    End With
    nG = NameRange(wb, N_GTAB).Rows.Count
    FindDataSheets wb, nG, dsh
    t = tpl.Range(tpl.Cells(1, 1), tpl.Cells(nR, nC)).Formula
    b = bs.Range(bs.Cells(1, 1), bs.Cells(nR, nC)).Value
    ' never compared: the cells MIDAS_AYAK writes (names on the template or on any data sheet), array formulas
    ReDim skipM(1 To nR, 1 To nC)
    For Each nm In Array(M_BLOCK, M_READ, "M_TESIR_STATIK", "M_TESIR_DEPREM", "M_TESIR_ELEMAN")
        MarkSkipArea tpl, CStr(nm), skipM, nR, nC
        For k = 1 To nG
            If Not dsh(k) Is Nothing Then MarkSkipArea dsh(k), CStr(nm), skipM, nR, nC
        Next k
    Next nm
    ha = tpl.Range(tpl.Cells(1, 1), tpl.Cells(nR, nC)).HasArray
    If IsNull(ha) Then
        For r = 1 To nR
            For c = 1 To nC
                If Left$(CellText(t(r, c)), 1) = "=" Then
                    If tpl.Cells(r, c).HasArray Then skipM(r, c) = True
                End If
            Next c
        Next r
    ElseIf ha = True Then
        Err.Raise vbObjectError + 44, , U("\u015Eablonun tamam\u0131 dizi form\u00FCl\u00FC; uygulanamaz.")
    End If

    ' cells changed on the template since the last apply, and the rectangle around them
    Set chg = New Collection
    For r = 1 To nR
        For c = 1 To nC
            If Not skipM(r, c) Then
                If CellText(t(r, c)) <> CellText(b(r, c)) Then
                    chg.Add Array(r, c)
                    If chg.Count = 1 Then r1 = r: r2 = r: c1 = c: c2 = c
                    If r > r2 Then r2 = r
                    If c < c1 Then c1 = c
                    If c > c2 Then c2 = c
                End If
            End If
        Next c
    Next r
    If chg.Count = 0 Then
        Application.ScreenUpdating = True
        RunDataApply = U("Veri \u015Fablonunda son uygulamadan beri de\u011Fi\u015Fen h\u00FCcre yok; veri sayfalar\u0131na dokunulmad\u0131. (Yaln\u0131z bi\u00E7im de\u011Fi\u015Fikli\u011Fi ta\u015F\u0131nmaz: bir h\u00FCcrenin i\u00E7eri\u011Fi de de\u011Fi\u015Fmeli.)")
        Exit Function
    End If
    ' the rectangle may not cut a merged area of the template or of a data sheet
    Do
        grown = False
        GrowRect tpl, r1, r2, c1, c2, grown
        For k = 1 To nG
            If Not dsh(k) Is Nothing Then GrowRect dsh(k), r1, r2, c1, c2, grown
        Next k
    Loop While grown
    rect = tpl.Range(tpl.Cells(r1, c1), tpl.Cells(r2, c2)).Address(False, False)

    ReDim upd(1 To nG): ReDim con(1 To nG)
    For k = 1 To nG
        Set upd(k) = New Collection: Set con(k) = New Collection
        If Not dsh(k) Is Nothing Then
            Set ds = dsh(k)
            nSh = nSh + 1
            g = ds.Range(ds.Cells(1, 1), ds.Cells(nR, nC)).Formula
            For Each it In chg
                r = it(0): c = it(1)
                tv = CellText(t(r, c)): bv = CellText(b(r, c)): gv = CellText(g(r, c))
                If gv = bv Then
                    upd(k).Add Array(r, c)
                ElseIf gv <> tv Then
                    con(k).Add Array(r, c, gv)
                End If
            Next it
            nUpd = nUpd + upd(k).Count: nCon = nCon + con(k).Count
            rep = rep & "  " & ds.Name & ": " & upd(k).Count & U(" h\u00FCcre") & IIf(con(k).Count > 0, ", " & con(k).Count & U(" \u00E7ak\u0131\u015Fma"), "") & vbLf
            For Each it In con(k)
                If Len(conList) < 3000 Then conList = conList & "  " & ds.Name & "!" & tpl.Cells(it(0), it(1)).Address(False, False) & _
                    U(": sayfada ") & Left$(it(2), 60) & U(" | \u015Fablonda ") & Left$(CellText(t(it(0), it(1))), 60) & vbLf
            Next it
        End If
    Next k
    lg = lg & U("Veri \u015Fablonunda son uygulamadan beri de\u011Fi\u015Fen ") & chg.Count & U(" h\u00FCcre (alan ") & rect & U("), veri sayfalar\u0131na:") & vbLf & rep
    If nCon > 0 Then lg = lg & U("Veri sayfas\u0131nda elle de\u011Fi\u015Ftirilmi\u015F (\u00E7ak\u0131\u015Fan) h\u00FCcreler:") & vbLf & conList
    If nSh = 0 Then
        lg = lg & U("Veri sayfas\u0131 yok.") & vbLf
        doApply = False
    ElseIf doApply And Not wasScripted Then
        Application.ScreenUpdating = True
        If Ask(lg & vbLf & nUpd & U(" h\u00FCcre \u015Fablona g\u00F6re yaz\u0131ls\u0131n, ") & rect & U(" alan\u0131n\u0131n bi\u00E7imi ve birle\u015Ftirilmi\u015F h\u00FCcreleri \u015Fablondan al\u0131ns\u0131n m\u0131?")) <> vbYes Then doApply = False
        If doApply And nCon > 0 Then writeCon = (Ask(nCon & U(" \u00E7ak\u0131\u015Fan h\u00FCcre de \u015Fablona g\u00F6re yaz\u0131ls\u0131n m\u0131? (Hay\u0131r = veri sayfas\u0131ndaki de\u011Ferler kal\u0131r)")) = vbYes)
        Application.ScreenUpdating = False
    Else
        writeCon = force
    End If

    If doApply Then
        If Application.Calculation <> xlCalculationManual Then Application.Calculation = xlCalculationManual
        Application.DisplayAlerts = False
        ' pictures must not travel with the cells: a data sheet keeps its own
        cow = Application.CopyObjectsWithCells: cowSet = True
        Application.CopyObjectsWithCells = False
        For k = 1 To nG
            If Not dsh(k) Is Nothing Then
                Set ds = dsh(k)
                ' the template's rectangle over the data sheet's (cells, formats, merged cells; no clipboard), then
                ' every cell that is not to be written gets its own content back
                Set rg = ds.Range(ds.Cells(r1, c1), ds.Cells(r2 + 1, c2))
                oldF = rg.Formula: oldV = rg.Value2
                ReDim wr(r1 To r2, c1 To c2)
                For Each it In upd(k)
                    wr(it(0), it(1)) = True
                Next it
                If writeCon Then
                    For Each it In con(k)
                        wr(it(0), it(1)) = True
                    Next it
                End If
                ds.Range(ds.Cells(r1, c1), ds.Cells(r2, c2)).UnMerge
                tpl.Range(tpl.Cells(r1, c1), tpl.Cells(r2, c2)).Copy Destination:=ds.Cells(r1, c1)
                For r = r1 To r2
                    For c = c1 To c2
                        If Not wr(r, c) Then
                            gv = CellText(oldF(r - r1 + 1, c - c1 + 1))
                            If gv <> CellText(t(r, c)) Then
                                On Error Resume Next
                                If Len(gv) = 0 Then
                                    ds.Cells(r, c).ClearContents
                                ElseIf Left$(gv, 1) = "=" Then
                                    ds.Cells(r, c).Formula = gv
                                Else
                                    ds.Cells(r, c).Value2 = oldV(r - r1 + 1, c - c1 + 1)
                                End If
                                On Error GoTo Fail
                                nKept = nKept + 1
                            End If
                        End If
                    Next c
                Next r
            End If
        Next k
        Application.CopyObjectsWithCells = cow: cowSet = False
        Application.DisplayAlerts = True
        WriteBaseline wb, tpl, nR, nC, DATA_BASE
        If Application.Calculation <> cm Then Application.Calculation = cm
        lg = lg & U("Yaz\u0131ld\u0131: ") & (nUpd + IIf(writeCon, nCon, 0)) & U(" h\u00FCcre; bi\u00E7im ve birle\u015Ftirilmi\u015F h\u00FCcreler ") & rect & U(" alan\u0131nda \u015Fablondan al\u0131nd\u0131. \u0130z sayfas\u0131 \u015Fablonla e\u015Flendi.") & vbLf
    Else
        lg = lg & U("Hi\u00E7bir h\u00FCcre yaz\u0131lmad\u0131.") & vbLf
    End If
    Application.ScreenUpdating = True
    RunDataApply = lg
    Exit Function
Fail:
    eTx = Err.Description
    On Error Resume Next
    Application.CutCopyMode = False
    If cowSet Then Application.CopyObjectsWithCells = cow
    Application.DisplayAlerts = True
    If cm <> 0 Then
        If Application.Calculation <> cm Then Application.Calculation = cm
    End If
    scripted = wasScripted
    Application.ScreenUpdating = True
    RunDataApply = "FAIL:" & eTx & IIf(Len(lg) > 0, vbLf & lg, "")
End Function

' cells of a sheet-level name (one area) marked as not compared
Private Sub MarkSkipArea(ws As Worksheet, ByVal nm As String, skipM() As Boolean, ByVal nR As Long, ByVal nC As Long)
    Dim rg As Range, r As Long, c As Long, rE As Long, cE As Long
    On Error Resume Next
    Set rg = ws.Names(nm).RefersToRange
    On Error GoTo 0
    If rg Is Nothing Then Exit Sub
    rE = rg.Row + rg.Rows.Count - 1: If rE > nR Then rE = nR
    cE = rg.Column + rg.Columns.Count - 1: If cE > nC Then cE = nC
    For r = rg.Row To rE
        For c = rg.Column To cE
            skipM(r, c) = True
        Next c
    Next r
End Sub

' the rectangle grown until no merged area of ws crosses its edge
Private Sub GrowRect(ws As Worksheet, ByRef r1 As Long, ByRef r2 As Long, ByRef c1 As Long, ByRef c2 As Long, ByRef grown As Boolean)
    Dim again As Boolean, r As Long, c As Long
    Do
        again = False
        For c = c1 To c2
            GrowCell ws.Cells(r1, c), r1, r2, c1, c2, again
            GrowCell ws.Cells(r2, c), r1, r2, c1, c2, again
        Next c
        For r = r1 To r2
            GrowCell ws.Cells(r, c1), r1, r2, c1, c2, again
            GrowCell ws.Cells(r, c2), r1, r2, c1, c2, again
        Next r
        If again Then grown = True
    Loop While again
End Sub

Private Sub GrowCell(cl As Range, ByRef r1 As Long, ByRef r2 As Long, ByRef c1 As Long, ByRef c2 As Long, ByRef again As Boolean)
    Dim m As Range
    If Not cl.MergeCells Then Exit Sub
    Set m = cl.MergeArea
    If m.Row < r1 Then r1 = m.Row: again = True
    If m.Column < c1 Then c1 = m.Column: again = True
    If m.Row + m.Rows.Count - 1 > r2 Then r2 = m.Row + m.Rows.Count - 1: again = True
    If m.Column + m.Columns.Count - 1 > c2 Then c2 = m.Column + m.Columns.Count - 1: again = True
End Sub
