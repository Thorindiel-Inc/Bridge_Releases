Attribute VB_Name = "TEMEL_GRUP"
Option Explicit
' Foundation report workbook (8'8_...) - 2026-11-26a
' One set of sheets per foundation group (owner 2026-11-26, docs/global-input/foundation-plan.md), made and kept in step
' from very hidden templates the way ORTA_AYAK does it for the 8'7 pier groups. Every cell is found by a name of the
' contract in foundation-plan.md, never by an address.
' TEMEL_GRUP_SAYFALARI: per foundation group
'   groups        = the typed Grup No column beside the pier table (T_GRUP_NO beside G_AYAK_TABLOSU); a pier without
'                   a group number is not in the report; group name = column 2 of T_GRUP_TABLO
'   one group     = one foundation: Temel Tip (KAZIKLI / YUZEYSEL), Kalinlik, En, Boy and Kazik Capi (pier table columns
'                   7-11) of all its piers equal; the tool stops and names the piers otherwise
'   group sheet   = copy of SABLON_TEMEL named <group>_<pier runs> without spaces (1_P1, 2_P2-P3,P16-P17), at most 31
'                   characters (cut with an ellipsis), <group>_TEMEL without a name; found by the group number in the
'                   first cell of its sheet-level name M_TEMEL_ZINCIR, renamed when its piers change
'   landscape     = copy of SABLON_TEMEL_Y named <group sheet>_Y (cut keeping _Y), found by its sheet-level number name
'                   M_TEMEL_Y_NO, right after its group sheet
'   data sheet    = copy of SABLON_MIDAS named MIDAS_<pier runs>, found by its sheet-level number name M_TESIR_NO, in
'                   group order right before INPUT
'   copies        = on the group's three sheets every formula reference to SABLON_TEMEL / SABLON_TEMEL_Y / SABLON_MIDAS is
'                   pointed at the group's own sheet of that kind; an existing sheet is never copied again (typed values
'                   stay); the sheets of a removed group are deleted after a question; renames go through free
'                   temporary names (two passes), so two groups that swap names never clash
'   order         = data sheets before INPUT; after AYRAC: 1_.., 1_.._Y, 2_.., 2_.._Y, ...; GEO last
'   pages         = rows whose gate (M_TEMEL_KAPI) is 0 are hidden with the pictures anchored on them, print area = the
'                   open pages (page index M_TEMEL_SAYFA, printed columns M_TEMEL_BASKI); a _Y sheet prints all its pages
'                   when it has these two names, else it keeps the print area of its template
'   numbers       = M_TEMEL_ZINCIR (2 rows x 4): row 1 = group, last table / figure / page before the sheet, row 2 = the
'                   sheet's own last ones; the first group continues from G_ONCEKI_NUMARA, each next one from the group
'                   sheet before it, G_SON_NUMARA = the last group's
'   per-pile table = its pattern page (M_TEMEL_KAZIK_SAYFA) is copied below itself once for every 36 piles over the first
'                   page (pile count = M_KAZIK_ADET of the group's data sheet; no copy for a YUZEYSEL group); each copy
'                   continues the table (relative row references to the data sheet moved), the page numbers and the
'                   page index; number of copies = M_TEMEL_KAZIK_EK
'   GEO           = rebuilt on every run: one block per group, stacked, copied from SABLON_GEO_KAZIK or SABLON_GEO_YUZEYSEL
'                   by the group's type, references pointed at the group's sheets, per-pile rows (the block template's
'                   M_KAZIK_GEO) grown with the pile count, a page break before every block, print area = all blocks
'   then the contents and the figure / table lists are rebuilt (RAPOR_LISTE).
'   The window asks first whether the base forces are pulled from 8'7 in the same run (TEMEL_GRUP_TESIR).
' TEMEL_GRUP_YENILE_SESSIZ: per-pile pages, gates, print areas, numbers, GEO and lists of the existing sheets again; nothing
'   made, renamed or deleted (for MIDAS_KAZIK and GSD_OKU after they write). Returns the log.
' TEMEL_GRUP_SAYFALARI_SESSIZ [doPull, 8'7 file] / TEMEL_GRUP_SAYFALARI_RUN report [, doPull, 8'7 file]: the group run
'   without questions (deletes without asking); returns the log / writes it to a text file.
' TEMEL_GRUP_TESIR: base forces from the 8'7 workbook in this workbook's folder (8'7_*.xlsm; exactly one, else a file
'   dialog), opened read-only without updating links and closed without saving (left open when it was open already):
'   for every pier of a group the 8'7 data sheet whose element list (M_TESIR_ELEMAN) names the pier, and its rows of
'   M_TESIR_BLOK at the column base (J end), into M_TEMEL_TESIR of the group's data sheet (pier, combination, element,
'   end, Pu, My, Mz, Vy, Vz, T); rows not used are cleared; stamp file | saved | read in M_TEMEL_TESIR_DAMGA.
'   TEMEL_GRUP_TESIR_SESSIZ [file] / TEMEL_GRUP_TESIR_RUN report [, file]: the same without dialogs.
' TEMEL_GRUP_SABLON_GOSTER: shows SABLON_TEMEL, SABLON_TEMEL_Y and the GEO block templates for editing (every gated row
'   and picture open); the next press hides them again (GEO is rebuilt from its block templates on every run).
' TEMEL_GRUP_SABLON_UYGULA: carries a change made once on SABLON_TEMEL to every group sheet and on SABLON_TEMEL_Y to every
'   _Y sheet, three-way against the baselines SABLON_TEMEL_IZ / SABLON_TEMEL_Y_IZ (the template at the last apply):
'     template = baseline                       -> nothing (a value typed on a sheet stays)
'     template changed, sheet = baseline        -> the sheet takes the template cell, references pointed at the group's sheets
'     template changed, sheet = template        -> nothing
'     template changed, sheet = something else  -> conflict: listed, written only when the owner says yes
'   Formulas are compared in template terms; the per-pile copies are taken out first and made again at the end with the
'   gates, print areas, numbers, GEO and lists; chain cells, page index and array formulas are not compared. A picture moved
'   or resized on a template (lists SABLON_TEMEL_IZ_R / SABLON_TEMEL_Y_IZ_R) puts the sheet's own picture of the same id at
'   the same place and size (its image is never copied); new or deleted pictures are only listed. Without a baseline the
'   tool writes it and changes nothing. TEMEL_GRUP_SABLON_UYGULA_SESSIZ [doApply, force]: without questions.
' TEMEL_GRUP_VERI_SABLON_GOSTER / TEMEL_GRUP_VERI_SABLON_UYGULA: the same for SABLON_MIDAS and the MIDAS_ data sheets
'   (baseline SABLON_MIDAS_IZ written at the first show; cells, formats and merged cells of the rectangle around the changed
'   cells; the areas of the names M_TESIR_*, M_KAZIK_*, M_GSD_* and M_TEMEL_TESIR* are never compared; pictures are not
'   copied). TEMEL_GRUP_VERI_SABLON_UYGULA_SESSIZ [doApply, force]: without questions.
' TEMEL_GRUP_SABLON_RESIM_IZ_SESSIZ [file]: install helper, picture lists of the last apply from this or another file.
' TEMEL_GRUP_TEST_RUN report [, steps, 8'7 file]: scripted test, several entries in a row without dialogs; steps separated
'   by | from SAYFA, SAYFA_TESIR, TESIR, YENILE, SABLON (list only), SABLON_YAZ, SABLON_ZORLA, VERI (list only), VERI_YAZ,
'   VERI_ZORLA, GOSTER (show and hide again), VERI_GOSTER.

Private Const SCRIPT_ID As String = "foundation-sheets"
Private Const SCRIPT_VERSION As String = "2026-11-26a"
Private Const SCRIPT_CHANGELOG As String = "new: one report sheet, landscape companion (_Y) and data sheet (MIDAS_) per foundation group of the 8'8 workbook from the templates SABLON_TEMEL / SABLON_TEMEL_Y / SABLON_MIDAS; per-pile pages grow with the pile count; GEO sheet with one block per group; base forces pulled from the 8'7 data sheets; template tools (show / hide, three-way apply)"
' macro-window row(s) of this module (PANEL reads them): order|group|title|macro|sheet|text|fields[|modules]
'@TOOL 40|Girdiler|Temel grubu sayfalar\u0131|TEMEL_GRUP.TEMEL_GRUP_SAYFALARI||Ayak tablosundaki Grup No s\u00FCtununa (T_GRUP_NO) g\u00F6re her temel grubu i\u00E7in bir rapor sayfas\u0131, yatay sayfas\u0131 (_Y) ve veri sayfas\u0131 (MIDAS_) a\u00E7ar. Bir gruptaki ayaklar\u0131n temel tipi, kal\u0131nl\u0131\u011F\u0131, eni, boyu ve kaz\u0131k \u00E7ap\u0131 ayn\u0131 olmal\u0131. Sayfa ad\u0131 grup numaras\u0131 ve grubun ayaklar\u0131d\u0131r; ayaklar de\u011Fi\u015Fince ad da de\u011Fi\u015Fir, sayfadaki de\u011Ferlere dokunulmaz. Temel tipine g\u00F6re sayfalar\u0131 g\u00F6sterir / gizler, kaz\u0131k tablosunu kaz\u0131k say\u0131s\u0131na g\u00F6re b\u00FCy\u00FCt\u00FCr, GEO sayfas\u0131n\u0131, numaralar\u0131 ve listeleri yeniler. Temel tesirlerinin 8'7'den de \u00E7ekilip \u00E7ekilmeyece\u011Fini ba\u015Fta sorar.|
'@TOOL 41|Girdiler|Temel \u015Fablonlar\u0131n\u0131 g\u00F6ster / gizle|TEMEL_GRUP.TEMEL_GRUP_SABLON_GOSTER||Gizli temel \u015Fablonlar\u0131n\u0131 (SABLON_TEMEL, SABLON_TEMEL_Y ve GEO bloklar\u0131) d\u00FCzenlemek i\u00E7in a\u00E7ar; tekrar bas\u0131l\u0131nca gizler. B\u00FCt\u00FCn grup sayfalar\u0131na gidecek bir d\u00FCzeltmeyi \u015Fablonda bir kez yap\u0131n, sonra 'Temel \u015Fablonlar\u0131n\u0131 grup sayfalar\u0131na uygula'. GEO sayfas\u0131 her \u00E7al\u0131\u015Fmada bloklar\u0131ndan yeniden kurulur.|
'@TOOL 42|Girdiler|Temel \u015Fablonlar\u0131n\u0131 grup sayfalar\u0131na uygula|TEMEL_GRUP.TEMEL_GRUP_SABLON_UYGULA||SABLON_TEMEL'de ve SABLON_TEMEL_Y'de son uygulamadan beri de\u011Fi\u015Fen h\u00FCcreleri her grup sayfas\u0131na ve _Y sayfas\u0131na yazar; form\u00FCller grubun kendi sayfalar\u0131na ba\u011Flan\u0131r. Sayfada elle de\u011Fi\u015Ftirdi\u011Finiz h\u00FCcreler listelenir ve yaln\u0131z siz evet derseniz yaz\u0131l\u0131r. Yazmadan \u00F6nce sayfa ba\u015F\u0131na ka\u00E7 h\u00FCcre de\u011Fi\u015Fece\u011Fini g\u00F6sterir.|
'@TOOL 43|Girdiler|Temel veri \u015Fablonunu g\u00F6ster / gizle|TEMEL_GRUP.TEMEL_GRUP_VERI_SABLON_GOSTER||Gizli veri \u015Fablonunu (SABLON_MIDAS: kaz\u0131k ba\u015F\u0131 y\u00FCkleri, GSD \u00E7\u0131kt\u0131s\u0131 ve temel tesirleri) d\u00FCzenlemek i\u00E7in a\u00E7ar; tekrar bas\u0131l\u0131nca gizler. B\u00FCt\u00FCn MIDAS_ veri sayfalar\u0131na gidecek bir d\u00FCzen de\u011Fi\u015Fikli\u011Fini burada bir kez yap\u0131n, sonra 'Temel veri \u015Fablonunu veri sayfalar\u0131na uygula'.|
'@TOOL 44|Girdiler|Temel veri \u015Fablonunu veri sayfalar\u0131na uygula|TEMEL_GRUP.TEMEL_GRUP_VERI_SABLON_UYGULA||SABLON_MIDAS'ta son uygulamadan beri de\u011Fi\u015Fen h\u00FCcreleri, bi\u00E7imleri ve birle\u015Ftirilmi\u015F h\u00FCcreleri her MIDAS_ veri sayfas\u0131na ta\u015F\u0131r. Veri sayfas\u0131nda sizin de\u011Fi\u015Ftirdi\u011Finiz h\u00FCcreler listelenir ve yaln\u0131z siz evet derseniz yaz\u0131l\u0131r. Kaz\u0131k, GSD ve tesir bloklar\u0131na dokunulmaz. Yazmadan \u00F6nce sorar.|
'@TOOL 45|Girdiler|Temel tesirleri (8'7)|TEMEL_GRUP.TEMEL_GRUP_TESIR||Bu dosyan\u0131n klas\u00F6r\u00FCndeki 8'7 orta ayak dosyas\u0131n\u0131 salt okunur a\u00E7ar ve her grubun ayaklar\u0131n\u0131n kolon taban\u0131 tesirlerini (8'7 veri sayfalar\u0131ndaki MIDAS_AYAK sat\u0131rlar\u0131) grubun MIDAS_ veri sayfas\u0131na yazar; dosyay\u0131 kaydetmeden kapat\u0131r. Klas\u00F6rde tek 8'7 dosyas\u0131 yoksa dosyay\u0131 sorar.|

Private Const TPL_G As String = "SABLON_TEMEL"          ' very hidden portrait template of a group sheet
Private Const TPL_Y As String = "SABLON_TEMEL_Y"        ' very hidden landscape template of a group's _Y sheet
Private Const TPL_D As String = "SABLON_MIDAS"          ' very hidden template of a group's data sheet
Private Const TPL_GEO_K As String = "SABLON_GEO_KAZIK"  ' very hidden GEO block of a piled group
Private Const TPL_GEO_Y As String = "SABLON_GEO_YUZEYSEL"  ' very hidden GEO block of a shallow group
Private Const TPL_PFX As String = "SABLON_"             ' every template name starts with it
Private Const GEO_SHEET As String = "GEO"
Private Const BASE_G As String = "SABLON_TEMEL_IZ"      ' very hidden: SABLON_TEMEL at the last apply (as text)
Private Const BASE_G_R As String = "SABLON_TEMEL_IZ_R"  ' very hidden: place and size of its pictures at the last apply
Private Const BASE_Y As String = "SABLON_TEMEL_Y_IZ"
Private Const BASE_Y_R As String = "SABLON_TEMEL_Y_IZ_R"
Private Const BASE_D As String = "SABLON_MIDAS_IZ"
Private Const SHEET_SFX As String = "_TEMEL"            ' name of a group sheet whose group has no name
Private Const Y_SFX As String = "_Y"
Private Const TMP_BASE As String = "~tmp_TEMEL_"        ' temporary names during the renames
Private Const MAX_NAME As Long = 31                     ' Excel's limit for a sheet name
Private Const AFTER_SHEET As String = "AYRAC"           ' the first group sheet goes right after this sheet
Private Const DATA_BEFORE As String = "INPUT"           ' data sheets go right before this sheet
Private Const DATA_PFX As String = "MIDAS_"             ' data sheet name = MIDAS_<pier runs>
Private Const N_TABLE As String = "G_AYAK_TABLOSU"
Private Const N_GROUP As String = "T_GRUP_NO"
Private Const N_GTAB As String = "T_GRUP_TABLO"
Private Const N_PREV As String = "G_ONCEKI_NUMARA"
Private Const N_LAST As String = "G_SON_NUMARA"
Private Const COL_SEC As Long = 3                       ' pier table: section type (KA = abutment)
Private Const COL_TYPE As Long = 7                      ' pier table: Temel Tip, then Kalinlik, En, Boy, Kazik Capi
Private Const COL_PILE As Long = 11
Private Const TYPE_PILED As String = "KAZIKLI"
Private Const TYPE_SHALLOW As String = "YUZEYSEL"
Private Const M_CHAIN As String = "M_TEMEL_ZINCIR"      ' sheet-level names of SABLON_TEMEL, copied with it
Private Const M_GATE As String = "M_TEMEL_KAPI"
Private Const M_PAGE As String = "M_TEMEL_SAYFA"
Private Const M_PRINT As String = "M_TEMEL_BASKI"
Private Const M_FIRST As String = "M_ILK_SAYFA"         ' first page number cell: its column = the page-number column
Private Const M_PILE_PAGE As String = "M_TEMEL_KAZIK_SAYFA"   ' rows of the pattern page of the per-pile table
Private Const M_PILE_EXTRA As String = "M_TEMEL_KAZIK_EK"     ' number of pattern copies below it
Private Const M_Y_NO As String = "M_TEMEL_Y_NO"         ' _Y sheet: sheet-level name holding its group number
Private Const M_DATA_NO As String = "M_TESIR_NO"        ' data sheet: sheet-level name holding its group number
Private Const M_PILES As String = "M_KAZIK_ADET"        ' data sheet: pile count (cell)
Private Const M_GEO_ROWS As String = "M_KAZIK_GEO"      ' GEO block template: its per-pile rows
Private Const M_BASE As String = "M_TEMEL_TESIR"        ' data sheet: base-force block (10 columns)
Private Const M_BASE_STAMP As String = "M_TEMEL_TESIR_DAMGA"
Private Const DATA_SKIP As String = "M_TESIR_*|M_KAZIK_*|M_GSD_*|M_TEMEL_TESIR*"   ' areas the data apply never compares
Private Const PILES_PER_PAGE As Long = 36
Private Const MAX_PILES As Long = 1000
Private Const OA_PATTERN As String = "8'7_*.xlsm"       ' the pier workbook in the same folder
Private Const OA_TPL As String = "SABLON_MIDAS"         ' 8'7: template of its data sheets (not read)
Private Const OA_BLOCK As String = "M_TESIR_BLOK"       ' 8'7 data sheet: counter, No, name, Pu, My, Mz, Vy, Vz, T
Private Const OA_ELEM As String = "M_TESIR_ELEMAN"      ' 8'7 data sheet: element list, 2nd column 'P1: 473-477, base 477 J'
Private Const BASE_END As String = "J"                  ' column base = J end of the lowest element
Private Const KIND_G As Long = 1
Private Const KIND_Y As Long = 2
Private Const KIND_D As Long = 3

Private scripted As Boolean
Private lg As String

' ==== entries =========================================================================================================
Public Sub TEMEL_GRUP_SAYFALARI()
    Dim s As String, doPull As Boolean
    scripted = False
    doPull = (Ask(U("Grup sayfalar\u0131ndan sonra temel tesirleri 8'7 orta ayak dosyas\u0131ndan da \u00E7ekilsin mi?") & vbLf & _
        U("(Her grubun ayaklar\u0131n\u0131n kolon taban\u0131 tesirleri grubun MIDAS_ veri sayfas\u0131na yaz\u0131l\u0131r; 8'7 dosyas\u0131 salt okunur a\u00E7\u0131l\u0131r.)")) = vbYes)
    s = RunSync(doPull, "")
    If Left$(s, 5) = "FAIL:" Then ShowMsg Mid$(s, 6), vbExclamation Else ShowMsg s, vbInformation
End Sub

Public Function TEMEL_GRUP_SAYFALARI_SESSIZ(Optional ByVal doPull As Boolean = False, Optional ByVal oaPath As String = "") As String
    scripted = True
    TEMEL_GRUP_SAYFALARI_SESSIZ = RunSync(doPull, oaPath)
End Function

Public Sub TEMEL_GRUP_SAYFALARI_RUN(ByVal reportFile As String, Optional ByVal doPull As Boolean = False, Optional ByVal oaPath As String = "")
    scripted = True
    WriteReport reportFile, RunSync(doPull, oaPath)
End Sub

' per-pile pages, gates, print areas, numbers, GEO and lists of the existing sheets again; returns the log
Public Function TEMEL_GRUP_YENILE_SESSIZ() As String
    scripted = True
    TEMEL_GRUP_YENILE_SESSIZ = RefreshTop()
End Function

Public Sub TEMEL_GRUP_TESIR()
    Dim s As String
    scripted = False
    s = RunPullTop("")
    If Left$(s, 5) = "FAIL:" Then ShowMsg Mid$(s, 6), vbExclamation Else ShowMsg s, vbInformation
End Sub

Public Function TEMEL_GRUP_TESIR_SESSIZ(Optional ByVal oaPath As String = "") As String
    scripted = True
    TEMEL_GRUP_TESIR_SESSIZ = RunPullTop(oaPath)
End Function

Public Sub TEMEL_GRUP_TESIR_RUN(ByVal reportFile As String, Optional ByVal oaPath As String = "")
    scripted = True
    WriteReport reportFile, RunPullTop(oaPath)
End Sub

Public Sub TEMEL_GRUP_SABLON_GOSTER()
    scripted = False
    ShowTemplates
End Sub

Public Sub TEMEL_GRUP_SABLON_GOSTER_SESSIZ()
    scripted = True
    ShowTemplates
End Sub

Public Sub TEMEL_GRUP_SABLON_UYGULA()
    Dim s As String
    scripted = False
    s = RunTemplateApply(True, False)
    If Left$(s, 5) = "FAIL:" Then ShowMsg Mid$(s, 6), vbExclamation Else ShowMsg s, vbInformation
End Sub

' tests: doApply False = list only (nothing written), force True = conflicts too
Public Function TEMEL_GRUP_SABLON_UYGULA_SESSIZ(Optional ByVal doApply As Boolean = True, Optional ByVal force As Boolean = False) As String
    scripted = True
    TEMEL_GRUP_SABLON_UYGULA_SESSIZ = RunTemplateApply(doApply, force)
End Function

Public Sub TEMEL_GRUP_VERI_SABLON_GOSTER()
    scripted = False
    ShowDataTemplate
End Sub

Public Sub TEMEL_GRUP_VERI_SABLON_GOSTER_SESSIZ()
    scripted = True
    ShowDataTemplate
End Sub

Public Sub TEMEL_GRUP_VERI_SABLON_UYGULA()
    Dim s As String
    scripted = False
    s = RunDataApply(True, False)
    If Left$(s, 5) = "FAIL:" Then ShowMsg Mid$(s, 6), vbExclamation Else ShowMsg s, vbInformation
End Sub

Public Function TEMEL_GRUP_VERI_SABLON_UYGULA_SESSIZ(Optional ByVal doApply As Boolean = True, Optional ByVal force As Boolean = False) As String
    scripted = True
    TEMEL_GRUP_VERI_SABLON_UYGULA_SESSIZ = RunDataApply(doApply, force)
End Function

' scripted test: steps separated by |, each step's log under its own heading in the report (UTF-8)
Public Sub TEMEL_GRUP_TEST_RUN(ByVal reportFile As String, Optional ByVal steps As String = "SAYFA|YENILE|SABLON|VERI", Optional ByVal oaPath As String = "")
    Dim st As Variant, r As String, s As String
    scripted = True
    For Each st In Split(steps, "|")
        Select Case UCase$(Trim$(CStr(st)))
            Case "SAYFA": r = RunSync(False, "")
            Case "SAYFA_TESIR": r = RunSync(True, oaPath)
            Case "TESIR": r = RunPullTop(oaPath)
            Case "YENILE": r = RefreshTop()
            Case "SABLON": r = RunTemplateApply(False, False)
            Case "SABLON_YAZ": r = RunTemplateApply(True, False)
            Case "SABLON_ZORLA": r = RunTemplateApply(True, True)
            Case "VERI": r = RunDataApply(False, False)
            Case "VERI_YAZ": r = RunDataApply(True, False)
            Case "VERI_ZORLA": r = RunDataApply(True, True)
            Case "GOSTER": ShowTemplates: ShowTemplates: r = "shown and hidden again"
            Case "VERI_GOSTER": ShowDataTemplate: ShowDataTemplate: r = "shown and hidden again"
            Case "": r = ""
            Case Else: r = "FAIL:unknown step " & CStr(st)
        End Select
        scripted = True
        s = s & "== " & Trim$(CStr(st)) & " ==" & vbLf & r & vbLf
    Next st
    WriteReport reportFile, s
End Sub

' install / tests: the picture lists of the last apply = the templates of another file (a backup from before pictures
' were moved), read without saving that file; no path = the templates of this workbook as they are
Public Function TEMEL_GRUP_SABLON_RESIM_IZ_SESSIZ(Optional ByVal path As String = "") As String
    Dim wb As Workbook, ob As Workbook, ev As Boolean, kd As Long, tn As String, bn As String, s As String
    On Error GoTo Fail
    scripted = True
    Set wb = ThisWorkbook
    Application.ScreenUpdating = False
    If Len(path) > 0 Then
        ev = Application.EnableEvents
        Application.EnableEvents = False
        Set ob = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
        Application.EnableEvents = ev
    End If
    For kd = 1 To 2
        If kd = 1 Then tn = TPL_G: bn = BASE_G_R Else tn = TPL_Y: bn = BASE_Y_R
        If ob Is Nothing Then
            If SheetExists(wb, tn) Then WriteShapeBase wb, wb.Worksheets(tn), bn: s = s & bn & ": " & ReadShapeBase(wb, bn).Count & vbLf
        ElseIf SheetExists(ob, tn) Then
            OpenRows ob.Worksheets(tn)
            WriteShapeList wb, GeoList(ob.Worksheets(tn)), bn
            s = s & bn & ": " & ReadShapeBase(wb, bn).Count & vbLf
        End If
    Next kd
    If Not ob Is Nothing Then ob.Close SaveChanges:=False
    Application.ScreenUpdating = True
    TEMEL_GRUP_SABLON_RESIM_IZ_SESSIZ = s
    Exit Function
Fail:
    s = Err.Description
    On Error Resume Next
    If Not ob Is Nothing Then ob.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    TEMEL_GRUP_SABLON_RESIM_IZ_SESSIZ = "FAIL:" & s
End Function

' ==== group run =======================================================================================================
Private Function RunSync(ByVal doPull As Boolean, ByVal oaPath As String) As String
    Dim wb As Workbook, groups As Collection, k As Variant, nG As Long, gs() As Worksheet, ys() As Worksheet, ds() As Worksheet
    Dim tps() As String, toDel As Collection, delList As String, s As Variant, madeG As String, madeY As String, madeD As String
    Dim t As String, n As Long, hasY As Boolean, hasD As Boolean, lastSh As Worksheet, pl As String, kk As Long
    Dim t0 As Double, tm As String
    On Error GoTo Fail
    Set wb = ThisWorkbook
    lg = ""
    Application.ScreenUpdating = False
    Set groups = ReadGroups(wb, tps)
    nG = NameRange(wb, N_GTAB).Rows.Count
    If Not SheetExists(wb, TPL_G) Then Err.Raise vbObjectError + 25, , U("\u015Eablon sayfa yok: ") & TPL_G
    hasY = SheetExists(wb, TPL_Y)
    hasD = SheetExists(wb, TPL_D)
    If Not hasY Then lg = lg & U("Not: ") & TPL_Y & U(" yok; _Y sayfalar\u0131 a\u00E7\u0131lmad\u0131.") & vbLf
    If Not hasD Then lg = lg & U("Not: ") & TPL_D & U(" yok; grup veri sayfalar\u0131 (") & DATA_PFX & U("<ayaklar>) a\u00E7\u0131lmad\u0131.") & vbLf

    ' sheets by their group number; sheets of groups that no longer exist
    Set toDel = New Collection
    CollectRemoved toDel, FindSheets(wb, nG, KIND_G, gs), gs, groups
    CollectRemoved toDel, FindSheets(wb, nG, KIND_Y, ys), ys, groups
    CollectRemoved toDel, FindSheets(wb, nG, KIND_D, ds), ds, groups
    For Each s In toDel
        delList = delList & "  " & s.Name & vbLf
    Next s
    If toDel.Count > 0 And Not scripted Then
        Application.ScreenUpdating = True
        If Ask(U("Bu sayfalar\u0131n grubu art\u0131k ayak tablosunda yok. Sayfalar ve \u00FCzerlerine yaz\u0131lm\u0131\u015F de\u011Ferler S\u0130L\u0130NECEK:") & vbLf & delList & vbLf & U("Devam edilsin mi?")) <> vbYes Then
            RunSync = "FAIL:" & U("\u0130\u015Flem durduruldu, hi\u00E7bir \u015Fey de\u011Fi\u015Fmedi.")
            Exit Function
        End If
        Application.ScreenUpdating = False
    End If

    ' no recalculation while sheets are copied and their formulas re-pointed (thousands of single-cell writes)
    t0 = Timer
    If Application.Calculation <> xlCalculationManual Then Application.Calculation = xlCalculationManual
    ' missing sheets first, so the workbook always keeps a visible sheet
    For Each k In groups
        kk = CLng(k)
        If gs(kk) Is Nothing Then Set gs(kk) = MakeGroupSheet(wb, KIND_G, kk): madeG = madeG & "|" & kk
        If hasY And ys(kk) Is Nothing Then Set ys(kk) = MakeGroupSheet(wb, KIND_Y, kk): madeY = madeY & "|" & kk
        If hasD And ds(kk) Is Nothing Then Set ds(kk) = MakeGroupSheet(wb, KIND_D, kk): madeD = madeD & "|" & kk
        If Not ys(kk) Is Nothing Then SetGroupNo ys(kk), KIND_Y, kk
        If Not ds(kk) Is Nothing Then SetGroupNo ds(kk), KIND_D, kk
    Next k
    If toDel.Count > 0 Then
        Application.DisplayAlerts = False
        For Each s In toDel
            s.Delete
        Next s
        Application.DisplayAlerts = True
        lg = lg & U("Silinen sayfa:") & vbLf & delList
    End If

    t = RenameSheets(wb, groups, gs, KIND_G) & RenameSheets(wb, groups, ys, KIND_Y) & RenameSheets(wb, groups, ds, KIND_D)
    lg = lg & MadeList(U("Yeni grup sayfas\u0131:"), madeG, gs) & MadeList(U("Yeni _Y sayfas\u0131:"), madeY, ys) & MadeList(U("Yeni veri sayfas\u0131:"), madeD, ds)
    If Len(t) > 0 Then lg = lg & U("Ad\u0131 de\u011Fi\u015Fen sayfa:") & vbLf & t

    OrderSheets wb, groups, gs, ys
    OrderData wb, groups, gs, ds
    ' formulas that still read a template are pointed at the group's own sheets
    For Each k In groups
        kk = CLng(k)
        n = RepointGroup(gs(kk), ys(kk), ds(kk))
        If n > 0 Then lg = lg & "  " & gs(kk).Name & U(": ") & n & U(" form\u00FCl grubun kendi sayfalar\u0131na ba\u011Fland\u0131") & vbLf
    Next k
    tm = U("sayfalar ") & Format$(Timer - t0, "0") & " s"
    If doPull Then
        t0 = Timer
        On Error Resume Next
        pl = PullForces(wb, oaPath)
        If Err.Number <> 0 Then pl = U("Temel tesirleri \u00E7ekilemedi: ") & Err.Description & vbLf
        Err.Clear
        On Error GoTo Fail
        lg = lg & pl
        tm = tm & U(", tesirler ") & Format$(Timer - t0, "0") & " s"
    End If
    t0 = Timer
    For Each k In groups
        kk = CLng(k)
        lg = lg & FitPilePages(gs(kk), ds(kk), tps(kk))
    Next k
    WriteChains wb, groups, gs
    If Application.Calculation <> xlCalculationAutomatic Then Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    For Each k In groups
        kk = CLng(k)
        n = ApplyGate(gs(kk))
        lg = lg & gs(kk).Name & ": " & GroupLabel(wb, kk, tps(kk)) & ", " & n & U(" sayfa") & vbLf
        If Not ys(kk) Is Nothing Then
            n = SetYPrint(ys(kk))
            If n > 0 Then lg = lg & ys(kk).Name & ": " & n & U(" sayfa") & vbLf
        End If
    Next k
    Application.Calculate
    tm = tm & U(", sayfa kap\u0131lar\u0131 ") & Format$(Timer - t0, "0") & " s"
    t0 = Timer
    kk = CLng(groups(groups.Count))
    Set lastSh = gs(kk)
    If Not ys(kk) Is Nothing Then Set lastSh = ys(kk)
    lg = lg & BuildGeo(wb, groups, gs, ys, ds, tps, lastSh)
    tm = tm & ", GEO " & Format$(Timer - t0, "0") & " s"
    t0 = Timer
    RunLists wb
    tm = tm & U(", listeler ") & Format$(Timer - t0, "0") & " s"
    Application.ScreenUpdating = True
    RunSync = groups.Count & U(" temel grubu.") & vbLf & lg & U("S\u00FCre: ") & tm & vbLf
    Exit Function
Fail:
    Application.DisplayAlerts = True
    If Application.Calculation <> xlCalculationAutomatic Then Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    RunSync = "FAIL:" & Err.Description & IIf(Len(lg) > 0, vbLf & lg, "")
End Function

Private Function MadeList(ByVal head As String, ByVal lst As String, sh() As Worksheet) As String
    Dim k As Variant
    If Len(lst) = 0 Then Exit Function
    MadeList = head & vbLf
    For Each k In Split(Mid$(lst, 2), "|")
        If Not sh(CLng(k)) Is Nothing Then MadeList = MadeList & "  " & sh(CLng(k)).Name & vbLf
    Next k
End Function

' sheets of groups that no longer exist (numbers above the table, or groups without piers) into toDel
Private Sub CollectRemoved(toDel As Collection, extra As Collection, sh() As Worksheet, groups As Collection)
    Dim s As Variant, n As Long
    For Each s In extra
        toDel.Add s
    Next s
    For n = LBound(sh) To UBound(sh)
        If Not sh(n) Is Nothing Then
            If Not InColl(groups, n) Then
                toDel.Add sh(n)
                Set sh(n) = Nothing
            End If
        End If
    Next n
End Sub

' existing sheets only: per-pile pages, numbers, gates, print areas, GEO, lists
Private Function RefreshTop() As String
    On Error GoTo Fail
    lg = ""
    Application.ScreenUpdating = False
    RefreshCore ThisWorkbook
    Application.ScreenUpdating = True
    RefreshTop = lg
    Exit Function
Fail:
    If Application.Calculation <> xlCalculationAutomatic Then Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    RefreshTop = "FAIL:" & Err.Description & IIf(Len(lg) > 0, vbLf & lg, "")
End Function

Private Sub RefreshCore(wb As Workbook)
    Dim nG As Long, gs() As Worksheet, ys() As Worksheet, ds() As Worksheet, tps() As String, groups As Collection
    Dim k As Long, n As Long, lastSh As Worksheet
    nG = NameRange(wb, N_GTAB).Rows.Count
    FindSheets wb, nG, KIND_G, gs
    FindSheets wb, nG, KIND_Y, ys
    FindSheets wb, nG, KIND_D, ds
    tps = TypesOf(wb, nG)
    Set groups = New Collection
    For k = 1 To nG
        If Not gs(k) Is Nothing Then groups.Add k
    Next k
    If groups.Count = 0 Then lg = lg & U("Grup sayfas\u0131 yok.") & vbLf: Exit Sub
    For k = 1 To nG
        If Not gs(k) Is Nothing Then lg = lg & FitPilePages(gs(k), ds(k), tps(k))
    Next k
    WriteChains wb, groups, gs
    If Application.Calculation <> xlCalculationAutomatic Then Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    For k = 1 To nG
        If Not gs(k) Is Nothing Then
            n = ApplyGate(gs(k))
            lg = lg & gs(k).Name & ": " & n & U(" sayfa") & vbLf
            If Not ys(k) Is Nothing Then SetYPrint ys(k)
        End If
    Next k
    Application.Calculate
    k = CLng(groups(groups.Count))
    Set lastSh = gs(k)
    If Not ys(k) Is Nothing Then Set lastSh = ys(k)
    lg = lg & BuildGeo(wb, groups, gs, ys, ds, tps, lastSh)
    RunLists wb
End Sub

' ==== groups ==========================================================================================================
' group numbers in use, ascending; tps(k) = Temel Tip of group k; stops when a group number or a Temel Tip cannot be read
' or the piers of one group differ in their foundation
Private Function ReadGroups(wb As Workbook, tps() As String) As Collection
    Dim t As Variant, g As Variant, nG As Long, i As Long, k As Long, v As Variant, ty As String, d As String
    Dim first() As Long, bad As String, noGrp As String, pier As String, gap As Boolean
    t = NameRange(wb, N_TABLE).Value
    g = NameRange(wb, N_GROUP).Value
    If UBound(g, 1) <> UBound(t, 1) Then Err.Raise vbObjectError + 20, , N_GROUP & U(" ile ") & N_TABLE & U(" ayn\u0131 sat\u0131r say\u0131s\u0131nda de\u011Fil.")
    If UBound(t, 2) < COL_PILE Then Err.Raise vbObjectError + 20, , N_TABLE & U(" en az ") & COL_PILE & U(" s\u00FCtun olmal\u0131 (temel s\u00FCtunlar\u0131 7-11).")
    nG = NameRange(wb, N_GTAB).Rows.Count
    ReDim first(1 To nG): ReDim tps(1 To nG)
    For i = 1 To UBound(t, 1)
        v = g(i, 1)
        pier = CellText(t(i, 1))
        ty = NormType(t(i, COL_TYPE))
        If IsError(v) Then Err.Raise vbObjectError + 21, , pier & U(": Grup No hata de\u011Feri.")
        If Len(Trim$(CStr(v))) > 0 Then
            k = 0
            If IsNumeric(v) Then
                If CDbl(v) = Int(CDbl(v)) And CDbl(v) >= 1 And CDbl(v) <= nG Then k = CLng(v)
            End If
            If k = 0 Then Err.Raise vbObjectError + 21, , pier & U(": Grup No 1 ile ") & nG & U(" aras\u0131nda bir tam say\u0131 olmal\u0131 ('") & CStr(v) & "')."
            If ty <> TYPE_PILED And ty <> TYPE_SHALLOW Then Err.Raise vbObjectError + 22, , pier & U(": temel tipi '") & CellText(t(i, COL_TYPE)) & U("' tan\u0131nm\u0131yor (KAZIKLI veya YUZEYSEL).")
            If first(k) = 0 Then
                first(k) = i
                tps(k) = ty
            Else
                d = FoundationDiff(t, first(k), i)
                If Len(d) > 0 Then bad = bad & "  Grup " & k & ": " & CellText(t(first(k), 1)) & " / " & pier & " (" & d & ")" & vbLf
            End If
        ElseIf (ty = TYPE_PILED Or ty = TYPE_SHALLOW) And Len(pier) > 0 And NormType(t(i, COL_SEC)) <> "KA" Then
            noGrp = noGrp & " " & pier
        End If
    Next i
    If Len(bad) > 0 Then Err.Raise vbObjectError + 23, , U("Bir grubun ayaklar\u0131n\u0131n temeli ayn\u0131 olmal\u0131 (Temel Tip, Kal\u0131nl\u0131k, En, Boy, Kaz\u0131k \u00C7ap\u0131). Farkl\u0131 olanlar:") & vbLf & bad & U("Grup numaralar\u0131n\u0131 veya ayak tablosunu d\u00FCzeltin.")
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

' Temel Tip of every group (from its first pier row), no checks: for the refresh
Private Function TypesOf(wb As Workbook, ByVal nG As Long) As String()
    Dim t As Variant, g As Variant, i As Long, k As Long, v As Variant, r() As String
    ReDim r(1 To nG)
    t = NameRange(wb, N_TABLE).Value
    g = NameRange(wb, N_GROUP).Value
    If UBound(t, 2) >= COL_TYPE Then
        For i = 1 To UBound(g, 1)
            If i > UBound(t, 1) Then Exit For
            v = g(i, 1)
            If Not IsError(v) Then
                If IsNumeric(v) And Not IsEmpty(v) Then
                    If CDbl(v) = Int(CDbl(v)) And CDbl(v) >= 1 And CDbl(v) <= nG Then
                        k = CLng(v)
                        If Len(r(k)) = 0 Then r(k) = NormType(t(i, COL_TYPE))
                    End If
                End If
            End If
        Next i
    End If
    TypesOf = r
End Function

' piers of group k in table order
Private Function PiersOf(wb As Workbook, ByVal k As Long) As Collection
    Dim t As Variant, g As Variant, i As Long, v As Variant
    Set PiersOf = New Collection
    t = NameRange(wb, N_TABLE).Value
    g = NameRange(wb, N_GROUP).Value
    For i = 1 To UBound(g, 1)
        If i > UBound(t, 1) Then Exit For
        v = g(i, 1)
        If Not IsError(v) Then
            If IsNumeric(v) And Not IsEmpty(v) Then
                If CDbl(v) = k And Len(CellText(t(i, 1))) > 0 Then PiersOf.Add CellText(t(i, 1))
            End If
        End If
    Next i
End Function

' what differs between the foundations of two pier rows of one group ("" = same)
Private Function FoundationDiff(t As Variant, ByVal a As Long, ByVal b As Long) As String
    Dim s As String
    If NormType(t(a, COL_TYPE)) <> NormType(t(b, COL_TYPE)) Then s = s & ", Temel Tip"
    If Not SameVal(t(a, COL_TYPE + 1), t(b, COL_TYPE + 1)) Then s = s & U(", Kal\u0131nl\u0131k")
    If Not SameVal(t(a, COL_TYPE + 2), t(b, COL_TYPE + 2)) Then s = s & ", En"
    If Not SameVal(t(a, COL_TYPE + 3), t(b, COL_TYPE + 3)) Then s = s & ", Boy"
    If Not SameVal(t(a, COL_PILE), t(b, COL_PILE)) Then s = s & U(", Kaz\u0131k \u00C7ap\u0131")
    If Len(s) > 0 Then FoundationDiff = Mid$(s, 3)
End Function

' equal numbers (relative 1e-9) or equal texts; empty and '-' are the same
Private Function SameVal(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim sa As String, sb As String
    If IsError(a) Or IsError(b) Then Exit Function
    If IsNumeric(a) And IsNumeric(b) And Not IsEmpty(a) And Not IsEmpty(b) Then
        SameVal = (Abs(CDbl(a) - CDbl(b)) <= 0.000000001 * (1 + Abs(CDbl(a))))
    Else
        sa = Trim$(CStr(a)): sb = Trim$(CStr(b))
        If sa = "-" Then sa = ""
        If sb = "-" Then sb = ""
        SameVal = (StrComp(sa, sb, vbTextCompare) = 0)
    End If
End Function

Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Then CellText = "#" Else CellText = Trim$(CStr(v))
End Function

' a type text in capitals without Turkish letters
Private Function NormType(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then Exit Function
    s = UCase$(Trim$(CStr(v)))
    s = Replace(s, ChrW(304), "I"): s = Replace(s, ChrW(305), "I")
    s = Replace(s, ChrW(214), "O"): s = Replace(s, ChrW(246), "O")
    s = Replace(s, ChrW(220), "U"): s = Replace(s, ChrW(252), "U")
    s = Replace(s, ChrW(286), "G"): s = Replace(s, ChrW(287), "G")
    s = Replace(s, ChrW(350), "S"): s = Replace(s, ChrW(351), "S")
    s = Replace(s, ChrW(199), "C"): s = Replace(s, ChrW(231), "C")
    NormType = s
End Function

Private Function GroupLabel(wb As Workbook, ByVal k As Long, ByVal ty As String) As String
    GroupLabel = CellText(NameRange(wb, N_GTAB).Cells(k, 2).Value) & " (" & ty & ")"
End Function

' ==== sheets of the groups ============================================================================================
' sh(1 To nG) = the sheet of each group of one kind (group sheet by M_TEMEL_ZINCIR, _Y sheet by M_TEMEL_Y_NO, data sheet
' by M_TESIR_NO); templates and baselines are skipped; returns the sheets whose number is above nG; stops when two sheets
' of one kind hold the same number
Private Function FindSheets(wb As Workbook, ByVal nG As Long, ByVal kind As Long, sh() As Worksheet) As Collection
    Dim ws As Worksheet, k As Long
    ReDim sh(1 To nG)
    Set FindSheets = New Collection
    For Each ws In wb.Worksheets
        If Not IsTemplateSheet(ws.Name) Then
            k = GroupOf(ws, kind)
            If k > nG Then
                FindSheets.Add ws
            ElseIf k > 0 Then
                If Not sh(k) Is Nothing Then Err.Raise vbObjectError + 32, , U("\u0130ki sayfa ayn\u0131 grup numaras\u0131n\u0131 ta\u015F\u0131yor (grup ") & k & "): '" & sh(k).Name & "' / '" & ws.Name & U("'. Birini silin.")
                Set sh(k) = ws
            End If
        End If
    Next ws
End Function

Private Function GroupOf(ws As Worksheet, ByVal kind As Long) As Long
    Dim r As Range, v As Variant
    Select Case kind
        Case KIND_G
            Set r = SheetRangeOpt(ws, M_CHAIN)
            If r Is Nothing Then Exit Function
            v = r.Cells(1, 1).Value
            If Not IsError(v) Then
                If IsNumeric(v) And Not IsEmpty(v) Then
                    If CDbl(v) = Int(CDbl(v)) And CDbl(v) >= 1 And CDbl(v) <= 100000 Then GroupOf = CLng(v)
                End If
            End If
            If GroupOf = 0 Then lg = lg & U("Uyar\u0131: '") & ws.Name & U("' sayfas\u0131n\u0131n zincir h\u00FCcresinde (") & M_CHAIN & U(") grup numaras\u0131 yok; grup sayfas\u0131 say\u0131lmad\u0131.") & vbLf
        Case KIND_Y
            GroupOf = NameNumber(ws, M_Y_NO)
        Case KIND_D
            GroupOf = NameNumber(ws, M_DATA_NO)
    End Select
    If GroupOf < 0 Then GroupOf = 0
End Function

Private Function IsTemplateSheet(ByVal nm As String) As Boolean
    Dim x As Variant
    For Each x In Array(TPL_G, TPL_Y, TPL_D, TPL_GEO_K, TPL_GEO_Y, BASE_G, BASE_G_R, BASE_Y, BASE_Y_R, BASE_D)
        If StrComp(nm, CStr(x), vbTextCompare) = 0 Then IsTemplateSheet = True: Exit Function
    Next x
End Function

Private Sub SetGroupNo(ws As Worksheet, ByVal kind As Long, ByVal k As Long)
    Select Case kind
        Case KIND_G
            With SheetRange(ws, M_CHAIN).Cells(1, 1)
                If CStr(.Value) <> CStr(k) Then .Value = k
            End With
        Case KIND_Y
            SetNameNumber ws, M_Y_NO, k
        Case KIND_D
            SetNameNumber ws, M_DATA_NO, k
    End Select
End Sub

' group sheet <group>_<pier runs>, _Y sheet <group>_<pier runs>_Y, data sheet MIDAS_<pier runs>; characters a sheet
' name cannot hold removed, at most MAX_NAME characters
Private Function SheetNameFor(wb As Workbook, ByVal k As Long, ByVal kind As Long) As String
    Dim v As Variant, s As String, c As Variant
    v = NameRange(wb, N_GTAB).Cells(k, 2).Value
    If Not IsError(v) Then s = CStr(v)
    For Each c In Array(" ", ChrW(160), vbTab, vbCr, vbLf, "\", "/", "?", "*", "[", "]", ":")
        s = Replace(s, c, "")
    Next c
    Select Case kind
        Case KIND_D
            If Len(s) = 0 Then s = CStr(k)
            SheetNameFor = CutName(DATA_PFX & s, "")
        Case KIND_Y
            If Len(s) = 0 Then s = Mid$(SHEET_SFX, 2)
            SheetNameFor = CutName(CStr(k) & "_" & s, Y_SFX)
        Case Else
            If Len(s) = 0 Then s = Mid$(SHEET_SFX, 2)
            SheetNameFor = CutName(CStr(k) & "_" & s, "")
    End Select
End Function

' base & sfx in at most MAX_NAME characters: base cut and ended with an ellipsis; no apostrophe at the end of base
Private Function CutName(ByVal base As String, ByVal sfx As String) As String
    If Len(base) + Len(sfx) > MAX_NAME Then base = Left$(base, MAX_NAME - Len(sfx) - 1) & ChrW(8230)
    Do While Right$(base, 1) = "'"
        base = Left$(base, Len(base) - 1)
    Loop
    CutName = base & sfx
End Function

' gives every sheet of one kind its name; two passes through free temporary names; one "old -> new" line per rename
Private Function RenameSheets(wb As Workbook, groups As Collection, sh() As Worksheet, ByVal kind As Long) As String
    Dim k As Variant, j As Variant, want() As String, old() As String, o As Object, own As Boolean
    Dim e As Long, d As String, u2 As String, kk As Long
    ReDim want(LBound(sh) To UBound(sh)): ReDim old(LBound(sh) To UBound(sh))
    For Each k In groups
        kk = CLng(k)
        If Not sh(kk) Is Nothing Then
            want(kk) = SheetNameFor(wb, kk, kind)
            Set o = Nothing
            On Error Resume Next
            Set o = wb.Sheets(want(kk))
            On Error GoTo 0
            If Not o Is Nothing Then
                own = False
                For Each j In groups
                    If Not sh(CLng(j)) Is Nothing Then
                        If StrComp(o.Name, sh(CLng(j)).Name, vbTextCompare) = 0 Then own = True: Exit For
                    End If
                Next j
                If Not own Then Err.Raise vbObjectError + 33, , "'" & o.Name & U("' ad\u0131nda bu t\u00FCrden olmayan bir sayfa var; grup ") & kk & U(" sayfas\u0131 bu ad\u0131 alam\u0131yor. O sayfan\u0131n ad\u0131n\u0131 de\u011Fi\u015Ftirin.")
            End If
            If kind = KIND_G Then
                u2 = UCase$(want(kk))
                If Right$(u2, 2) = "_Y" Or Right$(u2, 6) = "_YATAY" Or Right$(u2, 6) = "_KILIF" Then
                    lg = lg & U("Uyar\u0131: '") & want(kk) & U("' ad\u0131 _Y, _YATAY veya _KILIF ile bitiyor; i\u00E7indekiler ve listeler bu sayfay\u0131 almaz.") & vbLf
                End If
            End If
        End If
    Next k
    For Each k In groups
        kk = CLng(k)
        If Not sh(kk) Is Nothing Then
            If StrComp(sh(kk).Name, want(kk), vbBinaryCompare) <> 0 Then
                old(kk) = sh(kk).Name
                sh(kk).Name = FreeName(wb, TMP_BASE)
            End If
        End If
    Next k
    For Each k In groups
        kk = CLng(k)
        If Len(old(kk)) > 0 Then
            On Error Resume Next
            sh(kk).Name = want(kk)
            e = Err.Number: d = Err.Description
            On Error GoTo 0
            If e <> 0 Then Err.Raise vbObjectError + 34, , U("Sayfa ad\u0131 verilemedi (") & want(kk) & "): " & d
            If Left$(old(kk), Len(TMP_BASE)) <> TMP_BASE Then RenameSheets = RenameSheets & "  " & old(kk) & " -> " & want(kk) & vbLf
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

' copy of a template sheet after the last visible sheet, named newName, visible. The new sheet is the one whose name
' was not there before the copy: measured 2026-11-26, a copy placed after a very hidden last sheet lands right after
' its template instead, so the sheet at the expected index is another one.
Private Function MakeCopy(wb As Workbook, ByVal tplName As String, ByVal newName As String) As Worksheet
    Dim tpl As Worksheet, ws As Worksheet, vis As Long, e As Long, d As String, seen As Object, w As Worksheet, anchor As Worksheet, nNew As Long
    If Not SheetExists(wb, tplName) Then Err.Raise vbObjectError + 25, , U("\u015Eablon sayfa yok: ") & tplName
    Set tpl = wb.Worksheets(tplName)
    vis = tpl.Visible
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For Each w In wb.Worksheets
        seen(w.Name) = True
        If w.Visible = xlSheetVisible Then Set anchor = w
    Next w
    On Error Resume Next
    tpl.Visible = xlSheetVisible
    Application.DisplayAlerts = False
    If anchor Is Nothing Then tpl.Copy After:=tpl Else tpl.Copy After:=anchor
    e = Err.Number: d = Err.Description
    For Each w In wb.Worksheets
        If Not seen.Exists(w.Name) Then Set ws = w: nNew = nNew + 1
    Next w
    If e = 0 And nNew = 1 Then
        ws.Name = newName
        ws.Visible = xlSheetVisible
        e = Err.Number: d = Err.Description
    ElseIf e = 0 Then
        e = 1: d = "copy (" & nNew & U(" yeni sayfa)")
    End If
    Err.Clear
    tpl.Visible = vis
    Application.DisplayAlerts = True
    On Error GoTo 0
    If e <> 0 Then Err.Raise vbObjectError + 26, , U("Sayfa a\u00E7\u0131lamad\u0131 (") & tplName & " -> " & newName & "): " & d
    Set MakeCopy = ws
End Function

' copy of the kind's template with a temporary name (RenameSheets gives the real one), group number set
Private Function MakeGroupSheet(wb As Workbook, ByVal kind As Long, ByVal k As Long) As Worksheet
    Dim tn As String, ws As Worksheet
    Select Case kind
        Case KIND_Y: tn = TPL_Y
        Case KIND_D: tn = TPL_D
        Case Else: tn = TPL_G
    End Select
    Set ws = MakeCopy(wb, tn, FreeName(wb, TMP_BASE))
    SetGroupNo ws, kind, k
    Set MakeGroupSheet = ws
End Function

' after AYRAC: group sheet, its _Y sheet, next group ...; GEO after the last of them
Private Sub OrderSheets(wb As Workbook, groups As Collection, gs() As Worksheet, ys() As Worksheet)
    Dim k As Variant, before As Worksheet
    On Error Resume Next
    Set before = wb.Worksheets(AFTER_SHEET)
    On Error GoTo 0
    For Each k In groups
        MoveAfter gs(CLng(k)), before
        Set before = gs(CLng(k))
        If Not ys(CLng(k)) Is Nothing Then
            MoveAfter ys(CLng(k)), before
            Set before = ys(CLng(k))
        End If
    Next k
    If SheetExists(wb, GEO_SHEET) Then MoveAfter wb.Worksheets(GEO_SHEET), before
End Sub

' ws right after before in the visible tab order. Measured 2026-11-26: Move After places a sheet behind the hidden sheets
' that follow the anchor, so an index test would move every sheet on every run; hidden sheets between are allowed.
Private Sub MoveAfter(ws As Worksheet, before As Worksheet)
    Dim i As Long, sh As Object, wb As Workbook
    If before Is Nothing Then Exit Sub
    Set wb = before.Parent
    For i = before.Index + 1 To wb.Sheets.Count
        Set sh = wb.Sheets(i)
        If sh Is ws Then Exit Sub
        If sh.Visible = xlSheetVisible Then Exit For
    Next i
    ws.Move After:=before
End Sub

' ws right before anchor in the visible tab order (hidden sheets between are allowed)
Private Sub MoveBefore(ws As Worksheet, anchor As Worksheet)
    Dim i As Long, sh As Object, wb As Workbook
    Set wb = anchor.Parent
    For i = anchor.Index - 1 To 1 Step -1
        Set sh = wb.Sheets(i)
        If sh Is ws Then Exit Sub
        If sh.Visible = xlSheetVisible Then Exit For
    Next i
    ws.Move Before:=anchor
End Sub

' data sheets in group order right before INPUT (after the last group sheet when there is no INPUT)
Private Sub OrderData(wb As Workbook, groups As Collection, gs() As Worksheet, ds() As Worksheet)
    Dim anchor As Worksheet, i As Long, k As Long, lst As Collection, x As Variant
    Set lst = New Collection
    For Each x In groups
        If Not ds(CLng(x)) Is Nothing Then lst.Add ds(CLng(x))
    Next x
    If lst.Count = 0 Then Exit Sub
    On Error Resume Next
    Set anchor = wb.Worksheets(DATA_BEFORE)
    On Error GoTo 0
    If anchor Is Nothing Then
        k = CLng(groups(groups.Count))
        Set anchor = gs(k)
        For i = 1 To lst.Count
            MoveAfter lst(i), anchor
            Set anchor = lst(i)
        Next i
    Else
        For i = lst.Count To 1 Step -1
            MoveBefore lst(i), anchor
            Set anchor = lst(i)
        Next i
    End If
End Sub

' ==== references to the templates =====================================================================================
' the group's three sheets: references to SABLON_TEMEL / SABLON_TEMEL_Y / SABLON_MIDAS -> the group's own sheets
Private Function RepointGroup(gs As Worksheet, ys As Worksheet, ds As Worksheet) As Long
    If Not gs Is Nothing Then RepointGroup = RepointRange(gs.UsedRange, gs, ys, ds) + RepointCharts(gs, gs, ys, ds)
    If Not ys Is Nothing Then RepointGroup = RepointGroup + RepointRange(ys.UsedRange, gs, ys, ds) + RepointCharts(ys, gs, ys, ds)
    If Not ds Is Nothing Then RepointGroup = RepointGroup + RepointRange(ds.UsedRange, gs, ys, ds) + RepointCharts(ds, gs, ys, ds)
End Function

' chart series of a sheet that read a template (cell formulas do not reach them) -> the group's own sheets; returns the
' number of series written
Private Function RepointCharts(ws As Worksheet, gs As Worksheet, ys As Worksheet, ds As Worksheet) As Long
    Dim co As ChartObject, i As Long, f As String, n As String, e As Long
    For Each co In ws.ChartObjects
        For i = 1 To co.Chart.SeriesCollection.Count
            f = ""
            On Error Resume Next
            f = co.Chart.SeriesCollection(i).Formula
            On Error GoTo 0
            If InStr(1, f, TPL_PFX, vbTextCompare) > 0 Then
                n = ToGroup(f, gs, ys, ds)
                If n <> f Then
                    On Error Resume Next
                    co.Chart.SeriesCollection(i).Formula = n
                    e = Err.Number
                    Err.Clear
                    On Error GoTo 0
                    If e = 0 Then
                        RepointCharts = RepointCharts + 1
                    Else
                        lg = lg & U("Uyar\u0131: ") & ws.Name & " / " & co.Name & U(": grafik serisi \u015Fablona ba\u011Fl\u0131 kald\u0131: ") & f & vbLf
                    End If
                End If
            End If
        Next i
    Next co
End Function

' every formula of rg that names a template is written in the group's terms (quoted and plain sheet names, ReplaceSheetRef);
' geoName: references to the GEO block templates go to that sheet; returns the number of cells written
Private Function RepointRange(rg As Range, gs As Worksheet, ys As Worksheet, ds As Worksheet, Optional ByVal geoName As String = "") As Long
    Dim f As Variant, r As Long, c As Long, s As String, n As String, cl As Range, hit As Range
    On Error Resume Next
    Set hit = rg.Find(What:=TPL_PFX, LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
    On Error GoTo 0
    If hit Is Nothing Then Exit Function
    ' cells that name a template now (the count returned)
    f = Arr2(rg, True)
    For r = 1 To UBound(f, 1)
        For c = 1 To UBound(f, 2)
            If VarType(f(r, c)) = vbString Then
                If Left$(f(r, c), 1) = "=" Then
                    If InStr(1, f(r, c), TPL_PFX, vbTextCompare) > 0 Then RepointRange = RepointRange + 1
                End If
            End If
        Next c
    Next r
    ' Excel's own replace first, one call per template name and form (quoted and plain): thousands of single-cell writes
    ' took minutes on a group sheet; the template names are plain and no other sheet name ends with one of them
    BulkReplace rg, TPL_Y, ys
    BulkReplace rg, TPL_G, gs
    BulkReplace rg, TPL_D, ds
    If Len(geoName) > 0 Then
        BulkReplace rg, TPL_GEO_K, Nothing, geoName
        BulkReplace rg, TPL_GEO_Y, Nothing, geoName
    End If
    ' what is left (array formulas, anything the replace did not reach) one cell at a time through ReplaceSheetRef
    Set hit = Nothing
    On Error Resume Next
    Set hit = rg.Find(What:=TPL_PFX, LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
    On Error GoTo 0
    If hit Is Nothing Then Exit Function
    f = Arr2(rg, True)
    For r = 1 To UBound(f, 1)
        For c = 1 To UBound(f, 2)
            If VarType(f(r, c)) = vbString Then
                s = f(r, c)
                If Left$(s, 1) = "=" Then
                    If InStr(1, s, TPL_PFX, vbTextCompare) > 0 Then
                        n = ToGroup(s, gs, ys, ds, geoName)
                        If n <> s Then
                            Set cl = rg.Cells(r, c)
                            If cl.HasArray Then
                                lg = lg & U("Uyar\u0131: ") & rg.Worksheet.Name & "!" & cl.Address(False, False) & U(" dizi form\u00FCl\u00FC; \u015Fablona ba\u011Fl\u0131 kald\u0131.") & vbLf
                            Else
                                cl.Formula = n
                            End If
                        End If
                    End If
                End If
            End If
        Next c
    Next r
End Function

' every 'fromName'! and fromName! in the formulas of rg -> the sheet target (or the sheet name toName), in Excel's form
Private Sub BulkReplace(rg As Range, ByVal fromName As String, target As Worksheet, Optional ByVal toName As String = "")
    Dim tgt As String, da As Boolean
    If Not target Is Nothing Then toName = target.Name
    If Len(toName) = 0 Then Exit Sub
    If NeedsQuotes(toName) Then tgt = "'" & Replace(toName, "'", "''") & "'!" Else tgt = toName & "!"
    da = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    rg.Replace What:="'" & fromName & "'!", Replacement:=tgt, LookAt:=xlPart, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    rg.Replace What:=fromName & "!", Replacement:=tgt, LookAt:=xlPart, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Err.Clear
    On Error GoTo 0
    Application.DisplayAlerts = da
End Sub

' formulas (asFormula) or values of a range as a 2-D array, also for one cell
Private Function Arr2(rg As Range, ByVal asFormula As Boolean) As Variant
    Dim a As Variant
    If rg.Cells.CountLarge = 1 Then
        ReDim a(1 To 1, 1 To 1)
        If asFormula Then a(1, 1) = rg.Formula Else a(1, 1) = rg.Value
        Arr2 = a
    ElseIf asFormula Then
        Arr2 = rg.Formula
    Else
        Arr2 = rg.Value
    End If
End Function

' a template formula in the group's terms
Private Function ToGroup(ByVal f As String, gs As Worksheet, ys As Worksheet, ds As Worksheet, Optional ByVal geoName As String = "") As String
    If Left$(f, 1) <> "=" Then ToGroup = f: Exit Function
    If Not ys Is Nothing Then f = ReplaceSheetRef(f, TPL_Y, ys.Name)
    If Not gs Is Nothing Then f = ReplaceSheetRef(f, TPL_G, gs.Name)
    If Not ds Is Nothing Then f = ReplaceSheetRef(f, TPL_D, ds.Name)
    If Len(geoName) > 0 Then
        f = ReplaceSheetRef(f, TPL_GEO_K, geoName)
        f = ReplaceSheetRef(f, TPL_GEO_Y, geoName)
    End If
    ToGroup = f
End Function

' a group's formula in template terms
Private Function ToTemplate(ByVal f As String, gs As Worksheet, ys As Worksheet, ds As Worksheet) As String
    If Left$(f, 1) <> "=" Then ToTemplate = f: Exit Function
    If Not ys Is Nothing Then f = ReplaceSheetRef(f, ys.Name, TPL_Y)
    If Not gs Is Nothing Then f = ReplaceSheetRef(f, gs.Name, TPL_G)
    If Not ds Is Nothing Then f = ReplaceSheetRef(f, ds.Name, TPL_D)
    ToTemplate = f
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

' Excel writes a sheet name without quotes only when it holds ASCII letters, digits, _ and . , does not start with a
' digit and does not read as a cell address (letters then digits)
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

' ==== per-pile table pages ============================================================================================
' copies of the pattern page (M_TEMEL_KAZIK_SAYFA) below it: one for every PILES_PER_PAGE piles over the first page (pile
' count M_KAZIK_ADET of the data sheet), none for a YUZEYSEL group; returns a log line when the number changed. Without
' the pattern name, or without a pile count, nothing is done. forceN >= 0: exactly that many copies (the template apply
' takes them out with 0, then fits again).
Private Function FitPilePages(ws As Worksheet, ds As Worksheet, ByVal ty As String, Optional ByVal forceN As Long = -1) As String
    Dim pat As Range, p0 As Long, pr As Long, ex0 As Long, nOld As Long, nNew As Long, nP As Long
    Dim r As Long, i As Long, wOff As Long, wc As Long, top As Long, prevW As Long, co As Range, v As Variant, cPat As Variant
    Dim delta As Long, cm As Long, eNo As Long, eTx As String, nextOff As Long, gc As Long, gl As String
    Set pat = SheetRangeOpt(ws, M_PILE_PAGE)
    If pat Is Nothing Then Exit Function
    nOld = NameNumber(ws, M_PILE_EXTRA)
    If forceN >= 0 Then
        nNew = forceN
    ElseIf ty <> TYPE_PILED Then
        nNew = 0
    Else
        nP = PileCount(ds)
        If nP < 0 Then Exit Function
        If nP > PILES_PER_PAGE Then nNew = -Int(-(nP - PILES_PER_PAGE) / PILES_PER_PAGE)
    End If
    If nNew = nOld Then Exit Function
    p0 = pat.Row: pr = pat.Rows.Count: ex0 = p0 + pr
    wc = PageCol(ws)
    gc = SheetRange(ws, M_GATE).Column
    gl = Split(ws.Cells(1, gc).Address(True, False), "$")(0)
    wOff = -1
    For r = p0 To ex0 - 1
        If IsPageFormula(ws.Cells(r, wc), gl) Then wOff = r - p0
    Next r
    If wOff < 0 Then Err.Raise vbObjectError + 37, , ws.Name & U(": kaz\u0131k tablosu \u00F6rnek sayfas\u0131nda (") & M_PILE_PAGE & U(") sayfa numaras\u0131 yok.")
    Set co = SheetRange(ws, M_PAGE)
    cPat = ws.Cells(p0, co.Column).Value
    ' the page after the copies: its page-number cell, found before rows move (after a delete its formula would read
    ' #REF! and no longer look like a page number)
    nextOff = -1
    For r = ex0 + pr * nOld To ex0 + pr * nOld + 2 * pr
        If IsPageFormula(ws.Cells(r, wc), gl) Then nextOff = r - (ex0 + pr * nOld): Exit For
    Next r
    cm = Application.Calculation
    If cm <> xlCalculationManual Then Application.Calculation = xlCalculationManual
    On Error GoTo Bad
    If nNew > nOld Then
        For i = nOld + 1 To nNew
            pat.EntireRow.Copy
            ws.Rows(ex0 + pr * (i - 1)).Insert Shift:=xlDown
        Next i
        Application.CutCopyMode = False
        ' new copies: relative row references to the data sheet continue the table
        If Not ds Is Nothing Then
            For i = nOld + 1 To nNew
                ShiftCopy ws, ex0 + pr * (i - 1), pr, ds.Name, i * (PILES_PER_PAGE - pr)
            Next i
        End If
    Else
        ws.Rows((ex0 + pr * nNew) & ":" & (ex0 + pr * nOld - 1)).Delete
    End If
    ' page numbers: each copy follows the page before it, the page after the copies follows the last one
    For i = 1 To nNew
        top = ex0 + pr * (i - 1)
        prevW = IIf(i = 1, p0 + wOff, top - pr + wOff)
        ws.Cells(top + wOff, wc).Formula = "=" & ws.Cells(prevW, wc).Address & "+" & ws.Cells(top + wOff, gc).Address(False, True)
    Next i
    prevW = IIf(nNew = 0, p0 + wOff, ex0 + pr * (nNew - 1) + wOff)
    If nextOff >= 0 Then
        r = ex0 + pr * nNew + nextOff
        ws.Cells(r, wc).Formula = "=" & ws.Cells(prevW, wc).Address & "+" & ws.Cells(r, gc).Address(False, True)
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
    SetNameNumber ws, M_PILE_EXTRA, nNew
    If Application.Calculation <> cm Then Application.Calculation = cm
    FitPilePages = ws.Name & U(": kaz\u0131k tablosu ") & IIf(forceN >= 0, "", nP & U(" kaz\u0131k, ")) & U("ek sayfa ") & nOld & " -> " & nNew & vbLf
    Exit Function
Bad:
    eNo = Err.Number: eTx = Err.Description
    Application.CutCopyMode = False
    If Application.Calculation <> cm Then Application.Calculation = cm
    Err.Raise eNo, , ws.Name & U(" (kaz\u0131k tablosu): ") & eTx
End Function

' the column of the page numbers: the one of M_ILK_SAYFA, else W
Private Function PageCol(ws As Worksheet) As Long
    Dim r As Range
    Set r = SheetRangeOpt(ws, M_FIRST)
    If r Is Nothing Then PageCol = 23 Else PageCol = r.Column
End Function

' a page-number formula: previous page + the gate of its row (8'7 form =$W$105+$CG151); gl = column letters of the gate
Private Function IsPageFormula(c As Range, ByVal gl As String) As Boolean
    Dim f As String
    If Not c.HasFormula Then Exit Function
    f = UCase$(c.Formula)
    IsPageFormula = (f Like "=*+$" & gl & "#*" Or f Like "=*+" & gl & "#*")
End Function

' one copy (rows top .. top+pr-1): relative row references to the data sheet moved by d rows
Private Sub ShiftCopy(ws As Worksheet, ByVal top As Long, ByVal pr As Long, ByVal dsName As String, ByVal d As Long)
    Dim rg As Range, f As Variant, r As Long, c As Long, s As String, n As String, ur As Range
    If d = 0 Then Exit Sub
    Set ur = ws.UsedRange
    Set rg = ws.Range(ws.Cells(top, ur.Column), ws.Cells(top + pr - 1, ur.Column + ur.Columns.Count - 1))
    If rg.Cells.CountLarge < 2 Then Exit Sub
    f = rg.FormulaR1C1
    For r = 1 To UBound(f, 1)
        For c = 1 To UBound(f, 2)
            If VarType(f(r, c)) = vbString Then
                s = f(r, c)
                If Left$(s, 1) = "=" Then
                    If InStr(1, s, dsName, vbTextCompare) > 0 Then
                        n = ShiftRows(s, dsName, d)
                        If n <> s Then
                            If Not rg.Cells(r, c).HasArray Then rg.Cells(r, c).FormulaR1C1 = n
                        End If
                    End If
                End If
            End If
        Next c
    Next r
End Sub

' an R1C1 formula with every relative row reference to sheet sn (R[n], and R alone = this row) moved by d rows;
' absolute rows (INDEX ranges of the per-pile table) stay
Private Function ShiftRows(ByVal f As String, ByVal sn As String, ByVal d As Long) As String
    Dim q(1) As String, k As Long, p As Long, j As Long, pre As String
    q(0) = "'" & Replace(sn, "'", "''") & "'!"
    If NeedsQuotes(sn) Then q(1) = "" Else q(1) = sn & "!"
    For k = 0 To 1
        If Len(q(k)) > 0 Then
            p = InStr(1, f, q(k), vbTextCompare)
            Do While p > 0
                pre = ""
                If p > 1 Then pre = Mid$(f, p - 1, 1)
                j = p + Len(q(k))
                If Not (k = 1 And pre Like "[A-Za-z0-9_.'!]") Then
                    j = ShiftRef(f, j, d)
                    If Mid$(f, j, 1) = ":" Then j = ShiftRef(f, j + 1, d)
                End If
                If j <= p Then j = p + 1
                p = InStr(j, f, q(k), vbTextCompare)
            Loop
        End If
    Next k
    ShiftRows = f
End Function

' at position j of f a reference R..C..: a relative row moved by d; returns the position after the reference
Private Function ShiftRef(ByRef f As String, ByVal j As Long, ByVal d As Long) As Long
    Dim e As Long, n As Long, s As String, nx As String, nx2 As String
    ShiftRef = j
    If UCase$(Mid$(f, j, 1)) <> "R" Then Exit Function
    nx = Mid$(f, j + 1, 1)
    If nx = "[" Then
        e = InStr(j + 2, f, "]")
        If e = 0 Then Exit Function
        n = CLng(Val(Mid$(f, j + 2, e - j - 2)))
        If n + d = 0 Then s = "R" Else s = "R[" & (n + d) & "]"
        f = Left$(f, j - 1) & s & Mid$(f, e + 1)
        j = j + Len(s)
    ElseIf nx Like "#" Then
        j = j + 1
        Do While Mid$(f, j, 1) Like "#"
            j = j + 1
        Loop
    ElseIf UCase$(nx) = "C" Then
        nx2 = Mid$(f, j + 2, 1)
        If Not (nx2 = "[" Or nx2 Like "#" Or Not (nx2 Like "[A-Za-z_.]")) Then Exit Function
        s = "R[" & d & "]"
        f = Left$(f, j - 1) & s & Mid$(f, j + 1)
        j = j + Len(s)
    Else
        Exit Function
    End If
    If UCase$(Mid$(f, j, 1)) = "C" Then
        j = j + 1
        If Mid$(f, j, 1) = "[" Then
            e = InStr(j, f, "]")
            If e > 0 Then j = e + 1
        Else
            Do While Mid$(f, j, 1) Like "#"
                j = j + 1
            Loop
        End If
    End If
    ShiftRef = j
End Function

' pile count of a data sheet (M_KAZIK_ADET); -1 when it has none
Private Function PileCount(ds As Worksheet) As Long
    Dim r As Range, v As Variant
    PileCount = -1
    If ds Is Nothing Then Exit Function
    Set r = SheetRangeOpt(ds, M_PILES)
    If r Is Nothing Then Exit Function
    v = r.Cells(1, 1).Value
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then PileCount = 0: Exit Function
    If IsNumeric(v) Then
        If CDbl(v) >= 0 And CDbl(v) <= MAX_PILES Then PileCount = CLng(CDbl(v))
    End If
End Function

' rows the per-pile copies add below the pattern page now, and the pattern page's last row
Private Sub PileShift(ws As Worksheet, ByRef nShift As Long, ByRef lastRow As Long)
    Dim pat As Range
    nShift = 0: lastRow = 0
    Set pat = SheetRangeOpt(ws, M_PILE_PAGE)
    If pat Is Nothing Then Exit Sub
    lastRow = pat.Row + pat.Rows.Count - 1
    nShift = NameNumber(ws, M_PILE_EXTRA) * pat.Rows.Count
End Sub

' ==== numbers, gates, print areas =====================================================================================
Private Sub WriteChains(wb As Workbook, groups As Collection, gs() As Worksheet)
    Dim k As Variant, ws As Worksheet, c As Range, lastC As Range, j As Long, f As String, r As Range
    For Each k In groups
        Set ws = gs(CLng(k))
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
    If lastC Is Nothing Then Exit Sub
    Set r = NameRange(wb, N_LAST)
    For j = 1 To 3
        f = "='" & Replace(lastC.Worksheet.Name, "'", "''") & "'!" & lastC.Cells(2, j + 1).Address
        If r.Cells(j).Formula <> f Then r.Cells(j).Formula = f
    Next j
End Sub

' gated: hides the rows whose gate is 0, print area = the pages whose first row is open, pictures follow their row;
' not gated (_Y sheets): print area = every page of the page index; returns the page count
Private Function ApplyGate(ws As Worksheet, Optional ByVal gated As Boolean = True) As Long
    Dim rg As Range, rp As Range, rb As Range, g As Variant, p As Variant
    Dim n As Long, r0 As Long, i As Long, j As Long, o As Boolean, pg As Long, cur As Long, a As Long, addr As String
    Dim part As Range, pa As Range
    Set rp = SheetRange(ws, M_PAGE)
    Set rb = SheetRange(ws, M_PRINT)
    p = Arr2(rp, False)
    n = UBound(p, 1): r0 = rp.Row
    If gated Then
        Set rg = SheetRange(ws, M_GATE)
        g = Arr2(rg, False)
        If UBound(g, 1) <> n Or rg.Row <> r0 Then Err.Raise vbObjectError + 27, , ws.Name & ": " & M_GATE & " / " & M_PAGE & U(" ayn\u0131 sat\u0131rlarda de\u011Fil.")
        For i = 1 To n
            If IsError(g(i, 1)) Then Err.Raise vbObjectError + 28, , ws.Name & U(": sayfa kap\u0131s\u0131 hesaplanam\u0131yor (bu grubun temel tipi okunam\u0131yor).")
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
    Else
        ReDim g(1 To n, 1 To 1)
        For i = 1 To n
            g(i, 1) = 1
        Next i
    End If
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
    If gated Then GateShapes ws, g, r0
End Function

' a _Y sheet with a page index and printed columns: print area = all its pages; else its template's print area stays (0)
Private Function SetYPrint(ys As Worksheet) As Long
    If SheetRangeOpt(ys, M_PAGE) Is Nothing Or SheetRangeOpt(ys, M_PRINT) Is Nothing Then Exit Function
    SetYPrint = ApplyGate(ys, False)
End Function

' print area of many pieces (one per page): PageSetup and a text formula take at most 255 characters, so a longer one is
' given to Excel's own print-area name of the sheet as the union Range itself (its local name differs in a Turkish Excel)
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

' drawing objects follow the gate of their top-left row; comments and controls are left alone, objects outside the gated rows too
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
        If Len(CStr(r)) > 0 Then lg = lg & CStr(r) & vbLf
    End If
    On Error GoTo 0
End Sub

' ==== GEO =============================================================================================================
' GEO rebuilt: one block per group, stacked in group order (built from the last group upwards: inserting rows above a
' block moves its absolute references with it), block = rows of SABLON_GEO_KAZIK / SABLON_GEO_YUZEYSEL by the group's type,
' references pointed at the group's sheets, per-pile rows (M_KAZIK_GEO of the block template) grown with the pile count
Private Function BuildGeo(wb As Workbook, groups As Collection, gs() As Worksheet, ys() As Worksheet, ds() As Worksheet, tps() As String, afterSh As Worksheet) As String
    Dim geo As Worksheet, i As Long, k As Long, tpl As Worksheet, nR As Long, nC As Long, maxC As Long, made As Boolean
    Dim gr As Range, nP As Long, have As Long, d As Long, lastR As Long, c1 As Long, c2 As Long, tot As Long
    Dim cm As Long, s As String, cnt As Long, h() As Long, top As Long, eNo As Long, eTx As String, addr As String, cur As String
    Dim pr1() As Long, pr2() As Long, pc1() As Long, pc2() As Long, tpa As Range, part As Range, pa As Range
    If Not SheetExists(wb, TPL_GEO_K) And Not SheetExists(wb, TPL_GEO_Y) Then
        BuildGeo = U("Not: GEO blok \u015Fablonlar\u0131 (") & TPL_GEO_K & " / " & TPL_GEO_Y & U(") yok; GEO sayfas\u0131 kurulmad\u0131.") & vbLf
        Exit Function
    End If
    If groups.Count = 0 Then Exit Function
    cm = Application.Calculation
    Set geo = GeoSheet(wb, afterSh, made)
    If Application.Calculation <> xlCalculationManual Then Application.Calculation = xlCalculationManual
    On Error GoTo Bad
    ClearGeo geo
    ReDim h(1 To groups.Count): ReDim pr1(1 To groups.Count): ReDim pr2(1 To groups.Count): ReDim pc1(1 To groups.Count): ReDim pc2(1 To groups.Count)
    For i = groups.Count To 1 Step -1
        k = CLng(groups(i))
        Set tpl = Nothing
        If tps(k) = TYPE_PILED And SheetExists(wb, TPL_GEO_K) Then Set tpl = wb.Worksheets(TPL_GEO_K)
        If tps(k) = TYPE_SHALLOW And SheetExists(wb, TPL_GEO_Y) Then Set tpl = wb.Worksheets(TPL_GEO_Y)
        If tpl Is Nothing Then
            s = s & "  " & gs(k).Name & U(": GEO blok \u015Fablonu yok (") & tps(k) & ")" & vbLf
        Else
            LastCell tpl, nR, nC
            If nC > maxC Then maxC = nC
            ' the block prints what its template prints (first area of its print area), else all its used cells
            Set tpa = Nothing
            On Error Resume Next
            Set tpa = tpl.Range(Split(Replace(tpl.PageSetup.PrintArea, ";", ","), ",")(0))
            On Error GoTo Bad
            If tpa Is Nothing Then
                pr1(i) = 1: pr2(i) = nR: pc1(i) = 1: pc2(i) = nC
            Else
                pr1(i) = tpa.Row: pr2(i) = tpa.Row + tpa.Rows.Count - 1: pc1(i) = tpa.Column: pc2(i) = tpa.Column + tpa.Columns.Count - 1
            End If
            If cnt > 0 Then geo.Rows("1:" & nR).Insert Shift:=xlShiftDown
            tpl.Rows("1:" & nR).Copy Destination:=geo.Cells(1, 1)
            RepointRange geo.Range(geo.Cells(1, 1), geo.Cells(nR, nC)), gs(k), ys(k), ds(k), geo.Name
            d = 0
            Set gr = SheetRangeOpt(tpl, M_GEO_ROWS)
            If Not gr Is Nothing And tps(k) = TYPE_PILED Then
                nP = PileCount(ds(k))
                have = gr.Rows.Count
                If nP > have And have >= 2 And gr.Row + have - 1 <= nR Then
                    d = nP - have
                    lastR = gr.Row + have - 1
                    c1 = gr.Column: c2 = gr.Column + gr.Columns.Count - 1
                    geo.Rows(lastR & ":" & (lastR + d - 1)).Insert Shift:=xlShiftDown
                    geo.Range(geo.Cells(lastR - 1, c1), geo.Cells(lastR - 1, c2)).Copy Destination:=geo.Range(geo.Cells(lastR, c1), geo.Cells(lastR + d - 1, c2))
                End If
            End If
            If ds(k) Is Nothing Then s = s & "  " & gs(k).Name & U(": veri sayfas\u0131 yok; GEO blo\u011Fu \u015Fablonu okuyor.") & vbLf
            h(i) = nR + d
            If d > 0 And lastR <= pr2(i) Then pr2(i) = pr2(i) + d
            cnt = cnt + 1
        End If
    Next i
    ' a page break before every block but the first, print area = the printed part of every block
    top = 1
    For i = 1 To groups.Count
        If h(i) > 0 Then
            If top > 1 Then
                On Error Resume Next
                geo.HPageBreaks.Add Before:=geo.Cells(top, 1)
                Err.Clear
                On Error GoTo Bad
            End If
            Set part = geo.Range(geo.Cells(top + pr1(i) - 1, pc1(i)), geo.Cells(top + pr2(i) - 1, pc2(i)))
            addr = addr & "," & part.Address
            If pa Is Nothing Then Set pa = part Else Set pa = Union(pa, part)
            top = top + h(i)
        End If
    Next i
    tot = top - 1
    If tot > 0 And Not pa Is Nothing Then
        addr = Mid$(addr, 2)
        cur = Replace(Replace(Replace(geo.PageSetup.PrintArea, "'" & geo.Name & "'!", ""), geo.Name & "!", ""), ";", ",")
        If cur <> addr Then SetPrintAreaLong geo, addr, pa
        If cnt > 1 Then
            On Error Resume Next
            With geo.PageSetup
                If VarType(.Zoom) = vbBoolean Then
                    If .FitToPagesTall <> False Then .FitToPagesTall = False
                End If
            End With
            Err.Clear
            On Error GoTo Bad
        End If
    End If
    If Application.Calculation <> cm Then Application.Calculation = cm
    BuildGeo = geo.Name & ": " & cnt & U(" blok, ") & tot & U(" sat\u0131r") & IIf(made, U(" (yeni sayfa)"), "") & "." & vbLf & s
    Exit Function
Bad:
    eNo = Err.Number: eTx = Err.Description
    Application.CutCopyMode = False
    If Application.Calculation <> cm Then Application.Calculation = cm
    Err.Raise eNo, , "GEO: " & eTx
End Function

' the sheet GEO; made as a copy of a block template (page setup and column widths come with it) when missing
Private Function GeoSheet(wb As Workbook, afterSh As Worksheet, ByRef made As Boolean) As Worksheet
    Dim ws As Worksheet, i As Long, n As Name, sn As String
    If SheetExists(wb, GEO_SHEET) Then Set GeoSheet = wb.Worksheets(GEO_SHEET): Exit Function
    If SheetExists(wb, TPL_GEO_K) Then sn = TPL_GEO_K Else sn = TPL_GEO_Y
    Set ws = MakeCopy(wb, sn, GEO_SHEET)
    ' the block template's own names stay with the template
    For i = ws.Names.Count To 1 Step -1
        Set n = ws.Names(i)
        If NameOnSheet(n, ws) Then
            If Mid$(n.Name, InStr(n.Name, "!") + 1) Like "M_*" Then n.Delete
        End If
    Next i
    If Not afterSh Is Nothing Then MoveAfter ws, afterSh
    made = True
    Set GeoSheet = ws
End Function

' every row, picture and page break of GEO removed
Private Sub ClearGeo(geo As Worksheet)
    Dim i As Long, ur As Range
    On Error Resume Next
    For i = geo.Shapes.Count To 1 Step -1
        geo.Shapes(i).Delete
    Next i
    geo.ResetAllPageBreaks
    Err.Clear
    On Error GoTo 0
    Set ur = geo.UsedRange
    geo.Rows("1:" & (ur.Row + ur.Rows.Count - 1)).Delete
End Sub

' last used row and column of a sheet
Private Sub LastCell(ws As Worksheet, ByRef nR As Long, ByRef nC As Long)
    With ws.UsedRange
        nR = .Row + .Rows.Count - 1
        nC = .Column + .Columns.Count - 1
    End With
End Sub

' ==== base forces from 8'7 ============================================================================================
Private Function RunPullTop(ByVal oaPath As String) As String
    On Error GoTo Fail
    Application.ScreenUpdating = False
    RunPullTop = PullForces(ThisWorkbook, oaPath)
    Application.ScreenUpdating = True
    Exit Function
Fail:
    Application.ScreenUpdating = True
    RunPullTop = "FAIL:" & Err.Description
End Function

' every group with a data sheet: the rows of its piers at the column base from the 8'7 data sheets into M_TEMEL_TESIR;
' returns the log (raises on a fatal error, after closing the 8'7 file it opened)
Private Function PullForces(wb As Workbook, ByVal oaPath As String) As String
    Dim p As String, ob As Workbook, wasOpen As Boolean, ev As Boolean, evSet As Boolean, w As Workbook, out As String
    Dim nG As Long, dsA() As Worksheet, k As Long, idx As Object, blks As Object, ws As Worksheet, el As Range, bl As Range
    Dim v As Variant, i As Long, pier As String, ct As String, elem As String, ent As Collection, e As Variant, sh0 As String
    Dim pv As Variant, outRows As Collection, tb As Range, n As Long, o() As Variant, j As Long, st As Range, miss As String
    Dim saved As String, fn As String, cnt As Long, rw As Variant, two As String, multi As String, eNo As Long, eTx As String
    Dim piers As Collection, src As String, arr As Variant, nDs As Long
    On Error GoTo Bad
    p = FindOaFile(wb, oaPath)
    fn = Mid$(p, InStrRev(p, Application.PathSeparator) + 1)
    For Each w In Application.Workbooks
        If StrComp(w.FullName, p, vbTextCompare) = 0 Then Set ob = w: wasOpen = True: Exit For
    Next w
    If ob Is Nothing Then
        For Each w In Application.Workbooks
            If StrComp(w.Name, fn, vbTextCompare) = 0 Then Err.Raise vbObjectError + 50, , U("Ayn\u0131 adl\u0131 bir dosya ba\u015Fka klas\u00F6rden a\u00E7\u0131k: ") & w.FullName
        Next w
        ev = Application.EnableEvents: evSet = True
        Application.EnableEvents = False
        Set ob = Workbooks.Open(Filename:=p, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
        Application.EnableEvents = ev: evSet = False
    End If
    saved = Format$(FileDateTime(p), "yyyy-mm-dd hh:nn")
    ' piers of the 8'7 data sheets: pier -> (sheet, column tag, base element)
    Set idx = CreateObject("Scripting.Dictionary"): idx.CompareMode = vbTextCompare
    Set blks = CreateObject("Scripting.Dictionary"): blks.CompareMode = vbTextCompare
    For Each ws In ob.Worksheets
        If StrComp(ws.Name, OA_TPL, vbTextCompare) <> 0 Then
            Set el = SheetRangeOpt(ws, OA_ELEM)
            Set bl = SheetRangeOpt(ws, OA_BLOCK)
            If Not el Is Nothing And Not bl Is Nothing Then
                If el.Columns.Count >= 2 And bl.Columns.Count >= 9 Then
                    v = Arr2(el.Columns(2), False)
                    For i = 1 To UBound(v, 1)
                        ParseElemLine CellText(v(i, 1)), pier, ct, elem
                        If Len(pier) > 0 Then
                            If Not idx.Exists(pier) Then idx.Add pier, New Collection
                            Set ent = idx(pier)
                            ent.Add Array(ws.Name, ct, elem)
                        End If
                    Next i
                End If
            End If
        End If
    Next ws
    If idx.Count = 0 Then Err.Raise vbObjectError + 51, , fn & U(": eleman listesi (") & OA_ELEM & U(") dolu bir veri sayfas\u0131 yok; 8'7'de MIDAS_AYAK okunmam\u0131\u015F.")
    nG = NameRange(wb, N_GTAB).Rows.Count
    FindSheets wb, nG, KIND_D, dsA
    For k = 1 To nG
        If Not dsA(k) Is Nothing Then
            nDs = nDs + 1
            Set tb = SheetRangeOpt(dsA(k), M_BASE)
            If tb Is Nothing Then
                out = out & dsA(k).Name & ": " & M_BASE & U(" ad\u0131 yok; tesir yaz\u0131lmad\u0131.") & vbLf
            Else
                If tb.Columns.Count < 10 Then Err.Raise vbObjectError + 52, , dsA(k).Name & ": " & M_BASE & U(" en az 10 s\u00FCtun olmal\u0131 (ayak, kombinasyon, eleman, u\u00E7, Pu, My, Mz, Vy, Vz, T).")
                Set outRows = New Collection
                miss = "": two = "": multi = "": src = ""
                Set piers = PiersOf(wb, k)
                For Each pv In piers
                    If Not idx.Exists(CStr(pv)) Then
                        miss = miss & " " & pv
                    Else
                        Set ent = idx(CStr(pv))
                        e = ent(1)
                        sh0 = CStr(e(0))
                        For Each e In ent
                            If StrComp(CStr(e(0)), sh0, vbTextCompare) <> 0 Then two = two & " " & pv & " (" & sh0 & " / " & CStr(e(0)) & ")"
                            If Len(CStr(e(1))) > 0 And InStr(multi, " " & pv & " ") = 0 Then multi = multi & " " & pv & " "
                        Next e
                        If Not blks.Exists(sh0) Then blks.Add sh0, Arr2(SheetRangeOpt(ob.Worksheets(sh0), OA_BLOCK), False)
                        arr = blks(sh0)
                        cnt = AddPierRows(arr, CStr(pv), ent, outRows)
                        If cnt = 0 Then miss = miss & " " & pv & U("(sat\u0131r yok)")
                        If InStr(src, sh0) = 0 Then src = src & ", " & sh0
                    End If
                Next pv
                n = outRows.Count
                If n > tb.Rows.Count Then Err.Raise vbObjectError + 53, , dsA(k).Name & ": " & n & U(" sat\u0131r, ") & M_BASE & " " & tb.Rows.Count & U(" sat\u0131r al\u0131yor.")
                If n > 0 Then
                    ReDim o(1 To n, 1 To 10)
                    i = 0
                    For Each rw In outRows
                        i = i + 1
                        For j = 0 To 9
                            o(i, j + 1) = rw(j)
                        Next j
                    Next rw
                    tb.Cells(1, 1).Resize(n, 10).Value = o
                End If
                If n < tb.Rows.Count Then tb.Cells(n + 1, 1).Resize(tb.Rows.Count - n, 10).ClearContents
                Set st = SheetRangeOpt(dsA(k), M_BASE_STAMP)
                If st Is Nothing Then
                    out = out & dsA(k).Name & ": " & M_BASE_STAMP & U(" ad\u0131 yok; damga yaz\u0131lmad\u0131.") & vbLf
                Else
                    st.Cells(1, 1).Value = fn & " | " & saved & " | " & Format$(Now, "yyyy-mm-dd hh:nn")
                End If
                out = out & dsA(k).Name & ": " & n & U(" sat\u0131r") & IIf(Len(src) > 0, " (" & Mid$(src, 3) & ")", "") & vbLf
                If Len(miss) > 0 Then out = out & U("  Uyar\u0131: 8'7'de tesiri bulunmayan ayak:") & miss & vbLf
                If Len(two) > 0 Then out = out & U("  Uyar\u0131: iki 8'7 veri sayfas\u0131nda listelenen ayak (ilki al\u0131nd\u0131):") & two & vbLf
                If Len(multi) > 0 Then out = out & U("  Not: birden \u00E7ok kolonlu ayak, her kolon taban\u0131 ayr\u0131 sat\u0131r (toplanmad\u0131):") & Replace(multi, "  ", " ") & vbLf
            End If
        End If
    Next k
    If nDs = 0 Then out = out & U("Grup veri sayfas\u0131 (" & DATA_PFX & "<ayaklar>) yok; \u00F6nce 'Temel grubu sayfalar\u0131' \u00E7al\u0131\u015Ft\u0131r\u0131lmal\u0131.") & vbLf
    If Not wasOpen Then ob.Close SaveChanges:=False
    Set ob = Nothing
    Application.Calculate
    PullForces = U("Temel tesirleri: ") & fn & U(" (kay\u0131t ") & saved & "), " & idx.Count & U(" ayak 8'7 eleman listelerinde.") & vbLf & out
    Exit Function
Bad:
    eNo = Err.Number: eTx = Err.Description
    On Error Resume Next
    If evSet Then Application.EnableEvents = ev
    If Not ob Is Nothing And Not wasOpen Then ob.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise eNo, , eTx
End Function

' the 8'7 workbook: the given file, or the one 8'7_*.xlsm in this workbook's folder, else (window only) a file dialog
Private Function FindOaFile(wb As Workbook, ByVal oaPath As String) As String
    Dim f As String, n As Long, one As String, lst As String, v As Variant, sep As String
    If Len(oaPath) > 0 Then
        If Len(Dir$(oaPath)) = 0 Then Err.Raise vbObjectError + 54, , U("8'7 dosyas\u0131 yok: ") & oaPath
        FindOaFile = oaPath
        Exit Function
    End If
    If Len(wb.Path) = 0 Then Err.Raise vbObjectError + 54, , U("Bu dosya hen\u00FCz kaydedilmemi\u015F; 8'7 dosyas\u0131 klas\u00F6r\u00FCnde aranam\u0131yor.")
    sep = Application.PathSeparator
    f = Dir$(wb.Path & sep & OA_PATTERN)
    Do While Len(f) > 0
        n = n + 1: one = f: lst = lst & "  " & f & vbLf
        f = Dir$()
    Loop
    If n = 1 Then FindOaFile = wb.Path & sep & one: Exit Function
    If scripted Then
        If n = 0 Then Err.Raise vbObjectError + 54, , U("Klas\u00F6rde 8'7 dosyas\u0131 yok (") & OA_PATTERN & "): " & wb.Path
        Err.Raise vbObjectError + 54, , U("Klas\u00F6rde birden \u00E7ok 8'7 dosyas\u0131 var:") & vbLf & lst
    End If
    On Error Resume Next
    ChDrive wb.Path
    ChDir wb.Path
    On Error GoTo 0
    v = Application.GetOpenFilename(U("Excel (*.xlsm),*.xlsm"), , U("8'7 orta ayak dosyas\u0131n\u0131 se\u00E7in"))
    If VarType(v) = vbBoolean Then Err.Raise vbObjectError + 55, , U("8'7 dosyas\u0131 se\u00E7ilmedi; tesir \u00E7ekilmedi.")
    FindOaFile = CStr(v)
End Function

' 'P1: 473-477, base 477 J' -> P1, "", 477; 'P1 K2: 480-484, base 484 J, top 480 I' -> P1, K2, 484; other lines -> ""
Private Sub ParseElemLine(ByVal s As String, ByRef pier As String, ByRef ct As String, ByRef elem As String)
    Dim p As Long, head As String, q As Long, b As Long, j As Long
    pier = "": ct = "": elem = ""
    p = InStr(s, ":")
    If p < 2 Then Exit Sub
    head = Trim$(Left$(s, p - 1))
    pier = head
    q = InStrRev(head, " K")
    If q > 1 Then
        If AllDigits(Mid$(head, q + 2)) Then
            pier = Trim$(Left$(head, q - 1))
            ct = "K" & Mid$(head, q + 2)
        End If
    End If
    b = InStr(1, s, "base ", vbTextCompare)
    If b > 0 Then
        j = b + 5
        Do While Mid$(s, j, 1) Like "#"
            elem = elem & Mid$(s, j, 1)
            j = j + 1
        Loop
    End If
End Sub

Private Function AllDigits(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    AllDigits = (s Like String$(Len(s), "#"))
End Function

' rows of one pier at the column base: names '<combination>_<pier>[K<column>]_J' of the 8'7 block (counter, No, name, Pu,
' My, Mz, Vy, Vz, T) as Array(pier, combination, element, end, Pu, My, Mz, Vy, Vz, T); returns the number added
Private Function AddPierRows(arr As Variant, ByVal pier As String, ent As Collection, outRows As Collection) As Long
    Dim i As Long, nm As String, a As Long, b As Long, endTag As String, tok As String, comb As String, ct As String
    Dim elem As Variant, e As Variant
    For i = 1 To UBound(arr, 1)
        nm = CellText(arr(i, 3))
        a = InStrRev(nm, "_")
        If a > 1 Then
            endTag = Mid$(nm, a + 1)
            If StrComp(endTag, BASE_END, vbTextCompare) = 0 Then
                b = InStrRev(nm, "_", a - 1)
                If b > 1 Then
                    tok = Mid$(nm, b + 1, a - b - 1)
                    comb = Left$(nm, b - 1)
                    ct = PierToken(tok, pier)
                    If ct <> "-" Then
                        elem = Empty
                        For Each e In ent
                            If StrComp(CStr(e(1)), ct, vbTextCompare) = 0 Then
                                If Len(CStr(e(2))) > 0 Then elem = CLng(e(2))
                                Exit For
                            End If
                        Next e
                        outRows.Add Array(pier, comb, elem, endTag, arr(i, 4), arr(i, 5), arr(i, 6), arr(i, 7), arr(i, 8), arr(i, 9))
                        AddPierRows = AddPierRows + 1
                    End If
                End If
            End If
        End If
    Next i
End Function

' "" = the pier itself (one column), "K<n>" = its column n, "-" = another pier
Private Function PierToken(ByVal tok As String, ByVal pier As String) As String
    PierToken = "-"
    If StrComp(tok, pier, vbTextCompare) = 0 Then PierToken = "": Exit Function
    If Len(tok) > Len(pier) + 1 Then
        If StrComp(Left$(tok, Len(pier) + 1), pier & "K", vbTextCompare) = 0 Then
            If AllDigits(Mid$(tok, Len(pier) + 2)) Then PierToken = "K" & Mid$(tok, Len(pier) + 2)
        End If
    End If
End Function

' ==== template show / hide ============================================================================================
Private Sub ShowTemplates()
    Dim wb As Workbook, tpl As Worksheet, ws As Worksheet, nm As Variant, rg As Range, g As Variant, i As Long, lst As String
    Set wb = ThisWorkbook
    If Not SheetExists(wb, TPL_G) Then ShowMsg U("\u015Eablon sayfa yok: ") & TPL_G, vbExclamation: Exit Sub
    Set tpl = wb.Worksheets(TPL_G)
    If tpl.Visible = xlSheetVisible Then
        If ActiveWorkbook Is wb Then
            If IsTemplateSheet(ActiveSheet.Name) Then
                For Each ws In wb.Worksheets
                    If ws.Visible = xlSheetVisible And Not IsTemplateSheet(ws.Name) Then ws.Activate: Exit For
                Next ws
            End If
        End If
        ' back to the pages of its own gate (rows, pictures, print area), as the group sheets are
        Application.ScreenUpdating = False
        CloseRows tpl
        Application.ScreenUpdating = True
        For Each nm In Array(TPL_G, TPL_Y, TPL_GEO_K, TPL_GEO_Y)
            If SheetExists(wb, CStr(nm)) Then wb.Worksheets(CStr(nm)).Visible = xlSheetVeryHidden
        Next nm
        ShowMsg U("\u015Eablonlar gizlendi. De\u011Fi\u015Fiklikleri grup sayfalar\u0131na ta\u015F\u0131mak i\u00E7in: 'Temel \u015Fablonlar\u0131n\u0131 grup sayfalar\u0131na uygula'. GEO sayfas\u0131 bir sonraki 'Temel grubu sayfalar\u0131' \u00E7al\u0131\u015Fmas\u0131nda yeniden kurulur."), vbInformation
    Else
        For Each nm In Array(TPL_G, TPL_Y, TPL_GEO_K, TPL_GEO_Y)
            If SheetExists(wb, CStr(nm)) Then
                wb.Worksheets(CStr(nm)).Visible = xlSheetVisible
                lst = lst & ", " & CStr(nm)
            End If
        Next nm
        ' every foundation type editable: all gated rows and their pictures open
        On Error Resume Next
        Application.ScreenUpdating = False
        Set rg = SheetRangeOpt(tpl, M_GATE)
        If Not rg Is Nothing Then
            rg.EntireRow.Hidden = False
            g = Arr2(rg, False)
            For i = 1 To UBound(g, 1): g(i, 1) = 1: Next i
            GateShapes tpl, g, rg.Row
        End If
        Application.ScreenUpdating = True
        tpl.Activate
        On Error GoTo 0
        ShowMsg U("\u015Eablonlar a\u00E7\u0131k: ") & Mid$(lst, 3) & vbLf & _
            U("KAZIKLI ve YUZEYSEL sayfalar\u0131n\u0131n b\u00FCt\u00FCn sat\u0131rlar\u0131 ve resimleri a\u00E7\u0131k (gizlenince eski h\u00E2line d\u00F6ner). ") & _
            U("Burada yaln\u0131z t\u00FCm gruplarda ge\u00E7erli olacak d\u00FCzeltmeleri yap\u0131n; gruba \u00F6zel de\u011Ferleri grup sayfas\u0131na yaz\u0131n. Sonra 'Temel \u015Fablonlar\u0131n\u0131 grup sayfalar\u0131na uygula'."), vbInformation
    End If
End Sub

Private Sub ShowDataTemplate()
    Dim wb As Workbook, tpl As Worksheet, ws As Worksheet, nR As Long, nC As Long
    Set wb = ThisWorkbook
    If Not SheetExists(wb, TPL_D) Then ShowMsg U("\u015Eablon sayfa yok: ") & TPL_D, vbExclamation: Exit Sub
    Set tpl = wb.Worksheets(TPL_D)
    If tpl.Visible = xlSheetVisible Then
        If ActiveSheet Is tpl Then
            For Each ws In wb.Worksheets
                If ws.Visible = xlSheetVisible And Not ws Is tpl Then ws.Activate: Exit For
            Next ws
        End If
        tpl.Visible = xlSheetVeryHidden
        ShowMsg U("Veri \u015Fablonu gizlendi. De\u011Fi\u015Fiklikleri veri sayfalar\u0131na ta\u015F\u0131mak i\u00E7in: 'Temel veri \u015Fablonunu veri sayfalar\u0131na uygula'."), vbInformation
    Else
        ' the first apply compares with the template as it is before the owner edits it
        If Not SheetExists(wb, BASE_D) Then
            Application.ScreenUpdating = False
            LastCell tpl, nR, nC
            WriteBaseline wb, tpl, nR, nC, BASE_D
            Application.ScreenUpdating = True
        End If
        tpl.Visible = xlSheetVisible
        tpl.Activate
        ShowMsg U("Veri \u015Fablonu a\u00E7\u0131k: ") & TPL_D & vbLf & _
            U("Burada yaln\u0131z b\u00FCt\u00FCn gruplarda ge\u00E7erli olacak d\u00FCzeni (ba\u015Fl\u0131klar, h\u00FCcre yerleri, form\u00FCller, bi\u00E7im) de\u011Fi\u015Ftirin; gruba \u00F6zel \u00E7\u0131kt\u0131lar\u0131 grubun MIDAS_ sayfas\u0131na yap\u0131\u015Ft\u0131r\u0131n. Sonra 'Temel veri \u015Fablonunu veri sayfalar\u0131na uygula'. ") & _
            U("Bir h\u00FCcrenin yeri de\u011Fi\u015Fince onu okuyan rapor form\u00FCl\u00FCn\u00FC SABLON_TEMEL'de d\u00FCzeltin."), vbInformation
    End If
End Sub

' ==== template apply (group sheets and _Y sheets) =====================================================================
Private Function RunTemplateApply(ByVal doApply As Boolean, ByVal force As Boolean) As String
    Dim wb As Workbook, kd As Long, tplN(1 To 2) As String, baseN(1 To 2) As String, shpN(1 To 2) As String
    Dim has(1 To 2) As Boolean, tpl(1 To 2) As Worksheet, nR(1 To 2) As Long, nC(1 To 2) As Long, nT(1 To 2) As Long
    Dim t1 As Variant, b1 As Variant, t2 As Variant, b2 As Variant, sk1() As Boolean, sk2() As Boolean
    Dim tplHid(1 To 2) As Boolean, chg(1 To 2) As Collection, picNote As String, picLst As String, q As Variant
    Dim nG As Long, gsA() As Worksheet, ysA() As Worksheet, dsA() As Worksheet, k As Long, own As Worksheet
    Dim upd() As Collection, con() As Collection, nUpd As Long, nCon As Long, it As Variant, rep As String, conList As String
    Dim writeCon As Boolean, shrunk As Boolean, cm As Long, eTx As String, wasScripted As Boolean, anyTpl As Boolean
    Dim sft() As Long, pEnd() As Long, skipG() As Boolean, nMoved As Long, picMiss As String, nPic As Long
    On Error GoTo Fail
    wasScripted = scripted
    lg = ""
    Set wb = ThisWorkbook
    tplN(1) = TPL_G: baseN(1) = BASE_G: shpN(1) = BASE_G_R
    tplN(2) = TPL_Y: baseN(2) = BASE_Y: shpN(2) = BASE_Y_R
    cm = Application.Calculation
    Application.ScreenUpdating = False
    For kd = 1 To 2
        Set chg(kd) = New Collection
        If SheetExists(wb, tplN(kd)) Then
            anyTpl = True
            Set tpl(kd) = wb.Worksheets(tplN(kd))
            tplHid(kd) = (tpl(kd).Visible <> xlSheetVisible)
            LastCell tpl(kd), nR(kd), nC(kd)
            If SheetExists(wb, baseN(kd)) Then
                has(kd) = True
            Else
                WriteBaseline wb, tpl(kd), nR(kd), nC(kd), baseN(kd)
                WriteShapeBase wb, tpl(kd), shpN(kd)
                lg = lg & U("\u0130z sayfas\u0131 (") & baseN(kd) & U(") yoktu: ") & tplN(kd) & U(" \u015Fablonunun bug\u00FCnk\u00FC h\u00E2li kaydedildi, sayfalara dokunulmad\u0131. Bundan sonra bu \u015Fablonda yapt\u0131\u011F\u0131n\u0131z de\u011Fi\u015Fiklikler bu tu\u015Fla ta\u015F\u0131n\u0131r.") & vbLf
            End If
        End If
    Next kd
    If Not anyTpl Then Err.Raise vbObjectError + 40, , U("\u015Eablon sayfa yok: ") & TPL_G & " / " & TPL_Y
    ' cells and pictures changed on the templates since the last apply
    For kd = 1 To 2
        If has(kd) Then
            If kd = 1 Then
                LoadKind wb, tpl(1), baseN(1), nR(1), nC(1), t1, b1, sk1
                nT(1) = CountChanged(t1, b1, sk1, nR(1), nC(1))
            Else
                LoadKind wb, tpl(2), baseN(2), nR(2), nC(2), t2, b2, sk2
                nT(2) = CountChanged(t2, b2, sk2, nR(2), nC(2))
            End If
            If SheetExists(wb, shpN(kd)) Then
                OpenRows tpl(kd)
                picNote = "": picLst = ""
                Set chg(kd) = ShapeChanges(wb, tpl(kd), shpN(kd), picNote)
                For Each q In chg(kd)
                    picLst = picLst & "  " & CStr(q(1)) & " (" & tpl(kd).Cells(CLng(q(2)), CLng(q(3))).Address(False, False) & ")" & vbLf
                Next q
                If chg(kd).Count > 0 Then lg = lg & tplN(kd) & U(": yeri / boyutu de\u011Fi\u015Fen resimler (sayfadaki resim ayn\u0131 yere ve boyuta getirilir, g\u00F6r\u00FCnt\u00FCs\u00FC de\u011Fi\u015Fmez):") & vbLf & picLst
                If Len(picNote) > 0 Then lg = lg & tplN(kd) & U(": yeni (+) veya silinmi\u015F (-) resimler ta\u015F\u0131nmaz; sayfalarda elle yap\u0131n:") & vbLf & picNote
            Else
                WriteShapeBase wb, tpl(kd), shpN(kd)
                lg = lg & U("Resim izi (") & shpN(kd) & U(") yoktu: \u015Fablondaki resimlerin bug\u00FCnk\u00FC yeri kaydedildi; bundan sonra ta\u015F\u0131n\u0131r.") & vbLf
            End If
        End If
    Next kd
    nPic = chg(1).Count + chg(2).Count
    If nT(1) + nT(2) + nPic = 0 Then
        For kd = 1 To 2
            If tplHid(kd) Then CloseRows tpl(kd)
        Next kd
        Application.ScreenUpdating = True
        RunTemplateApply = lg & U("\u015Eablonlarda son uygulamadan beri de\u011Fi\u015Fen h\u00FCcre veya yeri de\u011Fi\u015Fen resim yok; sayfalara dokunulmad\u0131.")
        Exit Function
    End If

    nG = NameRange(wb, N_GTAB).Rows.Count
    FindSheets wb, nG, KIND_G, gsA
    FindSheets wb, nG, KIND_Y, ysA
    FindSheets wb, nG, KIND_D, dsA
    ReDim upd(1 To 2, 1 To nG): ReDim con(1 To 2, 1 To nG): ReDim sft(1 To nG): ReDim pEnd(1 To nG): ReDim skipG(1 To nG)
    ' rows the per-pile copies add below the pattern page now (conflicts are listed at the sheet's own rows)
    For k = 1 To nG
        If Not gsA(k) Is Nothing Then
            PileShift gsA(k), sft(k), pEnd(k)
            If sft(k) > 0 And dsA(k) Is Nothing And nT(1) > 0 Then
                skipG(k) = True
                lg = lg & U("Uyar\u0131: ") & gsA(k).Name & U(" sayfas\u0131nda kaz\u0131k tablosu ek sayfalar\u0131 var ama veri sayfas\u0131 yok; bu sayfa kar\u015F\u0131la\u015Ft\u0131r\u0131lmad\u0131.") & vbLf
            End If
        End If
    Next k
    ' per-pile copies out first: Excel moves every reference back to the template rows
    If nT(1) > 0 Then
        shrunk = True
        For k = 1 To nG
            If Not gsA(k) Is Nothing And Not skipG(k) Then FitPilePages gsA(k), dsA(k), "", 0
        Next k
    End If
    If Application.Calculation <> xlCalculationManual Then Application.Calculation = xlCalculationManual
    For k = 1 To nG
        Set upd(1, k) = New Collection: Set con(1, k) = New Collection
        Set upd(2, k) = New Collection: Set con(2, k) = New Collection
        If nT(1) > 0 And Not gsA(k) Is Nothing And Not skipG(k) Then
            PlanSheet t1, b1, sk1, nR(1), nC(1), gsA(k), gsA(k), ysA(k), dsA(k), upd(1, k), con(1, k)
            rep = rep & "  " & gsA(k).Name & ": " & upd(1, k).Count & U(" h\u00FCcre") & IIf(con(1, k).Count > 0, ", " & con(1, k).Count & U(" \u00E7ak\u0131\u015Fma"), "") & vbLf
            conList = conList & ConflictText(t1, con(1, k), tpl(1), gsA(k).Name, sft(k), pEnd(k), Len(conList))
        End If
        If nT(2) > 0 And Not ysA(k) Is Nothing Then
            PlanSheet t2, b2, sk2, nR(2), nC(2), ysA(k), gsA(k), ysA(k), dsA(k), upd(2, k), con(2, k)
            rep = rep & "  " & ysA(k).Name & ": " & upd(2, k).Count & U(" h\u00FCcre") & IIf(con(2, k).Count > 0, ", " & con(2, k).Count & U(" \u00E7ak\u0131\u015Fma"), "") & vbLf
            conList = conList & ConflictText(t2, con(2, k), tpl(2), ysA(k).Name, 0, 0, Len(conList))
        End If
        nUpd = nUpd + upd(1, k).Count + upd(2, k).Count
        nCon = nCon + con(1, k).Count + con(2, k).Count
    Next k
    If nT(1) + nT(2) > 0 Then lg = lg & U("\u015Eablonlarda son uygulamadan beri de\u011Fi\u015Fen h\u00FCcreler, sayfalara:") & vbLf & rep
    If nCon > 0 Then lg = lg & U("Sayfada elle de\u011Fi\u015Ftirilmi\u015F (\u00E7ak\u0131\u015Fan) h\u00FCcreler:") & vbLf & conList

    If nUpd + nCon + nPic = 0 Then
        lg = lg & U("Yaz\u0131lacak bir \u015Fey yok.") & vbLf
        doApply = False
    ElseIf doApply And Not wasScripted Then
        Application.ScreenUpdating = True
        If Ask(lg & vbLf & nUpd & U(" h\u00FCcre ve ") & nPic & U(" resim \u015Fablona g\u00F6re yaz\u0131ls\u0131n m\u0131?")) <> vbYes Then doApply = False
        If doApply And nCon > 0 Then writeCon = (Ask(nCon & U(" \u00E7ak\u0131\u015Fan h\u00FCcre de \u015Fablona g\u00F6re yaz\u0131ls\u0131n m\u0131? (Hay\u0131r = sayfadaki de\u011Ferler kal\u0131r)")) = vbYes)
        Application.ScreenUpdating = False
    Else
        writeCon = force
    End If

    If doApply Then
        For k = 1 To nG
            If Not gsA(k) Is Nothing Then
                For Each it In upd(1, k)
                    PutCell tpl(1), gsA(k), gsA(k), ysA(k), dsA(k), CLng(it(0)), CLng(it(1)), CellText(t1(it(0), it(1)))
                Next it
                If writeCon Then
                    For Each it In con(1, k)
                        PutCell tpl(1), gsA(k), gsA(k), ysA(k), dsA(k), CLng(it(0)), CLng(it(1)), CellText(t1(it(0), it(1)))
                    Next it
                End If
            End If
            If Not ysA(k) Is Nothing Then
                For Each it In upd(2, k)
                    PutCell tpl(2), ysA(k), gsA(k), ysA(k), dsA(k), CLng(it(0)), CLng(it(1)), CellText(t2(it(0), it(1)))
                Next it
                If writeCon Then
                    For Each it In con(2, k)
                        PutCell tpl(2), ysA(k), gsA(k), ysA(k), dsA(k), CLng(it(0)), CLng(it(1)), CellText(t2(it(0), it(1)))
                    Next it
                End If
            End If
        Next k
        ' pictures: place and size only
        For kd = 1 To 2
            If chg(kd).Count > 0 Then
                nMoved = 0
                For k = 1 To nG
                    If kd = 1 Then Set own = gsA(k) Else Set own = ysA(k)
                    If Not own Is Nothing Then
                        If Not (kd = 1 And skipG(k)) Then
                            OpenRows own
                            If kd = 1 And Not shrunk Then
                                nMoved = nMoved + MoveShapes(own, chg(kd), sft(k), pEnd(k), picMiss)
                                CloseRows own
                            Else
                                nMoved = nMoved + MoveShapes(own, chg(kd), 0, 0, picMiss)
                            End If
                        End If
                    End If
                Next k
                lg = lg & tplN(kd) & U(": yeri / boyutu d\u00FCzeltilen resim: ") & nMoved & vbLf
                WriteShapeBase wb, tpl(kd), shpN(kd), True
            End If
        Next kd
        If Len(picMiss) > 0 Then lg = lg & U("Bulunamayan (o sayfada silinmi\u015F veya de\u011Fi\u015Ftirilmi\u015F) resimler:") & vbLf & picMiss
        If nT(1) > 0 Then WriteBaseline wb, tpl(1), nR(1), nC(1), baseN(1)
        If nT(2) > 0 Then WriteBaseline wb, tpl(2), nR(2), nC(2), baseN(2)
        lg = lg & U("Yaz\u0131ld\u0131: ") & (nUpd + IIf(writeCon, nCon, 0)) & U(" h\u00FCcre. \u0130z sayfalar\u0131 \u015Fablonlarla e\u015Flendi.") & vbLf
    Else
        lg = lg & U("Hi\u00E7bir h\u00FCcre yaz\u0131lmad\u0131.") & vbLf
    End If

    ' per-pile copies back, gates, print areas, numbers, GEO, lists
    If Application.Calculation <> cm Then Application.Calculation = cm
    If shrunk Then
        lg = lg & U("Kaz\u0131k tablosu ek sayfalar\u0131 kar\u015F\u0131la\u015Ft\u0131rma i\u00E7in \u00E7\u0131kar\u0131l\u0131p yeniden kuruldu ('ek sayfa 0 ->' sat\u0131rlar\u0131):") & vbLf
        shrunk = False
        RefreshCore wb
    End If
    For kd = 1 To 2
        If tplHid(kd) Then CloseRows tpl(kd)
    Next kd
    scripted = wasScripted
    Application.ScreenUpdating = True
    RunTemplateApply = lg
    Exit Function
Fail:
    eTx = Err.Description
    On Error Resume Next
    If cm <> 0 Then
        If Application.Calculation <> cm Then Application.Calculation = cm
    End If
    If shrunk Then RefreshCore wb
    For kd = 1 To 2
        If tplHid(kd) Then CloseRows tpl(kd)
    Next kd
    scripted = wasScripted
    Application.ScreenUpdating = True
    RunTemplateApply = "FAIL:" & eTx & IIf(Len(lg) > 0, vbLf & lg, "")
End Function

' template formulas, baseline texts and the cells not compared (chain, page index, array formulas)
Private Sub LoadKind(wb As Workbook, tpl As Worksheet, ByVal baseName As String, ByRef nR As Long, ByRef nC As Long, ByRef t As Variant, ByRef b As Variant, sk() As Boolean)
    Dim bs As Worksheet, ha As Variant, r As Long, c As Long
    Set bs = wb.Worksheets(baseName)
    With bs.UsedRange
        If .Row + .Rows.Count - 1 > nR Then nR = .Row + .Rows.Count - 1
        If .Column + .Columns.Count - 1 > nC Then nC = .Column + .Columns.Count - 1
    End With
    If nR < 2 Then nR = 2
    If nC < 2 Then nC = 2
    t = tpl.Range(tpl.Cells(1, 1), tpl.Cells(nR, nC)).Formula
    b = bs.Range(bs.Cells(1, 1), bs.Cells(nR, nC)).Value
    ReDim sk(1 To nR, 1 To nC)
    MarkSkip tpl, M_CHAIN, sk, nR, nC
    MarkSkip tpl, M_PAGE, sk, nR, nC
    ha = tpl.Range(tpl.Cells(1, 1), tpl.Cells(nR, nC)).HasArray
    If IsNull(ha) Then
        For r = 1 To nR
            For c = 1 To nC
                If Left$(CellText(t(r, c)), 1) = "=" Then
                    If tpl.Cells(r, c).HasArray Then sk(r, c) = True
                End If
            Next c
        Next r
    ElseIf ha = True Then
        Err.Raise vbObjectError + 41, , tpl.Name & U(": \u015Fablonun tamam\u0131 dizi form\u00FCl\u00FC; uygulanamaz.")
    End If
End Sub

Private Function CountChanged(t As Variant, b As Variant, sk() As Boolean, ByVal nR As Long, ByVal nC As Long) As Long
    Dim r As Long, c As Long
    For r = 1 To nR
        For c = 1 To nC
            If Not sk(r, c) Then
                If CellText(t(r, c)) <> CellText(b(r, c)) Then CountChanged = CountChanged + 1
            End If
        Next c
    Next r
End Function

' one sheet against the template: upd = cells still at the baseline, con = cells changed on the sheet itself
Private Sub PlanSheet(t As Variant, b As Variant, sk() As Boolean, ByVal nR As Long, ByVal nC As Long, own As Worksheet, gs As Worksheet, ys As Worksheet, ds As Worksheet, upd As Collection, con As Collection)
    Dim g As Variant, r As Long, c As Long, tv As String, bv As String, gv As String
    g = own.Range(own.Cells(1, 1), own.Cells(nR, nC)).Formula
    For r = 1 To nR
        For c = 1 To nC
            If Not sk(r, c) Then
                tv = CellText(t(r, c)): bv = CellText(b(r, c))
                If tv <> bv Then
                    gv = ToTemplate(CellText(g(r, c)), gs, ys, ds)
                    If gv = bv Then
                        upd.Add Array(r, c)
                    ElseIf gv <> tv Then
                        con.Add Array(r, c, CellText(g(r, c)))
                    End If
                End If
            End If
        Next c
    Next r
End Sub

' conflict lines of one sheet, at the sheet's own rows (rows below lastRow lie sft lower), while the list is short
Private Function ConflictText(t As Variant, con As Collection, tpl As Worksheet, ByVal sheetName As String, ByVal sft As Long, ByVal lastRow As Long, ByVal room As Long) As String
    Dim it As Variant, r As Long
    For Each it In con
        If room + Len(ConflictText) >= 3000 Then Exit Function
        r = CLng(it(0))
        If lastRow > 0 And r > lastRow Then r = r + sft
        ConflictText = ConflictText & "  " & sheetName & "!" & tpl.Cells(r, CLng(it(1))).Address(False, False) & _
            U(": sayfada ") & Left$(CStr(it(2)), 60) & U(" | \u015Fablonda ") & Left$(CellText(t(it(0), it(1))), 60) & vbLf
    Next it
End Function

' one cell takes the template cell (number format too: a formula in a Text cell would stay text)
Private Sub PutCell(tpl As Worksheet, own As Worksheet, gs As Worksheet, ys As Worksheet, ds As Worksheet, ByVal r As Long, ByVal c As Long, ByVal tv As String)
    Dim cl As Range
    Set cl = own.Cells(r, c)
    If cl.NumberFormat <> tpl.Cells(r, c).NumberFormat Then cl.NumberFormat = tpl.Cells(r, c).NumberFormat
    If Len(tv) = 0 Then
        cl.ClearContents
    ElseIf Left$(tv, 1) = "=" Then
        cl.Formula = ToGroup(tv, gs, ys, ds)
    Else
        cl.Value = tpl.Cells(r, c).Value
    End If
End Sub

' baseline = the template's formulas and constants as text (apostrophe prefix: nothing calculates there)
Private Sub WriteBaseline(wb As Workbook, tpl As Worksheet, ByVal nR As Long, ByVal nC As Long, ByVal baseName As String)
    Dim bs As Worksheet, t As Variant, r As Long, c As Long
    If SheetExists(wb, baseName) Then
        Set bs = wb.Worksheets(baseName)
        bs.Cells.ClearContents
    Else
        Set bs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        bs.Name = baseName
    End If
    If nR < 2 Then nR = 2
    If nC < 2 Then nC = 2
    t = tpl.Range(tpl.Cells(1, 1), tpl.Cells(nR, nC)).Formula
    For r = 1 To nR
        For c = 1 To nC
            If Len(CellText(t(r, c))) > 0 Then t(r, c) = "'" & CellText(t(r, c)) Else t(r, c) = Empty
        Next c
    Next r
    bs.Range(bs.Cells(1, 1), bs.Cells(nR, nC)).Value = t
    bs.Visible = xlSheetVeryHidden
End Sub

' cells of a sheet-level name marked as not compared
Private Sub MarkSkip(ws As Worksheet, ByVal nm As String, sk() As Boolean, ByVal nR As Long, ByVal nC As Long)
    Dim rg As Range, cl As Range
    Set rg = SheetRangeOpt(ws, nm)
    If rg Is Nothing Then Exit Sub
    For Each cl In rg.Cells
        If cl.Row <= nR And cl.Column <= nC Then sk(cl.Row, cl.Column) = True
    Next cl
End Sub

' cells of every area of a range marked as not compared
Private Sub MarkSkipArea(rg As Range, sk() As Boolean, ByVal nR As Long, ByVal nC As Long)
    Dim ar As Range, r As Long, c As Long, rE As Long, cE As Long
    For Each ar In rg.Areas
        rE = ar.Row + ar.Rows.Count - 1: If rE > nR Then rE = nR
        cE = ar.Column + ar.Columns.Count - 1: If cE > nC Then cE = nC
        For r = ar.Row To rE
            For c = ar.Column To cE
                sk(r, c) = True
            Next c
        Next r
    Next ar
End Sub

' the areas of the sheet-level names of ws that match DATA_SKIP (M_TESIR_*, M_KAZIK_*, M_GSD_*, M_TEMEL_TESIR*)
Private Sub MarkSkipNames(ws As Worksheet, sk() As Boolean, ByVal nR As Long, ByVal nC As Long)
    Dim n As Name, sn As String, pt As Variant, rg As Range
    For Each n In ws.Names
        If NameOnSheet(n, ws) Then
            sn = UCase$(Mid$(n.Name, InStr(n.Name, "!") + 1))
            For Each pt In Split(DATA_SKIP, "|")
                If sn Like CStr(pt) Then
                    Set rg = Nothing
                    On Error Resume Next
                    Set rg = n.RefersToRange
                    On Error GoTo 0
                    If Not rg Is Nothing Then
                        If rg.Worksheet.Name = ws.Name Then MarkSkipArea rg, sk, nR, nC
                    End If
                    Exit For
                End If
            Next pt
        End If
    Next n
End Sub

' ==== pictures and drawings: place and size only (as ORTA_AYAK 16h) ===================================================
' every gated row open, so that a picture's place and size are its real ones (hidden rows have no height)
Private Sub OpenRows(ws As Worksheet)
    Dim rg As Range
    Set rg = SheetRangeOpt(ws, M_GATE)
    If Not rg Is Nothing Then rg.EntireRow.Hidden = False
End Sub

Private Sub CloseRows(ws As Worksheet)
    If SheetRangeOpt(ws, M_GATE) Is Nothing Then Exit Sub
    On Error Resume Next
    ApplyGate ws
    On Error GoTo 0
End Sub

' "id|name|row|column|x in the cell|y in the cell|width|height" (points) of every picture / drawing; comments and controls
' are left out, as in GateShapes. The id stays the same on the copies of the template.
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

' Top at which the shape starts exactly at the top edge of row r. Measured 2026-11-16 (8'7): on a sheet whose rows were just
' unhidden, Shape.Top and Range.Top do not agree, so a place is never taken from Range.Top: the shape is moved until its
' TopLeftCell changes row. Rows must be open.
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

Private Function ReadShapeBase(wb As Workbook, ByVal baseName As String) As Collection
    Dim c As New Collection, ws As Worksheet, r As Long
    Set ws = wb.Worksheets(baseName)
    r = 1
    Do While Len(CellText(ws.Cells(r, 1).Value)) > 0
        c.Add CellText(ws.Cells(r, 1).Value)
        r = r + 1
    Loop
    Set ReadShapeBase = c
End Function

Private Sub WriteShapeList(wb As Workbook, lst As Collection, ByVal baseName As String)
    Dim bs As Worksheet, q As Variant, r As Long
    If SheetExists(wb, baseName) Then
        Set bs = wb.Worksheets(baseName)
        bs.Cells.ClearContents
    Else
        Set bs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        bs.Name = baseName
    End If
    For Each q In lst
        r = r + 1
        bs.Cells(r, 1).Value = "'" & CStr(q)
    Next q
    bs.Visible = xlSheetVeryHidden
End Sub

' the template's pictures as they are now into the list baseName; isOpen = its rows are open and stay so
Private Sub WriteShapeBase(wb As Workbook, tpl As Worksheet, ByVal baseName As String, Optional ByVal isOpen As Boolean = False)
    Dim lst As Collection
    If Not isOpen Then OpenRows tpl
    Set lst = GeoList(tpl)
    If Not isOpen Then
        If tpl.Visible <> xlSheetVisible Then CloseRows tpl
    End If
    WriteShapeList wb, lst, baseName
End Sub

' template pictures whose place or size changed since the last apply: Array(id, name, row, column, x, y, w, h);
' note = pictures that are new or gone on the template (not carried: the image itself is never copied)
Private Function ShapeChanges(wb As Workbook, tpl As Worksheet, ByVal baseName As String, ByRef note As String) As Collection
    Dim c As New Collection, cur As Collection, old As Collection, q As Variant, o As Variant
    Dim a() As String, b() As String, hit As Boolean, i As Long, dif As Boolean
    Set cur = GeoList(tpl)
    Set old = ReadShapeBase(wb, baseName)
    For Each q In cur
        a = Split(CStr(q), "|")
        hit = False
        For Each o In old
            b = Split(CStr(o), "|")
            If UBound(b) >= 7 Then
                If b(0) = a(0) Then
                    hit = True
                    ' place within 2 points, also across a cell border; sizes 3 points + 1.5 % (a 'move and size' picture
                    ' reads a few points off from run to run)
                    dif = Not SameSpot(tpl, CLng(Val(a(2))), Val(a(5)), CLng(Val(b(2))), Val(b(5)), True)
                    If Not dif Then dif = Not SameSpot(tpl, CLng(Val(a(3))), Val(a(4)), CLng(Val(b(3))), Val(b(4)), False)
                    For i = 6 To 7
                        If Abs(Val(b(i)) - Val(a(i))) > 3 + 0.015 * Val(a(i)) Then dif = True
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

' two places (row or column + offset in points) are the same within 2 points. Measured 2026-11-26 (Picture 39 of
' SABLON_TEMEL, top edge on a row border): the top-left cell reads as row 753 with offset -0.2 in one run and as row 752
' with offset = its height - 0.2 in the next, so the place is compared through the row heights / column widths between.
Private Function SameSpot(ws As Worksheet, ByVal i1 As Long, ByVal o1 As Double, ByVal i2 As Long, ByVal o2 As Double, ByVal isRow As Boolean) As Boolean
    Dim lo As Long, hi As Long, k As Long, span As Double
    If i1 = i2 Then SameSpot = (Abs(o1 - o2) <= 2): Exit Function
    If Abs(i1 - i2) > 3 Then Exit Function
    If i1 < i2 Then lo = i1: hi = i2 Else lo = i2: hi = i1
    For k = lo To hi - 1
        If isRow Then span = span + ws.Rows(k).RowHeight Else span = span + ws.Columns(k).Width
    Next k
    If i1 < i2 Then SameSpot = (Abs((o1 - span) - o2) <= 2) Else SameSpot = (Abs((o2 - span) - o1) <= 2)
End Function

' the sheet's own picture of the same id goes to the template's place and size (its image is not touched); rows below
' lastRow are nShift lower on this sheet. Both sheets have their rows open. Returns the number moved.
Private Function MoveShapes(ws As Worksheet, chg As Collection, ByVal nShift As Long, ByVal lastRow As Long, ByRef miss As String) As Long
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
            If lastRow > 0 And r > lastRow Then r = r + nShift
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

' ==== data template to the data sheets ================================================================================
Private Function RunDataApply(ByVal doApply As Boolean, ByVal force As Boolean) As String
    Dim wb As Workbook, tpl As Worksheet, bs As Worksheet, nG As Long, dsA() As Worksheet, gsA() As Worksheet, ysA() As Worksheet
    Dim ds As Worksheet, nR As Long, nC As Long, t As Variant, b As Variant, g As Variant, k As Long, r As Long, c As Long
    Dim tv As String, gv As String, bv As String, sk() As Boolean, chg As Collection, upd() As Collection, con() As Collection
    Dim nUpd As Long, nCon As Long, it As Variant, rep As String, conList As String, writeCon As Boolean, cm As Long
    Dim ha As Variant, r1 As Long, r2 As Long, c1 As Long, c2 As Long, grown As Boolean, eTx As String
    Dim wasScripted As Boolean, nSh As Long, rect As String, rg As Range, oldF As Variant, oldV As Variant, wr() As Boolean
    Dim nKept As Long, cow As Boolean, cowSet As Boolean, cur As Variant, want As String
    On Error GoTo Fail
    wasScripted = scripted
    lg = ""
    Set wb = ThisWorkbook
    If Not SheetExists(wb, TPL_D) Then Err.Raise vbObjectError + 43, , U("\u015Eablon sayfa yok: ") & TPL_D
    Set tpl = wb.Worksheets(TPL_D)
    cm = Application.Calculation
    Application.ScreenUpdating = False
    LastCell tpl, nR, nC
    If Not SheetExists(wb, BASE_D) Then
        WriteBaseline wb, tpl, nR, nC, BASE_D
        Application.ScreenUpdating = True
        RunDataApply = U("\u0130z sayfas\u0131 (") & BASE_D & U(") yoktu: veri \u015Fablonunun bug\u00FCnk\u00FC h\u00E2li kaydedildi, hi\u00E7bir veri sayfas\u0131 de\u011Fi\u015Fmedi. Bundan sonra \u015Fablonda yapt\u0131\u011F\u0131n\u0131z de\u011Fi\u015Fiklikler bu tu\u015Fla veri sayfalar\u0131na ta\u015F\u0131n\u0131r.")
        Exit Function
    End If
    Set bs = wb.Worksheets(BASE_D)
    With bs.UsedRange
        If .Row + .Rows.Count - 1 > nR Then nR = .Row + .Rows.Count - 1
        If .Column + .Columns.Count - 1 > nC Then nC = .Column + .Columns.Count - 1
    End With
    If nR < 2 Then nR = 2
    If nC < 2 Then nC = 2
    nG = NameRange(wb, N_GTAB).Rows.Count
    FindSheets wb, nG, KIND_D, dsA
    FindSheets wb, nG, KIND_G, gsA
    FindSheets wb, nG, KIND_Y, ysA
    t = tpl.Range(tpl.Cells(1, 1), tpl.Cells(nR, nC)).Formula
    b = bs.Range(bs.Cells(1, 1), bs.Cells(nR, nC)).Value
    ' never compared: the areas the readers write (names on the template or on any data sheet), array formulas
    ReDim sk(1 To nR, 1 To nC)
    MarkSkipNames tpl, sk, nR, nC
    For k = 1 To nG
        If Not dsA(k) Is Nothing Then MarkSkipNames dsA(k), sk, nR, nC
    Next k
    ha = tpl.Range(tpl.Cells(1, 1), tpl.Cells(nR, nC)).HasArray
    If IsNull(ha) Then
        For r = 1 To nR
            For c = 1 To nC
                If Left$(CellText(t(r, c)), 1) = "=" Then
                    If tpl.Cells(r, c).HasArray Then sk(r, c) = True
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
            If Not sk(r, c) Then
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
            If Not dsA(k) Is Nothing Then GrowRect dsA(k), r1, r2, c1, c2, grown
        Next k
    Loop While grown
    rect = tpl.Range(tpl.Cells(r1, c1), tpl.Cells(r2, c2)).Address(False, False)

    ReDim upd(1 To nG): ReDim con(1 To nG)
    For k = 1 To nG
        Set upd(k) = New Collection: Set con(k) = New Collection
        If Not dsA(k) Is Nothing Then
            Set ds = dsA(k)
            nSh = nSh + 1
            g = ds.Range(ds.Cells(1, 1), ds.Cells(nR, nC)).Formula
            For Each it In chg
                r = it(0): c = it(1)
                tv = CellText(t(r, c)): bv = CellText(b(r, c)): gv = ToTemplate(CellText(g(r, c)), gsA(k), ysA(k), ds)
                If gv = bv Then
                    upd(k).Add Array(r, c)
                ElseIf gv <> tv Then
                    con(k).Add Array(r, c, CellText(g(r, c)))
                End If
            Next it
            nUpd = nUpd + upd(k).Count: nCon = nCon + con(k).Count
            rep = rep & "  " & ds.Name & ": " & upd(k).Count & U(" h\u00FCcre") & IIf(con(k).Count > 0, ", " & con(k).Count & U(" \u00E7ak\u0131\u015Fma"), "") & vbLf
            conList = conList & ConflictText(t, con(k), tpl, ds.Name, 0, 0, Len(conList))
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
            If Not dsA(k) Is Nothing Then
                Set ds = dsA(k)
                ' the template's rectangle over the data sheet's (cells, formats, merged cells; no clipboard), then every
                ' cell takes what it should hold: a written cell the template content in the group's terms, any other
                ' cell its own content back
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
                cur = rg.Formula
                For r = r1 To r2
                    For c = c1 To c2
                        If wr(r, c) Then
                            tv = CellText(t(r, c))
                            If Left$(tv, 1) = "=" Then
                                want = ToGroup(tv, gsA(k), ysA(k), ds)
                                If CellText(cur(r - r1 + 1, c - c1 + 1)) <> want Then
                                    On Error Resume Next
                                    ds.Cells(r, c).Formula = want
                                    On Error GoTo Fail
                                End If
                            End If
                        Else
                            gv = CellText(oldF(r - r1 + 1, c - c1 + 1))
                            If CellText(cur(r - r1 + 1, c - c1 + 1)) <> gv Then
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
        WriteBaseline wb, tpl, nR, nC, BASE_D
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

' ==== names and small helpers =========================================================================================
' number held by a sheet-level name (=n); 0 when the sheet has no such name
Private Function NameNumber(ws As Worksheet, ByVal nm As String) As Long
    Dim n As Name, v As String
    Set n = SheetNameObj(ws, nm)
    If n Is Nothing Then Exit Function
    v = Trim$(Mid$(n.RefersTo, 2))
    If Len(v) > 0 And Len(v) <= 6 Then
        If v Like String$(Len(v), "#") Then NameNumber = CLng(v)
    End If
End Function

Private Sub SetNameNumber(ws As Worksheet, ByVal nm As String, ByVal v As Long)
    Dim n As Name
    Set n = SheetNameObj(ws, nm)
    If n Is Nothing Then
        ws.Names.Add Name:=nm, RefersTo:="=" & v
    ElseIf n.RefersTo <> "=" & v Then
        n.RefersTo = "=" & v
    End If
End Sub

' the sheet-level name nm of ws (never a workbook-level one of the same name)
Private Function SheetNameObj(ws As Worksheet, ByVal nm As String) As Name
    Dim n As Name
    On Error Resume Next
    Set n = ws.Names(nm)
    On Error GoTo 0
    If n Is Nothing Then Exit Function
    If NameOnSheet(n, ws) Then Set SheetNameObj = n
End Function

' the cells of the sheet-level name nm of ws, Nothing when it is missing or shows another sheet
Private Function SheetRangeOpt(ws As Worksheet, ByVal nm As String) As Range
    Dim n As Name, r As Range
    Set n = SheetNameObj(ws, nm)
    If n Is Nothing Then Exit Function
    On Error Resume Next
    Set r = n.RefersToRange
    On Error GoTo 0
    If r Is Nothing Then Exit Function
    If r.Worksheet.Name <> ws.Name Then Exit Function
    Set SheetRangeOpt = r
End Function

Private Function SheetRange(ws As Worksheet, ByVal nm As String) As Range
    Set SheetRange = SheetRangeOpt(ws, nm)
    If SheetRange Is Nothing Then Err.Raise vbObjectError + 31, , ws.Name & U(": sayfa ad\u0131 yok veya bu sayfay\u0131 g\u00F6stermiyor: ") & nm
End Function

' the cells of the workbook-level name nm
Private Function NameRange(wb As Workbook, ByVal nm As String) As Range
    Dim n As Name
    Set n = WbName(wb, nm)
    If Not n Is Nothing Then
        On Error Resume Next
        Set NameRange = n.RefersToRange
        On Error GoTo 0
    End If
    If NameRange Is Nothing Then Err.Raise vbObjectError + 30, , U("Ad yok veya h\u00FCcre g\u00F6stermiyor: ") & nm
End Function

Private Function WbName(wb As Workbook, ByVal nm As String) As Name
    Dim n As Name
    On Error Resume Next
    Set n = wb.Names(nm)
    On Error GoTo 0
    If Not n Is Nothing Then
        If TypeOf n.Parent Is Workbook Then Set WbName = n: Exit Function
    End If
    For Each n In wb.Names
        If StrComp(n.Name, nm, vbTextCompare) = 0 Then
            If TypeOf n.Parent Is Workbook Then Set WbName = n: Exit Function
        End If
    Next n
End Function

Private Function NameOnSheet(nm As Name, ws As Worksheet) As Boolean
    On Error Resume Next
    NameOnSheet = (TypeOf nm.Parent Is Worksheet)
    If NameOnSheet Then NameOnSheet = (nm.Parent.Name = ws.Name)
    On Error GoTo 0
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

' report of a scripted run as UTF-8 (the test harness reads UTF-8); plain text when ADODB is not there
Private Sub WriteReport(ByVal path As String, ByVal s As String)
    Dim st As Object, f As Integer
    s = Replace(Replace(s, vbCrLf, vbLf), vbLf, vbCrLf)
    On Error Resume Next
    Set st = CreateObject("ADODB.Stream")
    If Not st Is Nothing Then
        st.Type = 2
        st.Charset = "utf-8"
        st.Open
        st.WriteText s
        st.SaveToFile path, 2
        st.Close
        If Err.Number = 0 Then Exit Sub
        Err.Clear
    End If
    On Error GoTo 0
    f = FreeFile
    Open path For Output As #f
    Print #f, s
    Close #f
End Sub

' The macro window (module PANEL) is optional: without it, a plain message box.
Private Sub ShowMsg(ByVal msg As String, ByVal icon As Long)
    Dim title As String
    If scripted Then Exit Sub
    title = U("Temel grubu sayfalar\u0131")
    On Error Resume Next
    Application.Run "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PANEL.UiMsg", msg, icon, title
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox msg, icon, title
    End If
End Sub

Private Function Ask(ByVal msg As String) As Long
    Dim title As String, r As Variant
    If scripted Then Ask = vbYes: Exit Function
    title = U("Temel grubu sayfalar\u0131")
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
