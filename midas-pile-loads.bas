Attribute VB_Name = "MIDAS_KAZIK"
Option Explicit
' ============================================================================
'  MIDAS_KAZIK - pile-head envelope loads of the abutment model open in Civil NX into KAZIK_UC of the
'  abutment workbook (8'6 KA1 / KA2, piled = KAZIKLI). Owner 2026-11-17 / 2026-11-19.
'
'  Macro window (MAKROLAR) > MIDAS sonuclari > Kazik uc yukleri, or MIDAS_KAZIK_OKU.
'  Reads the model OPEN in Civil NX through the MAPI; it writes nothing to the model.
'  Piles (hook PileHeads, owner 2026-11-19: no check against the KA_KAZIK_* inputs, the foundation
'  builder will be overhauled): the beam elements of structure group KAZIK (MIDAS_KENARAYAK) form one
'  chain per pile (J node of one = I node of the next); the head is the upper end of the chain
'  (I end of the top element for the KA builder). A new builder only needs a new PileHeads.
'  Pile choice (owner 2026-11-23: choose the piles, default the regular ones, addable): window field Kaziklar =
'  1 group KAZIK (default), 2 group KAZIK + added, 3 only the chosen ones. For 2 and 3 a list is asked: element
'  numbers as Civil NX writes them (2597 2617, 2597to3177by20) and structure-group names, separated by spaces or
'  commas. The box is filled with the elements selected in Civil NX (GET view/SELECT), or with the last list when
'  nothing is selected (it is kept in the source cell between square brackets, so the static and the seismic
'  model are read with the same piles). Connected line elements of the list are one chain = one pile, head = its
'  upper end; a single element is read at its upper end. Order: the regular piles by element, then the added ones.
'  Loads: envelopes ENV_SER, ENV_STR, ENV_DEP of db/LCOM-GEN, read as NAME(CB:max) and NAME(CB:min)
'  (a plain (CB) answers nothing, measured 2026-11-16), Beam Force table at the head end.
' Separate models (owner 2026-11-19, MIDAS_KENARAYAK model gate Statik / Deprem): only the envelopes the open
' model has are read (at least one); the columns and raw rows of the others are kept when the pile list
' (elements in order) is the one already in KAZIK_UC, else they are emptied; the stamp names what was read.
'  Sign (owner 2026-11-17): MIDAS max = tension side, min = compression side; report compression +:
'     Min = -Axial(ENV max), Mak = -Axial(ENV min), for Servis (ENV_SER), Dayanim (ENV_STR), Deprem (ENV_DEP).
'  KAZIK_UC (names, created at the defaults when missing):
'     M_KAZIK_DAMGA  stamp (model, date)        M_KAZIK_ADET  pile count     M_KAZIK_KAYNAK  pile-list source
'     M_KAZIK_TABLO  one row per pile, by element: No, Elem, node, x, y, Sira, Kolon, Servis Min / Mak,
'                    Dayanim Min / Mak, Deprem Min / Mak (values)
'     M_KAZIK_HAM    raw head forces as MIDAS gives them, 6 rows per pile (Elem, Load, Part, Axial,
'                    Shear-y, Shear-z, Torsion, Moment-y, Moment-z)
'  The report tables read row k of M_KAZIK_TABLO by formula and grow with the pile count
'  (KAZIK_SAYFALARI_YENILE_SESSIZ, run after every read):
'     2A_KAZIK per-pile table  36 piles per page, extra pages = copies of its base page (M_KAZIK_SAYFA1)
'     2A_KAZIK 8.6.7           45 piles per page, extra pages = copies of its base page (M_KAZIK_SAYFA2)
'                              without the heading rows; M_KAZIK_SAYFA3 = the page after them (8.6.8)
'     GEO_KAZIK                one row per pile (at least 35) in M_KAZIK_GEO, cells inserted / deleted
'  then the page numbers of 2A_KAZIK are chained again, the print areas set and the lists rebuilt
'  (RAPOR_LISTE); G_SON_NUMARA follows the last page, so the next workbook continues from it.
'  Second tool (owner 2026-11-23: the same for MIDAS_CIKTISI_KAZIK, let me pick the elements): MIDAS_KAZIK_GSD_OKU,
'  window > Kazik tesirleri (GSD). Reads the forces of chosen pile elements for every combination that ENV_SER,
'  ENV_STR and ENV_DEP of the open model are made of (load case or combination, as the envelope lists them) into
'  the GSD load block of MIDAS_CIKTISI_KAZIK: M_TESIR_BLOK (CW26:DE3839) = counter, No, name
'  '<combination>_<element>_<I|J>', Pu = -Axial (compression +), My, Mz, Vy, Vz, T (kN, kN.m) as values; formulas
'  of the block are replaced, the pasted table right of it is not read and not touched. Elements: 1 pile heads of
'  group KAZIK, 2 pile heads + the list, 3 only the list; here every listed element is read by itself, at its upper
'  end. Static (ENV_SER, ENV_STR) and seismic (ENV_DEP) models are read one after the other, in any order: the rows
'  of the other kind stay (seismic = names that start with TKDY- or H-TKDY-), static rows first. Stamps M_TESIR_STATIK
'  / M_TESIR_DEPREM (DG2, DG3), element list M_TESIR_SECIM (DG4, used as the default of the next read).
'  Foundation workbook (8'8, 2026-11-23): the report sheet is 2_KAZIK (no 8.6.7 page; M_KAZIK_SAYFA1 = per-pile page,
'  M_KAZIK_SAYFA3 = the page after it), GEO_KAZIK has K:Q only (M_KAZIK_GEO).
'  MAPI key: read by MIDAS_CORE (G_MAPI_KEY, INPUT!T3, MIDAS_MAPI_KEY, asked; memory only).
'  Test entries: MIDAS_KAZIK_RUN <report file> [, 1 / 2 / 3, list] (no questions), MIDAS_KAZIK_PLAN [1 / 2 / 3,
'  list] (piles and envelopes of the open model; nothing is read from the result tables or written),
'  MIDAS_KAZIK_PLANYAZ <report file> [, 1 / 2 / 3, list] (the plan into a file), MIDAS_KAZIK_GSD_RUN and
'  MIDAS_KAZIK_GSD_PLANYAZ <report file> [, 1 / 2 / 3, list] (the same for the GSD block), MIDAS_KAZIK_TEST_YAZ n,
'  <report file> (n made-up piles into KAZIK_UC without a model, then the pages; for the growth tests).
' ============================================================================

Private Const SCRIPT_ID As String = "midas-pile-loads"
Private Const SCRIPT_VERSION As String = "2026-11-23c"
Private Const SCRIPT_CHANGELOG As String = "Foundation workbook (8'8) too: KAZIK_UC with 2_KAZIK and GEO_KAZIK; New tool: pile forces for GSD (every combination of the envelopes, chosen elements) into the load block of MIDAS_CIKTISI_KAZIK; Choice of the piles: group KAZIK (default), group KAZIK + added elements, or only chosen elements (typed as in Civil NX, or the elements selected there); Reads the envelopes the open model has (static or seismic abutment model of the model gate) and keeps the others from the last read when the piles are the same; New: pile-head envelope loads (ENV_SER / ENV_STR / ENV_DEP, max and min) of the open abutment model into KAZIK_UC; the 2A_KAZIK per-pile table, section 8.6.7 and GEO_KAZIK grow with the pile count, page numbers and lists follow"
Private Const SCRIPT_NEEDS As String = "midas-core"   ' UPDATE adds this module when the workbook lacks it
' macro-window row(s) of this module (PANEL reads them): order|group|title|macro|sheet|text|fields[|modules]
'@TOOL 84|MIDAS sonu\u00E7lar\u0131|Kaz\u0131k u\u00E7 y\u00FCkleri|MIDAS_KAZIK_OKU||Civil NX'te a\u00E7\u0131k kenar ayak modelinden kaz\u0131k ba\u015Flar\u0131n\u0131n zarf y\u00FCklerini (ENV_SER / ENV_STR / ENV_DEP, en b\u00FCy\u00FCk ve en k\u00FC\u00E7\u00FCk) KAZIK_UC sayfas\u0131na okur; rapor tablolar\u0131 (2A_KAZIK, 8.6.7, GEO_KAZIK) kaz\u0131k say\u0131s\u0131na g\u00F6re b\u00FCy\u00FCr. Kaz\u0131klar alan\u0131: KAZIK grubu, gruba ek elemanlar ya da yaln\u0131z se\u00E7ilen elemanlar; liste, Civil NX'te se\u00E7ili elemanlarla dolu sorulur. Modele bir \u015Fey yazmaz.|kazik~Kaz\u0131klar~c~KAZIK grubu/KAZIK grubu + ek/Yaln\u0131z se\u00E7ilen
'@TOOL 85|MIDAS sonu\u00E7lar\u0131|Kaz\u0131k tesirleri (GSD)|MIDAS_KAZIK_GSD_OKU||Civil NX'te a\u00E7\u0131k kenar ayak modelinden kaz\u0131k elemanlar\u0131n\u0131n tesirlerini (ENV_SER / ENV_STR / ENV_DEP zarflar\u0131n\u0131n kombinasyonlar\u0131, tek tek) MIDAS_CIKTISI_KAZIK sayfas\u0131ndaki GSD y\u00FCk blo\u011Funa okur: ad, Pu, My, Mz, Vy, Vz, T. Statik ve deprem modeli ayr\u0131 ayr\u0131 okunur, di\u011Ferinin sat\u0131rlar\u0131 kal\u0131r. Elemanlar: kaz\u0131k ba\u015Flar\u0131, ba\u015Flar + ek ya da yaln\u0131z se\u00E7ilen; liste, Civil NX'te se\u00E7ili elemanlarla dolu sorulur. Modele bir \u015Fey yazmaz.|kazik~Elemanlar~c~Kaz\u0131k ba\u015Flar\u0131/Kaz\u0131k ba\u015Flar\u0131 + ek/Yaln\u0131z se\u00E7ilen

Private Const MSG_TITLE As String = "MIDAS_KAZIK"
Private Const SH_UC As String = "KAZIK_UC"
Private Const SH_GSD As String = "MIDAS_CIKTISI_KAZIK"   ' GSD data sheet of the piles
Private Const SH_REP As String = "2A_KAZIK"
Private Const SH_REP2 As String = "2_KAZIK"               ' foundation workbook (8'8): no 8.6.7 page
Private Const SH_GEO As String = "GEO_KAZIK"
Private Const PILE_GROUP As String = "KAZIK"
Private Const MODE_GROUP As Long = 1       ' piles of group KAZIK
Private Const MODE_ADD As Long = 2         ' group KAZIK + the list
Private Const MODE_ONLY As Long = 3        ' only the list
Private Const ENV_LIST As String = "ENV_SER|ENV_STR|ENV_DEP"
Private Const MAX_PILES As Long = 1000
Private Const PG As Long = 46              ' rows per report page (printed 45 + 1 blank)
Private Const PG_FIRST As Long = 2         ' first row of page 1
Private Const PER1 As Long = 36            ' piles per page, 2A_KAZIK per-pile table
Private Const PER2 As Long = 45            ' piles per page, 8.6.7 (3 blocks of 15)
Private Const GEO_MIN As Long = 35         ' rows of the GEO_KAZIK table at least
Private Const COL_W As Long = 23           ' W: page number, row top + 43
Private Const TYPE_NAMES As String = "KA_TEMEL_TIPI|T_TEMEL_TIPI"

Private mScripted As Boolean
Private mLog As String

' ============================================================================
Public Sub MIDAS_KAZIK_OKU()
    Dim s As String
    mScripted = False
    s = RunRead(False, UiIdx("kazik", MODE_GROUP), "")
    UiMsg s, IIf(Left$(s, 8) = "STOPPED:", vbExclamation, vbInformation), MSG_TITLE
End Sub

Public Sub MIDAS_KAZIK_RUN(ByVal reportFile As String, Optional ByVal pileMode As Long = 1, Optional ByVal pileList As String = "")
    mScripted = True
    WriteText reportFile, RunRead(False, pileMode, pileList)
End Sub

Public Function MIDAS_KAZIK_PLAN(Optional ByVal pileMode As Long = 1, Optional ByVal pileList As String = "") As String
    mScripted = True
    MIDAS_KAZIK_PLAN = RunRead(True, pileMode, pileList)
End Function

Public Sub MIDAS_KAZIK_PLANYAZ(ByVal reportFile As String, Optional ByVal pileMode As Long = 1, Optional ByVal pileList As String = "")
    WriteText reportFile, MIDAS_KAZIK_PLAN(pileMode, pileList)
End Sub

' n made-up piles (elements 14000.., 4 per row, 3 m spacing) into KAZIK_UC, then the pages; for tests
Public Sub MIDAS_KAZIK_TEST_YAZ(ByVal n As Long, ByVal reportFile As String)
    Dim heads As Collection, res As Object, i As Long, e As Long, s As String, env As Variant, j As Long
    On Error GoTo Fail
    mScripted = True: mLog = ""
    If n < 1 Or n > MAX_PILES Then Err.Raise vbObjectError + 1, , "n must be 1.." & MAX_PILES & "."
    Set heads = New Collection
    Set res = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        e = 13999 + i
        heads.Add Array(e, e, ((i - 1) Mod 4) * 3#, ((i - 1) \ 4) * 3#, "I")
        j = 0
        For Each env In Split(ENV_LIST, "|")
            j = j + 1
            res(e & "|" & env & "|max") = Array(-(j * 1000 + i), 1, 2, 3, 4, 5, "I[" & e & "]")
            res(e & "|" & env & "|min") = Array(-(j * 1000 + 500 + i), 1, 2, 3, 4, 5, "I[" & e & "]")
        Next env
    Next i
    s = WriteUc(ThisWorkbook, heads, res, "TEST " & n & " piles | " & Format$(Now, "yyyy-mm-dd hh:nn"), "test (MIDAS_KAZIK_TEST_YAZ)")
    s = s & PagesRefresh(ThisWorkbook)
    WriteText reportFile, "Done. " & s & mLog
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    WriteText reportFile, "STOPPED: " & Err.Description & IIf(Len(mLog) > 0, vbCrLf & mLog, "")
End Sub

' report pages (2A_KAZIK, GEO_KAZIK) for the pile count in KAZIK_UC; "" or a log, "FAIL:..." on an error
Public Function KAZIK_SAYFALARI_YENILE_SESSIZ() As String
    On Error GoTo Fail
    KAZIK_SAYFALARI_YENILE_SESSIZ = PagesRefresh(ThisWorkbook)
    Exit Function
Fail:
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    KAZIK_SAYFALARI_YENILE_SESSIZ = "FAIL:" & Err.Description
End Function

' ============================================================================
Private Function RunRead(ByVal planOnly As Boolean, ByVal pileMode As Long, ByVal pileList As String) As String
    Dim wb As Workbook, pj As String, tag As String, t As String, lc As Object, heads As Collection, src As String
    Dim keys As String, h As Variant, res As Object, s As String, k As Variant, envs As Object, miss As String, envUse As String, ans As Variant
    On Error GoTo Fail
    mLog = ""
    Set wb = ThisWorkbook

    ' ---- the workbook ----
    t = TemelTipi(wb)
    If Len(t) > 0 And t <> "KAZIKLI" Then RunRead = "STOPPED: the foundation type is " & t & ", not KAZIKLI: the pile sheets are hidden, nothing was read.": Exit Function
    If SheetOf(wb, SH_UC) Is Nothing Then RunRead = "STOPPED: this workbook has no sheet " & SH_UC & ".": Exit Function
    CheckUcLayout wb.Worksheets(SH_UC)

    ' ---- the model ----
    If Not McGetKey(wb, Not mScripted, MSG_TITLE) Then RunRead = "STOPPED: no MAPI key: nothing was read.": Exit Function
    Logp "Civil NX: model"
    If Not Api("GET", "db/PJCF", "", pj) Then
        If Left$(pj, 10) = "HTTP error" Or pj = "" Then RunRead = "STOPPED: Civil NX does not answer (is it open, is the MAPI connected?). " & Left$(pj, 200): Exit Function
        pj = ""
    End If
    tag = ModelTag(pj)
    Logp "Civil NX: combinations"
    Set lc = TableGet("LCOM-GEN")
    If lc Is Nothing Then Err.Raise vbObjectError + 4, , "The open model (" & tag & ") has no load combinations (db/LCOM-GEN)."
    Set envs = CreateObject("Scripting.Dictionary")
    For Each k In lc.keys
        envs(UCase$(CStr(lc(k)("NAME")))) = True
    Next k
    For Each k In Split(ENV_LIST, "|")
        If envs.Exists(CStr(k)) Then
            envUse = envUse & IIf(envUse = "", "", "|") & k
        Else
            miss = miss & IIf(miss = "", "", ", ") & k
        End If
    Next k
    If Len(envUse) = 0 Then Err.Raise vbObjectError + 5, , "The open model (" & tag & ") has none of the envelopes " & Replace(ENV_LIST, "|", ", ") & ": open the abutment model built by MIDAS_KENARAYAK."
    If UiStopped() Then RunRead = "STOPPED: by the user.": Exit Function
    If pileMode < MODE_GROUP Or pileMode > MODE_ONLY Then pileMode = MODE_GROUP
    If pileMode <> MODE_GROUP And Not mScripted Then
        Logp "Civil NX: selected elements"
        ans = AskPiles(CellTxt(McAnchor(wb.Worksheets(SH_UC), "M_KAZIK_KAYNAK", "B4").Value), pileMode, False)
        If VarType(ans) = vbBoolean Then RunRead = "STOPPED: nothing was read.": Exit Function
        pileList = CStr(ans)
    End If
    Logp "Civil NX: piles (" & IIf(pileMode = MODE_ONLY, "", PILE_GROUP & ", ") & "elements, nodes)"
    Set heads = PileHeads(src, pileMode, pileList)
    If heads.Count > MAX_PILES Then Err.Raise vbObjectError + 6, , heads.Count & " piles, KAZIK_UC holds " & MAX_PILES & "."
    For Each h In heads
        keys = keys & IIf(keys = "", "", ",") & h(0)
    Next h

    If planOnly Then
        s = "PLAN. Model: " & tag & vbCrLf & "Envelopes: " & Replace(envUse, "|", ", ") & IIf(miss = "", "", " (not in this model: " & miss & ")") & vbCrLf & "Piles: " & heads.Count & " (" & src & ")" & vbCrLf
        For Each h In heads
            s = s & "  " & h(0) & " " & h(4) & "-end node " & h(1) & " x " & Format$(h(2), "0.00") & " y " & Format$(h(3), "0.00") & vbCrLf
        Next h
        RunRead = s
        Exit Function
    End If

    If Not mScripted Then
        s = "Model: " & tag & vbCrLf & "Piles: " & heads.Count & " (" & src & ")" & vbCrLf & "Envelopes: " & Replace(envUse, "|", ", ") & " (max and min)" & vbCrLf & _
            IIf(miss = "", "", "Not in this model: " & miss & IIf(UcSame(wb.Worksheets(SH_UC), heads), " - their columns are kept from the last read (same piles).", " - WARNING: the piles differ from the last read, their columns will be EMPTY until that model is read with the same piles.") & vbCrLf) & vbCrLf & _
            "The pile-head loads of this model replace the table in " & SH_UC & "; the report tables follow the pile count. Continue?"
        If UiMsg(s, vbOKCancel + vbQuestion, MSG_TITLE) <> vbOK Then RunRead = "STOPPED: nothing was read.": Exit Function
    End If

    ' ---- forces ----
    Set res = CreateObject("Scripting.Dictionary")
    Logp "Civil NX: pile-head forces (max)"
    ReadTable keys, "max", res, envUse
    Logp "Civil NX: pile-head forces (min)"
    ReadTable keys, "min", res, envUse
    If UiStopped() Then RunRead = "STOPPED: by the user (nothing was written).": Exit Function

    ' ---- write ----
    Logp "Workbook: " & SH_UC & " and the report pages"
    s = WriteUc(wb, heads, res, tag & " | " & Format$(Now, "yyyy-mm-dd hh:nn"), src, envUse)
    s = s & PagesRefresh(wb)
    RunRead = "Done. Model: " & tag & vbCrLf & s & mLog
    Exit Function
Fail:
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    RunRead = "STOPPED: " & Err.Description & IIf(Len(mLog) > 0, vbCrLf & mLog, "")
End Function

' ---- hook: the pile heads of the open model --------------------------------------------------
' items Array(headElem, headNode, x, y, "I" | "J"): the piles of group KAZIK by element, then the added ones by
' element; src = where the list came from. Regular piles: the beam elements of structure group KAZIK
' (MIDAS_KENARAYAK). A new foundation builder replaces only the list of regular elements.
Private Function PileHeads(ByRef src As String, Optional ByVal pileMode As Long = 1, Optional ByVal pileList As String = "") As Collection
    Dim gr As Object, k As Variant, el As Variant, ev As Object, nd As Object, reg As Collection, extra As Collection, out As Collection
    Dim seen As Object, h As Variant, n As Long, els As Object, e As Variant
    src = ""
    Set gr = TableGet("GRUP")
    Set ev = TableGet("ELEM")
    Set nd = TableGet("NODE")
    If ev Is Nothing Or nd Is Nothing Then Err.Raise vbObjectError + 7, , "Elements or nodes could not be read."
    Set out = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    If pileMode <> MODE_ONLY Then
        If gr Is Nothing Then Err.Raise vbObjectError + 7, , "The open model has no structure groups (db/GRUP)."
        For Each k In gr.keys
            If UCase$(CStr(gr(k)("NAME"))) = PILE_GROUP Then
                If gr(k).Exists("E_LIST") Then el = gr(k)("E_LIST")
                Exit For
            End If
        Next k
        If IsEmpty(el) Then Err.Raise vbObjectError + 7, , "Structure group " & PILE_GROUP & " not found or without elements: not a piled abutment model of MIDAS_KENARAYAK (choose the third entry of the window field to read chosen elements)."
        Set els = CreateObject("Scripting.Dictionary")
        For Each e In el
            els(CStr(CLng(e))) = True
        Next e
        Set reg = ChainHeads(els, ev, nd, PILE_GROUP, False)
        For Each h In reg
            seen(h(0) & h(4)) = True
            out.Add h
        Next h
        src = "structure group " & PILE_GROUP & ", " & reg.Count & " element chains"
    End If
    If pileMode <> MODE_GROUP Then
        Set els = ListElems(pileList, gr, ev)
        If els.Count > 0 Then
            Set extra = ChainHeads(els, ev, nd, "the list", True)
            For Each h In extra
                If Not seen.Exists(h(0) & h(4)) Then
                    seen(h(0) & h(4)) = True
                    out.Add h
                    n = n + 1
                End If
            Next h
            src = src & IIf(src = "", "chosen elements [", " + added elements [") & ShortList(els) & "], " & n & " chains"
            If n < extra.Count Then src = src & " (" & (extra.Count - n) & " already there)"
        ElseIf pileMode = MODE_ONLY Then
            Err.Raise vbObjectError + 7, , "No elements chosen: nothing was read."
        End If
    End If
    If out.Count = 0 Then Err.Raise vbObjectError + 7, , "No piles found."
    src = src & ", head = upper end"
    Set PileHeads = out
End Function

' pile heads of a set of elements (keys = element numbers): connected line elements are one chain, head = its
' upper end; sorted by element. strict: an element that is not a line element stops the read.
Private Function ChainHeads(els As Object, ev As Object, nd As Object, ByVal what As String, ByVal strict As Boolean) As Collection
    Dim inSet As Object, byI As Object, jSet As Object, e As Variant, k As Variant, rec As Object, tops() As Long
    Dim n As Long, i As Long, cur As Long, nI As Long, nJ As Long, first As Long, steps As Long
    Dim zt As Double, zb As Double, out As Collection, nodes As Variant, isLine As Boolean
    Set inSet = CreateObject("Scripting.Dictionary")
    Set byI = CreateObject("Scripting.Dictionary")
    Set jSet = CreateObject("Scripting.Dictionary")
    For Each e In els.keys
        If Not ev.Exists(CStr(e)) Then Err.Raise vbObjectError + 7, , "Element " & e & " of " & what & " is not in the model."
        Set rec = ev(CStr(e))
        nodes = rec("NODE")
        nI = CLng(NumOr0(nodes(1))): nJ = CLng(NumOr0(nodes(2)))
        isLine = (nI > 0 And nJ > 0)
        If isLine Then If UBound(nodes) >= 3 Then isLine = (CLng(NumOr0(nodes(3))) = 0)
        If isLine Then
            inSet(CStr(e)) = Array(nI, nJ)
            If byI.Exists(CStr(nI)) Then Err.Raise vbObjectError + 7, , "Two pile elements start at node " & nI & ": the piles cannot be followed."
            byI(CStr(nI)) = CLng(e)
            jSet(CStr(nJ)) = True
        ElseIf strict Then
            Err.Raise vbObjectError + 7, , "Element " & e & " of " & what & " is not a line element (pile)."
        End If
    Next e
    If inSet.Count = 0 Then Err.Raise vbObjectError + 7, , IIf(what = PILE_GROUP, "Structure group " & PILE_GROUP, "The list") & " has no line elements."
    ReDim tops(1 To inSet.Count)
    For Each k In inSet.keys
        If Not jSet.Exists(CStr(inSet(k)(0))) Then n = n + 1: tops(n) = CLng(k)
    Next k
    If n = 0 Then Err.Raise vbObjectError + 7, , "No pile end found in " & what & " (the elements form loops)."
    ReDim Preserve tops(1 To n)
    SortLongs tops
    Set out = New Collection
    For i = 1 To n
        first = tops(i): cur = first: steps = 0
        Do
            nJ = inSet(CStr(cur))(1)
            If Not byI.Exists(CStr(nJ)) Then Exit Do
            cur = byI(CStr(nJ))
            steps = steps + 1
            If steps > inSet.Count Then Err.Raise vbObjectError + 7, , "The pile elements from " & first & " form a loop."
        Loop
        zt = NodeC(nd, inSet(CStr(first))(0), "Z")
        zb = NodeC(nd, inSet(CStr(cur))(1), "Z")
        If zt >= zb Then
            nI = inSet(CStr(first))(0)
            out.Add Array(first, nI, NodeC(nd, nI, "X"), NodeC(nd, nI, "Y"), "I")
        Else
            nJ = inSet(CStr(cur))(1)
            out.Add Array(cur, nJ, NodeC(nd, nJ, "X"), NodeC(nd, nJ, "Y"), "J")
        End If
    Next i
    Set ChainHeads = SortByElem(out)
End Function

' the list of elements typed by the user: numbers, a to b [by c] as Civil NX writes them (2597to3177by20),
' structure-group names; separated by spaces, commas, semicolons or line ends. Keys = element numbers.
Private Function ListElems(ByVal txt As String, gr As Object, ev As Object) As Object
    Dim d As Object, part As Variant, tok As Variant, p As String
    Set d = CreateObject("Scripting.Dictionary")
    txt = Replace(Replace(Replace(Replace(txt, vbCr, ","), vbLf, ","), ";", ","), vbTab, " ")
    For Each part In Split(txt, ",")
        p = Trim$(CStr(part))
        If Len(p) > 0 Then
            If Not AddGroup(p, gr, d) Then                      ' a group name may hold spaces
                For Each tok In Split(p, " ")
                    If Len(Trim$(CStr(tok))) > 0 Then AddToken Trim$(CStr(tok)), gr, ev, d
                Next tok
            End If
        End If
    Next part
    Set ListElems = d
End Function

Private Function AddGroup(ByVal nm As String, gr As Object, d As Object) As Boolean
    Dim k As Variant, e As Variant, el As Variant
    If gr Is Nothing Then Exit Function
    For Each k In gr.keys
        If StrComp(CStr(gr(k)("NAME")), nm, vbTextCompare) = 0 Then
            If gr(k).Exists("E_LIST") Then el = gr(k)("E_LIST")
            If Not IsArray(el) Then Err.Raise vbObjectError + 7, , "Structure group " & nm & " has no elements."
            For Each e In el
                d(CStr(CLng(e))) = True
            Next e
            AddGroup = True
            Exit Function
        End If
    Next k
End Function

Private Sub AddToken(ByVal tok As String, gr As Object, ev As Object, d As Object)
    Dim s As String, p As Long, q As Long, a As Long, b As Long, c As Long, i As Long, n As Long
    s = LCase$(tok)
    If AllDigits(s) Then
        If Not ev.Exists(CStr(CLng(s))) Then Err.Raise vbObjectError + 7, , "Element " & s & " is not in the model."
        d(CStr(CLng(s))) = True
        Exit Sub
    End If
    p = InStr(s, "to")
    If p > 1 Then
        q = InStr(p + 2, s, "by")
        If q = 0 Then q = Len(s) + 1
        If AllDigits(Left$(s, p - 1)) And AllDigits(Mid$(s, p + 2, q - p - 2)) And (q > Len(s) Or AllDigits(Mid$(s, q + 2))) Then
            a = CLng(Left$(s, p - 1)): b = CLng(Mid$(s, p + 2, q - p - 2)): c = 1
            If q <= Len(s) Then c = CLng(Mid$(s, q + 2))
            If c < 1 Or b < a Then Err.Raise vbObjectError + 7, , "The range " & tok & " cannot be read."
            If (b - a) \ c > 100000 Then Err.Raise vbObjectError + 7, , "The range " & tok & " is too long."
            For i = a To b Step c
                If ev.Exists(CStr(i)) Then d(CStr(i)) = True: n = n + 1
            Next i
            If n = 0 Then Err.Raise vbObjectError + 7, , "No element of " & tok & " is in the model."
            Exit Sub
        End If
    End If
    If AddGroup(tok, gr, d) Then Exit Sub
    Err.Raise vbObjectError + 7, , "'" & tok & "' is neither an element number, a range (2597to3177by20) nor a structure group of the open model."
End Sub

Private Function AllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Or Len(s) > 9 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    AllDigits = True
End Function

' element numbers (keys of d) as Civil NX writes them: 2597to3177by20 3200 3205
Private Function ShortList(d As Object) As String
    Dim a() As Long, k As Variant, n As Long, i As Long, j As Long, st As Long, s As String
    If d.Count = 0 Then Exit Function
    ReDim a(1 To d.Count)
    For Each k In d.keys
        n = n + 1: a(n) = CLng(k)
    Next k
    SortLongs a
    i = 1
    Do While i <= n
        j = i
        If i + 2 <= n Then
            st = a(i + 1) - a(i)
            If a(i + 2) - a(i + 1) = st Then
                j = i + 2
                Do While j < n
                    If a(j + 1) - a(j) <> st Then Exit Do
                    j = j + 1
                Loop
            End If
        End If
        If j > i Then
            s = s & IIf(s = "", "", " ") & a(i) & "to" & a(j) & IIf(st = 1, "", "by" & st)
        Else
            s = s & IIf(s = "", "", " ") & a(i)
        End If
        i = j + 1
    Loop
    ShortList = s
End Function

' the list for modes 2 and 3: the box holds the elements selected in Civil NX, else the list of the last read
' (between square brackets in the cell 'last'); False on Cancel
Private Function AskPiles(ByVal last As String, ByVal pileMode As Long, ByVal forGsd As Boolean) As Variant
    Dim def As String, note As String, p As Long, q As Long, head As String, rule As String
    def = SelectedInCivil()
    If Len(def) > 0 Then
        note = U("Kutu, Civil NX'te se\u00E7ili elemanlarla dolduruldu.")
    Else
        p = InStr(last, "[")
        If p > 0 Then q = InStr(p, last, "]")
        If q > p + 1 Then
            def = Mid$(last, p + 1, q - p - 1)
            note = U("Civil NX'te se\u00E7ili eleman yok; kutuda son okuman\u0131n listesi var.")
        Else
            note = U("Civil NX'te se\u00E7ili eleman yok: elemanlar\u0131 orada se\u00E7ip yeniden a\u00E7\u0131n ya da listeyi yaz\u0131n.")
        End If
    End If
    If forGsd Then
        head = IIf(pileMode = MODE_ADD, U("Kaz\u0131k ba\u015Flar\u0131na eklenecek elemanlar"), U("Okunacak elemanlar (yaln\u0131z bunlar)"))
        rule = U("Her eleman tek ba\u015F\u0131na, \u00FCst ucunda okunur.")
    Else
        head = IIf(pileMode = MODE_ADD, U("KAZIK grubuna eklenecek kaz\u0131k elemanlar\u0131"), U("Okunacak kaz\u0131k elemanlar\u0131 (yaln\u0131z bunlar)"))
        rule = U("Birbirine ba\u011Fl\u0131 elemanlar tek kaz\u0131kt\u0131r ve \u00FCst ucu okunur; tek eleman kendi \u00FCst ucunda okunur.")
    End If
    AskPiles = UiInput(head & vbCrLf & U("Eleman no: 2597 2617 ya da 2597to3177by20; yap\u0131 grubu ad\u0131 da yaz\u0131labilir (bo\u015Fluk ya da virg\u00FClle).") & vbCrLf & _
        rule & vbCrLf & note, MSG_TITLE, def)
End Function

' the elements selected in Civil NX, as a short list; empty when none (or when the program does not tell)
Private Function SelectedInCivil() As String
    Dim resp As String, root As Object, el As Variant, e As Variant, d As Object
    On Error GoTo Quit
    If Not Api("GET", "view/SELECT", "", resp) Then Exit Function
    Set root = ParseJson(resp)
    If root Is Nothing Then Exit Function
    If Not root.Exists("SELECT") Then Exit Function
    If Not root("SELECT").Exists("ELEM_LIST") Then Exit Function
    el = root("SELECT")("ELEM_LIST")
    If Not IsArray(el) Then Exit Function
    Set d = CreateObject("Scripting.Dictionary")
    For Each e In el
        d(CStr(CLng(NumOr0(e)))) = True
    Next e
    SelectedInCivil = ShortList(d)
Quit:
End Function

' Beam Force table of the pile heads for the three envelopes, side "max" or "min":
' res("<elem>|<ENV>|<side>") = Array(Axial, Shear-y, Shear-z, Torsion, Moment-y, Moment-z, Part text) as MIDAS gives
Private Sub ReadTable(ByVal keys As String, ByVal side As String, res As Object, ByVal envUse As String)
    Dim nm As Variant, lcs As String, body As String, resp As String, root As Object, tb As Object, head As Variant, data As Variant
    Dim iE As Long, iL As Long, iP As Long, iA As Long, iVy As Long, iVz As Long, iT As Long, iMy As Long, iMz As Long
    Dim j As Long, r As Long, ld As String, want As Object
    Set want = CreateObject("Scripting.Dictionary")
    For Each nm In Split(envUse, "|")
        lcs = lcs & IIf(lcs = "", "", ",") & QS(CStr(nm) & "(CB:" & side & ")")
        want(CStr(nm)) = True
    Next nm
    body = "{""Argument"":{""TABLE_NAME"":""F"",""TABLE_TYPE"":""BEAMFORCE"",""UNIT"":{""FORCE"":""KN"",""DIST"":""M""}," & _
           """NODE_ELEMS"":{""KEYS"":[" & keys & "]},""LOAD_CASE_NAMES"":[" & lcs & "],""PARTS"":[""PartI"",""PartJ""]}}"
    If Not Api("POST", "post/TABLE", body, resp) Then Err.Raise vbObjectError + 10, , "Beam forces (" & side & "): " & Left$(resp, 300)
    Set root = ParseJson(resp)
    If root Is Nothing Then Err.Raise vbObjectError + 10, , "Beam forces (" & side & "): no answer."
    If Not root.Exists("F") Then Err.Raise vbObjectError + 10, , "Beam forces (" & side & "): no table (has the analysis been run?). " & Left$(resp, 200)
    Set tb = root("F")
    head = tb("HEAD"): data = tb("DATA")
    For j = 1 To UBound(head)
        Select Case CStr(head(j))
            Case "Elem": iE = j
            Case "Load": iL = j
            Case "Part": iP = j
            Case "Axial": iA = j
            Case "Shear-y": iVy = j
            Case "Shear-z": iVz = j
            Case "Torsion": iT = j
            Case "Moment-y": iMy = j
            Case "Moment-z": iMz = j
        End Select
    Next j
    If iE * iL * iP * iA * iVy * iVz * iT * iMy * iMz = 0 Then Err.Raise vbObjectError + 10, , "Beam forces (" & side & "): unexpected columns."
    If IsEmpty(data) Then Err.Raise vbObjectError + 10, , "Beam forces (" & side & "): empty table (has the analysis been run?)."
    For r = 1 To UBound(data)
        ld = UCase$(BaseName(CStr(data(r)(iL))))
        If want.Exists(ld) Then
            res(CStr(CLng(Val(CStr(data(r)(iE))))) & "|" & ld & "|" & side & "|" & UCase$(Left$(CStr(data(r)(iP)), 1))) = Array( _
                Val(CStr(data(r)(iA))), Val(CStr(data(r)(iVy))), Val(CStr(data(r)(iVz))), Val(CStr(data(r)(iT))), _
                Val(CStr(data(r)(iMy))), Val(CStr(data(r)(iMz))), CStr(data(r)(iP)))
        End If
    Next r
End Sub

' the force of a pile head at its end; the test writer stores them without the end letter
Private Function ForceOf(res As Object, h As Variant, ByVal env As String, ByVal side As String) As Variant
    Dim key As String
    key = h(0) & "|" & env & "|" & side
    If res.Exists(key & "|" & h(4)) Then
        ForceOf = res(key & "|" & h(4))
    ElseIf res.Exists(key) Then
        ForceOf = res(key)
    Else
        Err.Raise vbObjectError + 12, , "No force for " & env & "(" & side & ") at the " & h(4) & " end of element " & h(0) & "."
    End If
End Function

' ---- KAZIK_UC -------------------------------------------------------------------------------
Private Sub CheckUcLayout(uc As Worksheet)
    Dim tbl As Range, raw As Range
    Set tbl = McAnchor(uc, "M_KAZIK_TABLO", "A8:M1007")
    Set raw = McAnchor(uc, "M_KAZIK_HAM", "O8:W6007")
    If tbl.Columns.Count <> 13 Or tbl.Rows.Count < MAX_PILES Then Err.Raise vbObjectError + 11, , "Name M_KAZIK_TABLO must be 13 columns and " & MAX_PILES & " rows."
    If raw.Columns.Count <> 9 Or raw.Rows.Count < 6 * MAX_PILES Then Err.Raise vbObjectError + 11, , "Name M_KAZIK_HAM must be 9 columns and " & 6 * MAX_PILES & " rows."
    If tbl.Row < 2 Or raw.Row < 2 Then Err.Raise vbObjectError + 11, , SH_UC & ": the tables need their header row above them."
    If StrComp(CellTxt(tbl.Cells(0, 2).Value), "Elem", vbTextCompare) <> 0 Or StrComp(CellTxt(raw.Cells(0, 1).Value), "Elem", vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 11, , SH_UC & " does not have the MIDAS_KAZIK layout (header 'Elem' above the pile table and the raw table): install the new " & SH_UC & " first."
End Sub

' True when KAZIK_UC holds these piles (elements in the same order)
Private Function UcSame(uc As Worksheet, heads As Collection) As Boolean
    Dim tbl As Range, v As Variant, h As Variant, i As Long
    Set tbl = McAnchor(uc, "M_KAZIK_TABLO", "A8:M1007")
    If CLng(NumOr0(McAnchor(uc, "M_KAZIK_ADET", "B3").Value)) <> heads.Count Then Exit Function
    v = tbl.Resize(heads.Count, 2).Value
    For Each h In heads
        i = i + 1
        If CLng(NumOr0(v(i, 2))) <> CLng(h(0)) Then Exit Function
    Next h
    UcSame = True
End Function

' writes the pile table, the raw head forces, the stamp, the count and the source; returns a summary line.
' envUse = the envelopes read ("" = all); the others keep their cells when the pile list is the one in KAZIK_UC
Private Function WriteUc(wb As Workbook, heads As Collection, res As Object, ByVal stamp As String, ByVal src As String, Optional ByVal envUse As String = "") As String
    Dim uc As Worksheet, tbl As Range, raw As Range, n As Long, out() As Variant, rw() As Variant, h As Variant, i As Long, r As Long
    Dim env As Variant, side As Variant, v As Variant, c As Long, xs As Object, ys As Object, j As Long
    Dim use As Object, oldT As Variant, oldR As Variant, same As Boolean, kept As String, emptied As String, prev As String, note As String
    Set uc = SheetOf(wb, SH_UC)
    If uc Is Nothing Then Err.Raise vbObjectError + 11, , "This workbook has no sheet " & SH_UC & "."
    CheckUcLayout uc
    Set tbl = McAnchor(uc, "M_KAZIK_TABLO", "A8:M1007")
    Set raw = McAnchor(uc, "M_KAZIK_HAM", "O8:W6007")
    n = heads.Count
    If n > tbl.Rows.Count Then Err.Raise vbObjectError + 11, , n & " piles, " & SH_UC & " holds " & tbl.Rows.Count & "."
    Set xs = Ranks(heads, 2)
    Set ys = Ranks(heads, 3)
    Set use = CreateObject("Scripting.Dictionary")
    For Each env In Split(IIf(envUse = "", ENV_LIST, envUse), "|")
        use(CStr(env)) = True
    Next env
    ' a partial read keeps the other envelopes when the piles (elements in order) are the ones already written
    If use.Count < 3 Then
        oldT = tbl.Resize(n).Value
        oldR = raw.Resize(6 * n).Value
        same = UcSame(uc, heads)
        prev = CellTxt(McAnchor(uc, "M_KAZIK_DAMGA", "B2").Value)
        If InStr(prev, " || ") > 0 Then prev = Left$(prev, InStr(prev, " || ") - 1)
    End If
    ReDim out(1 To n, 1 To 13)
    ReDim rw(1 To 6 * n, 1 To 9)
    i = 0: r = 0
    For Each h In heads
        i = i + 1
        out(i, 1) = i: out(i, 2) = h(0): out(i, 3) = h(1)
        out(i, 4) = Round(h(2), 3): out(i, 5) = Round(h(3), 3)
        out(i, 6) = ys(RKey(h(3))): out(i, 7) = xs(RKey(h(2)))
        c = 7
        For Each env In Split(ENV_LIST, "|")
            For Each side In Array("max", "min")
                c = c + 1
                r = r + 1
                If use.Exists(CStr(env)) Then
                    v = ForceOf(res, h, CStr(env), CStr(side))
                    out(i, c) = Round(-v(0), 2)            ' compression +: Min from (max), Mak from (min)
                    rw(r, 1) = h(0): rw(r, 2) = env & "(" & side & ")": rw(r, 3) = v(6)
                    rw(r, 4) = v(0): rw(r, 5) = v(1): rw(r, 6) = v(2): rw(r, 7) = v(3): rw(r, 8) = v(4): rw(r, 9) = v(5)
                ElseIf same Then
                    out(i, c) = oldT(i, c)
                    For j = 1 To 9: rw(r, j) = oldR(r, j): Next j
                End If
            Next side
        Next env
    Next h
    Application.ScreenUpdating = False
    tbl.ClearContents
    raw.ClearContents
    tbl.Resize(n).Value = out
    raw.Resize(6 * n).Value = rw
    If use.Count < 3 Then
        For Each env In Split(ENV_LIST, "|")
            If Not use.Exists(CStr(env)) Then
                If same Then kept = kept & IIf(kept = "", "", ", ") & env Else emptied = emptied & IIf(emptied = "", "", ", ") & env
            End If
        Next env
        stamp = stamp & " [" & Replace(IIf(envUse = "", ENV_LIST, envUse), "|", ", ") & "]"
        If kept <> "" Then stamp = stamp & " || " & kept & " kept from: " & prev
        If kept <> "" Then note = note & "Kept from the last read (same piles): " & kept & "." & vbCrLf
        If emptied <> "" Then note = note & "Emptied (not in this model, and the piles differ from the last read): " & emptied & "." & vbCrLf
    End If
    With McAnchor(uc, "M_KAZIK_DAMGA", "B2")
        If Len(CellTxt(.Offset(0, -1).Value)) = 0 Then .Offset(0, -1).Value = "Model"
        .Value = stamp
    End With
    With McAnchor(uc, "M_KAZIK_ADET", "B3")
        If Len(CellTxt(.Offset(0, -1).Value)) = 0 Then .Offset(0, -1).Value = U("Kaz\u0131k say\u0131s\u0131")
        .Value = n
    End With
    With McAnchor(uc, "M_KAZIK_KAYNAK", "B4")
        If Len(CellTxt(.Offset(0, -1).Value)) = 0 Then .Offset(0, -1).Value = U("Kaz\u0131k listesi")
        .Value = src
    End With
    Application.Calculate
    WriteUc = SH_UC & ": " & n & " piles, " & 6 * n & " raw rows." & vbCrLf & note
End Function

' distinct coordinates (0.01 m) of item field f -> rank 1, 2 ... from the smallest
Private Function Ranks(heads As Collection, ByVal f As Long) As Object
    Dim d As Object, h As Variant, a() As Double, n As Long, i As Long, j As Long, x As Double
    Set d = CreateObject("Scripting.Dictionary")
    For Each h In heads
        If Not d.Exists(RKey(h(f))) Then d(RKey(h(f))) = 0
    Next h
    ReDim a(1 To d.Count)
    For Each h In d.keys: n = n + 1: a(n) = CDbl(h) / 100#: Next h
    For i = 2 To n
        x = a(i): j = i - 1
        Do While j >= 1
            If a(j) <= x Then Exit Do
            a(j + 1) = a(j): j = j - 1
        Loop
        a(j + 1) = x
    Next i
    For i = 1 To n: d(RKey(a(i))) = i: Next i
    Set Ranks = d
End Function

Private Function RKey(ByVal v As Double) As String
    RKey = CStr(CLng(Round(v * 100#, 0)))
End Function

' ---- pile forces for GSD (MIDAS_CIKTISI_KAZIK) -----------------------------------------------
Public Sub MIDAS_KAZIK_GSD_OKU()
    Dim s As String
    mScripted = False
    s = RunGsd(False, UiIdx("kazik", MODE_GROUP), "")
    UiMsg s, IIf(Left$(s, 8) = "STOPPED:", vbExclamation, vbInformation), MSG_TITLE
End Sub

Public Sub MIDAS_KAZIK_GSD_RUN(ByVal reportFile As String, Optional ByVal pileMode As Long = 1, Optional ByVal pileList As String = "")
    mScripted = True
    WriteText reportFile, RunGsd(False, pileMode, pileList)
End Sub

Public Sub MIDAS_KAZIK_GSD_PLANYAZ(ByVal reportFile As String, Optional ByVal pileMode As Long = 1, Optional ByVal pileList As String = "")
    mScripted = True
    WriteText reportFile, RunGsd(True, pileMode, pileList)
End Sub

Private Function RunGsd(ByVal planOnly As Boolean, ByVal pileMode As Long, ByVal pileList As String) As String
    Dim wb As Workbook, ds As Worksheet, pj As String, tag As String, t As String, lc As Object, combos As Collection, els As Collection
    Dim src As String, s As String, c As Variant, h As Variant, keys As String, res As Object, two As Object, ans As Variant
    Dim hasS As Boolean, hasD As Boolean, nS As Long, nD As Long, seenE As Object
    On Error GoTo Fail
    mLog = ""
    Set wb = ThisWorkbook
    t = TemelTipi(wb)
    If Len(t) > 0 And t <> "KAZIKLI" Then RunGsd = "STOPPED: the foundation type is " & t & ", not KAZIKLI: nothing was read.": Exit Function
    Set ds = SheetOf(wb, SH_GSD)
    If ds Is Nothing Then RunGsd = "STOPPED: this workbook has no sheet " & SH_GSD & ".": Exit Function
    If GsdBlock(ds).Columns.Count <> 9 Then RunGsd = "STOPPED: M_TESIR_BLOK must be 9 columns wide.": Exit Function
    If Not McGetKey(wb, Not mScripted, MSG_TITLE) Then RunGsd = "STOPPED: no MAPI key: nothing was read.": Exit Function
    Logp "Civil NX: model"
    If Not Api("GET", "db/PJCF", "", pj) Then
        If Left$(pj, 10) = "HTTP error" Or pj = "" Then RunGsd = "STOPPED: Civil NX does not answer (is it open, is the MAPI connected?). " & Left$(pj, 200): Exit Function
        pj = ""
    End If
    tag = ModelTag(pj)
    Logp "Civil NX: combinations"
    Set lc = TableGet("LCOM-GEN")
    If lc Is Nothing Then Err.Raise vbObjectError + 4, , "The open model (" & tag & ") has no load combinations (db/LCOM-GEN)."
    Set combos = EnvMembers(lc)
    If combos.Count = 0 Then Err.Raise vbObjectError + 5, , "The open model (" & tag & ") has none of the envelopes " & Replace(ENV_LIST, "|", ", ") & ": open the abutment model built by MIDAS_KENARAYAK."
    For Each c In combos
        If c(2) Then hasD = True: nD = nD + 1 Else hasS = True: nS = nS + 1
    Next c
    If UiStopped() Then RunGsd = "STOPPED: by the user.": Exit Function
    If pileMode < MODE_GROUP Or pileMode > MODE_ONLY Then pileMode = MODE_GROUP
    If pileMode <> MODE_GROUP And Not mScripted Then
        Logp "Civil NX: selected elements"
        ans = AskPiles(CellTxt(McAnchor(ds, "M_TESIR_SECIM", "DG4", True).Value), pileMode, True)
        If VarType(ans) = vbBoolean Then RunGsd = "STOPPED: nothing was read.": Exit Function
        pileList = CStr(ans)
    End If
    Logp "Civil NX: elements"
    Set els = GsdElems(src, pileMode, pileList)
    Set seenE = CreateObject("Scripting.Dictionary")
    For Each h In els
        keys = keys & IIf(keys = "", "", ",") & h(0)
        seenE(CStr(h(0))) = True
    Next h
    If els.Count * combos.Count > GsdBlock(ds).Rows.Count Then Err.Raise vbObjectError + 13, , els.Count & " elements x " & combos.Count & " combinations = " & els.Count * combos.Count & " rows, M_TESIR_BLOK holds " & GsdBlock(ds).Rows.Count & ": choose fewer elements."
    s = "Model: " & tag & vbCrLf & "Combinations: " & combos.Count & " (static " & nS & ", seismic " & nD & ")" & vbCrLf & "Elements: " & els.Count & " (" & src & ")" & vbCrLf & "Rows: " & els.Count * combos.Count
    If planOnly Then
        s = "PLAN. " & s & vbCrLf
        For Each c In combos: s = s & "  " & c(1) & " (" & c(0) & ")" & vbCrLf: Next c
        For Each h In els: s = s & "  " & h(0) & " " & h(4) & "-end x " & Format$(h(2), "0.00") & " y " & Format$(h(3), "0.00") & vbCrLf: Next h
        RunGsd = s
        Exit Function
    End If
    If Not mScripted Then
        s = s & vbCrLf & KeptText(ds, hasS, hasD, seenE) & vbCrLf & vbCrLf & "The rows of this model replace its rows in the GSD load block (" & SH_GSD & "); cells of the block become values. Continue?"
        If UiMsg(s, vbOKCancel + vbQuestion, MSG_TITLE) <> vbOK Then RunGsd = "STOPPED: nothing was read.": Exit Function
    End If
    Set res = CreateObject("Scripting.Dictionary")
    Set two = GsdForces(keys, combos, res)
    If UiStopped() Then RunGsd = "STOPPED: by the user (nothing was written).": Exit Function
    Logp "Workbook: " & SH_GSD
    s = GsdWrite(ds, els, combos, two, res, hasS, hasD, tag, src)
    RunGsd = "Done. Model: " & tag & vbCrLf & s & mLog
    Exit Function
Fail:
    Application.ScreenUpdating = True
    RunGsd = "STOPPED: " & Err.Description & IIf(Len(mLog) > 0, vbCrLf & mLog, "")
End Function

' the cases the envelopes are made of: items Array(type "CB" / "ST" ..., name, True = seismic), in envelope order
Private Function EnvMembers(lc As Object) As Collection
    Dim env As Variant, k As Variant, it As Variant, seen As Object, v As Variant
    Set EnvMembers = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    For Each env In Split(ENV_LIST, "|")
        For Each k In lc.keys
            If StrComp(CStr(lc(k)("NAME")), CStr(env), vbTextCompare) = 0 Then
                v = lc(k)("vCOMB")
                If IsArray(v) Then
                    For Each it In v
                        If Not seen.Exists(it("ANAL") & "|" & it("LCNAME")) Then
                            seen(it("ANAL") & "|" & it("LCNAME")) = True
                            EnvMembers.Add Array(CStr(it("ANAL")), CStr(it("LCNAME")), (env = "ENV_DEP"))
                        End If
                    Next it
                End If
            End If
        Next k
    Next env
End Function

' elements of the GSD block: items Array(elem, node, x, y, "I" | "J" = upper end); 1 pile heads of group KAZIK,
' 2 pile heads + every listed element, 3 only the listed elements
Private Function GsdElems(ByRef src As String, ByVal pileMode As Long, ByVal pileList As String) As Collection
    Dim gr As Object, ev As Object, nd As Object, heads As Collection, els As Object, one As Object, e As Variant, h As Variant, seen As Object, n As Long, a() As Long, i As Long
    If pileMode = MODE_GROUP Then Set GsdElems = PileHeads(src): Exit Function
    Set GsdElems = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    src = ""
    If pileMode = MODE_ADD Then
        Set heads = PileHeads(src)
        For Each h In heads
            seen(h(0) & h(4)) = True
            GsdElems.Add h
        Next h
    End If
    Set gr = TableGet("GRUP"): Set ev = TableGet("ELEM"): Set nd = TableGet("NODE")
    If ev Is Nothing Or nd Is Nothing Then Err.Raise vbObjectError + 7, , "Elements or nodes could not be read."
    Set els = ListElems(pileList, gr, ev)
    If els.Count = 0 Then
        If pileMode = MODE_ONLY Then Err.Raise vbObjectError + 7, , "No elements chosen: nothing was read."
        Exit Function
    End If
    ReDim a(1 To els.Count)
    For Each e In els.keys
        i = i + 1: a(i) = CLng(e)
    Next e
    SortLongs a
    For i = 1 To UBound(a)
        Set one = CreateObject("Scripting.Dictionary")
        one(CStr(a(i))) = True
        For Each h In ChainHeads(one, ev, nd, "the list", True)
            If Not seen.Exists(h(0) & h(4)) Then
                seen(h(0) & h(4)) = True
                GsdElems.Add h
                n = n + 1
            End If
        Next h
    Next i
    src = src & IIf(src = "", "chosen elements [", " + added elements [") & ShortList(els) & "], " & n & " elements, upper end"
End Function

' forces of the elements: res("<elem>|<I|J>|<case>|<suffix>") = Array(Pu, My, Mz, Vy, Vz, T); a combination that
' answers no row as (CB) is read as (CB:max) / (CB:min); returns those
Private Function GsdForces(ByVal keys As String, combos As Collection, res As Object) As Object
    Dim c As Variant, req As String, miss As String, got As Object, two As Object
    Set two = CreateObject("Scripting.Dictionary")
    Set got = CreateObject("Scripting.Dictionary")
    For Each c In combos
        req = req & "|" & c(1) & "(" & c(0) & ")"
    Next c
    Logp "Civil NX: pile forces"
    GsdTable keys, Mid$(req, 2), "", res, got
    For Each c In combos
        If Not got.Exists(CStr(c(1))) Then
            If c(0) <> "CB" Then Err.Raise vbObjectError + 10, , "No result for " & c(1) & "(" & c(0) & ") (has the analysis been run?)."
            two(CStr(c(1))) = True
            miss = miss & "|" & c(1)
        End If
    Next c
    If Len(miss) > 0 Then
        miss = Mid$(miss, 2)
        GsdTable keys, Replace(miss, "|", "(CB:max)|") & "(CB:max)", "(max)", res, got
        GsdTable keys, Replace(miss, "|", "(CB:min)|") & "(CB:min)", "(min)", res, got
        For Each c In combos
            If Not got.Exists(CStr(c(1))) Then Err.Raise vbObjectError + 10, , "No result for " & c(1) & " (has the analysis been run?)."
        Next c
        mLog = mLog & "Read as (max) / (min), no plain result: " & Replace(miss, "|", ", ") & "." & vbCrLf
    End If
    Set GsdForces = two
End Function

Private Sub GsdTable(ByVal keys As String, ByVal reqList As String, ByVal sfx As String, res As Object, got As Object)
    Dim nm As Variant, lcs As String, body As String, resp As String, root As Object, tb As Object, head As Variant, data As Variant
    Dim iE As Long, iL As Long, iP As Long, iA As Long, iVy As Long, iVz As Long, iT As Long, iMy As Long, iMz As Long
    Dim j As Long, r As Long, ld As String, want As Object
    Set want = CreateObject("Scripting.Dictionary")
    For Each nm In Split(reqList, "|")
        lcs = lcs & IIf(lcs = "", "", ",") & QS(CStr(nm))
        want(BaseName(CStr(nm))) = True
    Next nm
    body = "{""Argument"":{""TABLE_NAME"":""F"",""TABLE_TYPE"":""BEAMFORCE"",""UNIT"":{""FORCE"":""KN"",""DIST"":""M""}," & _
           """NODE_ELEMS"":{""KEYS"":[" & keys & "]},""LOAD_CASE_NAMES"":[" & lcs & "],""PARTS"":[""PartI"",""PartJ""]}}"
    If Not Api("POST", "post/TABLE", body, resp) Then Err.Raise vbObjectError + 10, , "Pile forces: " & Left$(resp, 300)
    Set root = ParseJson(resp)
    If root Is Nothing Then Err.Raise vbObjectError + 10, , "Pile forces: no answer."
    If Not root.Exists("F") Then Exit Sub
    Set tb = root("F")
    head = tb("HEAD"): data = tb("DATA")
    If IsEmpty(data) Or Not IsArray(data) Then Exit Sub
    For j = 1 To UBound(head)
        Select Case CStr(head(j))
            Case "Elem": iE = j
            Case "Load": iL = j
            Case "Part": iP = j
            Case "Axial": iA = j
            Case "Shear-y": iVy = j
            Case "Shear-z": iVz = j
            Case "Torsion": iT = j
            Case "Moment-y": iMy = j
            Case "Moment-z": iMz = j
        End Select
    Next j
    If iE * iL * iP * iA * iVy * iVz * iT * iMy * iMz = 0 Then Err.Raise vbObjectError + 10, , "Pile forces: unexpected columns."
    For r = 1 To UBound(data)
        ld = CStr(data(r)(iL))
        If Not want.Exists(ld) Then ld = BaseName(ld)
        If want.Exists(ld) Then
            got(ld) = True
            res(CStr(CLng(Val(CStr(data(r)(iE))))) & "|" & UCase$(Left$(CStr(data(r)(iP)), 1)) & "|" & ld & "|" & sfx) = Array( _
                -Val(CStr(data(r)(iA))), Val(CStr(data(r)(iMy))), Val(CStr(data(r)(iMz))), _
                Val(CStr(data(r)(iVy))), Val(CStr(data(r)(iVz))), Val(CStr(data(r)(iT))))
        End If
    Next r
End Sub

Private Function IsSeismicRow(ByVal nm As String) As Boolean
    IsSeismicRow = (Left$(nm, 5) = "TKDY-" Or Left$(nm, 7) = "H-TKDY-" Or Left$(nm, 8) = "ENV_DEP_")
End Function

' element number in a row name '<combination>_<element>_<I|J>', "" when it has none
Private Function ElemOfRow(ByVal nm As String) As String
    Dim p As Long, q As Long
    p = InStrRev(nm, "_")
    If p < 3 Then Exit Function
    q = InStrRev(nm, "_", p - 1)
    If q < 1 Then Exit Function
    If AllDigits(Mid$(nm, q + 1, p - q - 1)) Then ElemOfRow = Mid$(nm, q + 1, p - q - 1)
End Function

' rows of the block down to the last name
Private Function BlockUsed(blk As Range) As Long
    Dim f As Range
    Set f = blk.Columns(3).Find(What:="*", After:=blk.Cells(1, 3), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not f Is Nothing Then BlockUsed = f.Row - blk.Row + 1
End Function

' what happens to the rows of the other model (text for the question)
Private Function KeptText(ds As Worksheet, ByVal hasS As Boolean, ByVal hasD As Boolean, seenE As Object) As String
    Dim blk As Range, used As Long, v As Variant, i As Long, nm As String, n As Long, other As Object, e As Variant, dif As Boolean
    If hasS And hasD Then Exit Function
    Set blk = GsdBlock(ds)
    used = BlockUsed(blk)
    Set other = CreateObject("Scripting.Dictionary")
    If used > 0 Then
        v = blk.Resize(used).Value
        For i = 1 To used
            nm = CellTxt(v(i, 3))
            If Len(nm) > 0 And IsSeismicRow(nm) = hasS Then
                n = n + 1
                other(ElemOfRow(nm)) = True
            End If
        Next i
    End If
    If n = 0 Then KeptText = "The block has no " & IIf(hasS, "seismic", "static") & " rows yet: read that model too.": Exit Function
    If other.Count <> seenE.Count Then dif = True
    For Each e In other.keys
        If Not seenE.Exists(CStr(e)) Then dif = True
    Next e
    KeptText = n & IIf(hasS, " seismic", " static") & " rows of the other model are kept" & IIf(dif, " - WARNING: they are of other elements (" & other.Count & ").", " (same elements).")
End Function

Private Function GsdWrite(ds As Worksheet, els As Collection, combos As Collection, two As Object, res As Object, ByVal hasS As Boolean, ByVal hasD As Boolean, ByVal tag As String, ByVal src As String) As String
    Dim blk As Range, used As Long, v As Variant, i As Long, j As Long, nm As String, rowsS As Collection, rowsD As Collection, rw As Variant
    Dim c As Variant, h As Variant, sfx As Variant, sfxs As Variant, key As String, n As Long, out() As Variant, f As Variant, nKept As Long, stamp As String
    Set blk = GsdBlock(ds)
    Set rowsS = New Collection: Set rowsD = New Collection
    used = BlockUsed(blk)
    If used > 0 Then
        v = blk.Resize(used).Value
        For i = 1 To used
            nm = CellTxt(v(i, 3))
            If Len(nm) > 0 Then
                rw = Array(v(i, 3), v(i, 4), v(i, 5), v(i, 6), v(i, 7), v(i, 8), v(i, 9))
                If IsSeismicRow(nm) Then
                    If Not hasD Then rowsD.Add rw: nKept = nKept + 1
                ElseIf Not hasS Then
                    rowsS.Add rw: nKept = nKept + 1
                End If
            End If
        Next i
    End If
    For Each c In combos
        If two.Exists(CStr(c(1))) Then sfxs = Array("(max)", "(min)") Else sfxs = Array("")
        For Each h In els
            For Each sfx In sfxs
                key = h(0) & "|" & h(4) & "|" & c(1) & "|" & sfx
                If Not res.Exists(key) Then Err.Raise vbObjectError + 12, , "No force for " & c(1) & sfx & " at the " & h(4) & " end of element " & h(0) & "."
                f = res(key)
                rw = Array(c(1) & sfx & "_" & h(0) & "_" & h(4), Round(f(0), 2), Round(f(1), 2), Round(f(2), 2), Round(f(3), 2), Round(f(4), 2), Round(f(5), 2))
                If c(2) Then rowsD.Add rw Else rowsS.Add rw
            Next sfx
        Next h
    Next c
    n = rowsS.Count + rowsD.Count
    If n > blk.Rows.Count Then Err.Raise vbObjectError + 13, , n & " rows, M_TESIR_BLOK holds " & blk.Rows.Count & ": choose fewer elements."
    ReDim out(1 To n, 1 To 9)
    i = 0
    For Each rw In rowsS
        i = i + 1: out(i, 1) = i: out(i, 2) = i
        For j = 0 To 6: out(i, 3 + j) = rw(j): Next j
    Next rw
    For Each rw In rowsD
        i = i + 1: out(i, 1) = i: out(i, 2) = i
        For j = 0 To 6: out(i, 3 + j) = rw(j): Next j
    Next rw
    Application.ScreenUpdating = False
    blk.Resize(n).Value = out
    If used > n Then blk.Rows(n + 1).Resize(used - n).ClearContents
    stamp = tag & " | " & Format$(Now, "yyyy-mm-dd hh:nn") & " | " & src
    GsdStamp McAnchor(ds, "M_TESIR_STATIK", "DG2", True), "Statik model", IIf(hasS, stamp, "")
    GsdStamp McAnchor(ds, "M_TESIR_DEPREM", "DG3", True), "Deprem modeli", IIf(hasD, stamp, "")
    GsdStamp McAnchor(ds, "M_TESIR_SECIM", "DG4", True), "Elemanlar", src
    Application.Calculate
    Application.ScreenUpdating = True
    GsdWrite = SH_GSD & ": " & n & " rows (static " & rowsS.Count & ", seismic " & rowsD.Count & "), " & nKept & " of them kept from the other model." & vbCrLf & _
        "Copy No..T (second to ninth column of the block) to GSD, run it and read the export with GSD sonuclari; until then the ratios are those of the old run." & vbCrLf
End Function

Private Sub GsdStamp(cell As Range, ByVal label As String, ByVal txt As String)
    If Len(CellTxt(cell.Offset(0, -1).Value)) = 0 Then cell.Offset(0, -1).Value = label
    If Len(txt) > 0 Then cell.Value = txt
End Sub

' GSD load block of the data sheet (9 columns), as GSD_OKU reads it
Private Function GsdBlock(ds As Worksheet) As Range
    Set GsdBlock = McAnchor(ds, "M_TESIR_BLOK", "CW26:DE3839", True)
End Function

' ---- report pages ---------------------------------------------------------------------------
Private Function PagesRefresh(wb As Workbook) As String
    Dim uc As Worksheet, rp As Worksheet, gs As Worksheet, n As Long, lg As String, r As Variant
    Set uc = SheetOf(wb, SH_UC)
    If uc Is Nothing Then Err.Raise vbObjectError + 20, , "This workbook has no sheet " & SH_UC & "."
    n = CLng(NumOr0(McAnchor(uc, "M_KAZIK_ADET", "B3").Value))
    If n < 0 Or n > MAX_PILES Then Err.Raise vbObjectError + 20, , SH_UC & ": pile count " & n & " outside 0.." & MAX_PILES & "."
    Application.ScreenUpdating = False
    Set rp = SheetOf(wb, SH_REP)
    If rp Is Nothing Then Set rp = SheetOf(wb, SH_REP2)
    If rp Is Nothing Then
        lg = lg & "No sheet " & SH_REP & " / " & SH_REP2 & ": its tables were not resized." & vbCrLf
    Else
        lg = lg & ReportPages(rp, n)
    End If
    Set gs = SheetOf(wb, SH_GEO)
    If gs Is Nothing Then
        lg = lg & "No sheet " & SH_GEO & ": its table was not resized." & vbCrLf
    Else
        lg = lg & GeoRows(gs, n)
    End If
    Application.Calculate
    On Error Resume Next
    r = Application.Run("'" & Replace(wb.Name, "'", "''") & "'!RAPOR_LISTE.LISTELERI_AYARLA_SESSIZ")
    If Err.Number <> 0 Then
        lg = lg & "Lists not rebuilt (RAPOR_LISTE.LISTELERI_AYARLA_SESSIZ): " & Err.Description & vbCrLf
        Err.Clear
    Else
        lg = lg & "Contents and figure / table lists rebuilt." & vbCrLf
        If Len(CStr(r)) > 0 Then lg = lg & CStr(r) & vbCrLf
    End If
    On Error GoTo 0
    Application.ScreenUpdating = True
    PagesRefresh = lg
End Function

' 2A_KAZIK: pages of the per-pile table and of 8.6.7, page numbers, print area
Private Function ReportPages(rp As Worksheet, ByVal n As Long) As String
    Dim t1 As Long, t2 As Long, t3 As Long, have1 As Long, have2 As Long, need1 As Long, need2 As Long, after As Long
    Dim nPages As Long, k As Long, top As Long, f As String, pa As Range, addr As String, changed As Boolean, lg As String, two As Boolean
    two = (StrComp(rp.Name, SH_REP, vbTextCompare) = 0)      ' 8.6.7 page: abutment only
    PageTops rp, t1, t2, t3, two
    have1 = (t2 - t1) \ PG: have2 = (t3 - t2) \ PG
    need1 = PagesFor(n, PER1)
    If two Then need2 = PagesFor(n, PER2)
    after = PagesAfter(rp, t3)
    If need2 <> have2 Then ChangePages rp, t2, t3, need2 - have2, True: changed = True
    If need1 <> have1 Then ChangePages rp, t1, t2, need1 - have1, False: changed = True
    PageTops rp, t1, t2, t3, two
    nPages = (t3 - PG_FIRST) \ PG + after
    ' page numbers: each page = the page before + 1 (page 1 keeps its link to the sheet before)
    For k = 2 To nPages
        top = PG_FIRST + PG * (k - 1)
        f = "=W" & (top - 3) & "+1"
        If rp.Cells(top + 43, COL_W).Formula <> f Then rp.Cells(top + 43, COL_W).Formula = f
    Next k
    ' print area: B:X of every page
    For k = 1 To nPages
        top = PG_FIRST + PG * (k - 1)
        If pa Is Nothing Then
            Set pa = rp.Range(rp.Cells(top, 2), rp.Cells(top + 44, 24))
        Else
            Set pa = Union(pa, rp.Range(rp.Cells(top, 2), rp.Cells(top + 44, 24)))
        End If
        addr = addr & "," & rp.Range(rp.Cells(top, 2), rp.Cells(top + 44, 24)).Address
    Next k
    If Replace(Replace(rp.PageSetup.PrintArea, ";", ","), "'" & rp.Name & "'!", "") <> Mid$(addr, 2) Then SetPrintAreaLong rp, Mid$(addr, 2), pa
    lg = rp.Name & ": per-pile table " & need1 & " page(s)" & IIf(need1 <> have1, " (was " & have1 & ")", "") & _
         IIf(two, ", 8.6.7 " & need2 & " page(s)" & IIf(need2 <> have2, " (was " & have2 & ")", ""), "") & ", " & nPages & " pages, page numbers chained." & vbCrLf
    ReportPages = lg
End Function

Private Sub PageTops(rp As Worksheet, ByRef t1 As Long, ByRef t2 As Long, ByRef t3 As Long, ByVal two As Boolean)
    t1 = McAnchor(rp, "M_KAZIK_SAYFA1", "A186:A231", True).Row
    If Not two Then
        t3 = McAnchor(rp, "M_KAZIK_SAYFA3", "A232:A277", True).Row
        t2 = t3
        If (t1 - PG_FIRST) Mod PG <> 0 Or t3 <= t1 Or (t3 - t1) Mod PG <> 0 Then _
            Err.Raise vbObjectError + 21, , rp.Name & ": the names M_KAZIK_SAYFA1 / 3 (rows " & t1 & ", " & t3 & ") are not on page starts " & PG & " rows apart."
        If StrComp(CellTxt(rp.Cells(t1 + 5, 3).Value), "Elem", vbTextCompare) <> 0 Then _
            Err.Raise vbObjectError + 21, , rp.Name & ": no 'Elem' header in row " & t1 + 5 & " of the per-pile page (M_KAZIK_SAYFA1)."
        Exit Sub
    End If
    t2 = McAnchor(rp, "M_KAZIK_SAYFA2", "A232:A277", True).Row
    t3 = McAnchor(rp, "M_KAZIK_SAYFA3", "A278:A323", True).Row
    If (t1 - PG_FIRST) Mod PG <> 0 Or t2 <= t1 Or t3 <= t2 Or (t2 - t1) Mod PG <> 0 Or (t3 - t2) Mod PG <> 0 Then _
        Err.Raise vbObjectError + 21, , SH_REP & ": the names M_KAZIK_SAYFA1 / 2 / 3 (rows " & t1 & ", " & t2 & ", " & t3 & ") are not on page starts " & PG & " rows apart."
    If StrComp(CellTxt(rp.Cells(t1 + 5, 3).Value), "Elem", vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 21, , SH_REP & ": no 'Elem' header in row " & t1 + 5 & " of the per-pile page (M_KAZIK_SAYFA1)."
    If StrComp(CellTxt(rp.Cells(t2 + 9, 4).Value), "Elem", vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 21, , SH_REP & ": no 'Elem' header in row " & t2 + 9 & " of the 8.6.7 page (M_KAZIK_SAYFA2)."
End Sub

Private Function PagesFor(ByVal n As Long, ByVal per As Long) As Long
    PagesFor = (n + per - 1) \ per
    If PagesFor < 1 Then PagesFor = 1
End Function

' pages from row t3 to the end of the current print area (at least 1)
Private Function PagesAfter(rp As Worksheet, ByVal t3 As Long) As Long
    Dim parts As Variant, p As Variant, lastRow As Long, rg As Range
    On Error Resume Next
    parts = Split(Replace(rp.PageSetup.PrintArea, ";", ","), ",")
    For Each p In parts
        Set rg = Nothing
        Set rg = rp.Range(CStr(p))
        If Not rg Is Nothing Then
            If rg.Row + rg.Rows.Count - 1 > lastRow Then lastRow = rg.Row + rg.Rows.Count - 1
        End If
    Next p
    On Error GoTo 0
    PagesAfter = (lastRow - t3 + 2) \ PG
    If PagesAfter < 1 Then PagesAfter = 1
End Function

' d > 0: d copies of the base page (rows tBase ..) inserted before row tEnd; d < 0: the last -d pages
' before tEnd deleted (never the base page). clearHead: the copies lose the heading rows 4..7 of the page.
Private Sub ChangePages(rp As Worksheet, ByVal tBase As Long, ByVal tEnd As Long, ByVal d As Long, ByVal clearHead As Boolean)
    Dim i As Long, nt As Long
    If d > 0 Then
        For i = 1 To d
            nt = tEnd + PG * (i - 1)
            rp.Range(rp.Rows(tBase), rp.Rows(tBase + PG - 1)).Copy
            rp.Rows(nt).Insert Shift:=xlDown
            Application.CutCopyMode = False
            If clearHead Then rp.Range(rp.Cells(nt + 4, 2), rp.Cells(nt + 7, 54)).ClearContents
        Next i
    ElseIf d < 0 Then
        If tEnd + PG * d <= tBase Then Err.Raise vbObjectError + 22, , SH_REP & ": cannot delete the base page."
        rp.Range(rp.Rows(tEnd + PG * d), rp.Rows(tEnd - 1)).Delete Shift:=xlUp
    End If
End Sub

' GEO_KAZIK: one table row per pile, at least GEO_MIN; cells of the table columns inserted / deleted
' above its last row, the formulas of the row above copied down; print area to the last row
Private Function GeoRows(gs As Worksheet, ByVal n As Long) As String
    Dim tb As Range, have As Long, want As Long, d As Long, lastR As Long, c1 As Long, c2 As Long, pa As Range
    Set tb = McAnchor(gs, "M_KAZIK_GEO", "K12:S46", True)
    If tb.Rows.Count < 3 Then Err.Raise vbObjectError + 23, , SH_GEO & ": M_KAZIK_GEO must have at least 3 rows."
    If tb.Row < 3 Then Err.Raise vbObjectError + 23, , SH_GEO & ": M_KAZIK_GEO needs its header above it."
    If StrComp(CellTxt(gs.Cells(tb.Row - 2, tb.Column).Value), "Elem", vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 23, , SH_GEO & ": no 'Elem' header two rows above M_KAZIK_GEO."
    have = tb.Rows.Count
    want = n
    If want < GEO_MIN Then want = GEO_MIN
    c1 = tb.Column: c2 = tb.Column + tb.Columns.Count - 1
    lastR = tb.Row + have - 1
    If want > have Then
        d = want - have
        gs.Range(gs.Cells(lastR, c1), gs.Cells(lastR + d - 1, c2)).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
        gs.Range(gs.Cells(lastR - 1, c1), gs.Cells(lastR - 1, c2)).Copy Destination:=gs.Range(gs.Cells(lastR, c1), gs.Cells(lastR + d - 1, c2))
        Application.CutCopyMode = False
    ElseIf want < have Then
        d = have - want
        gs.Range(gs.Cells(lastR - d, c1), gs.Cells(lastR - 1, c2)).Delete Shift:=xlShiftUp
    End If
    Set tb = McAnchor(gs, "M_KAZIK_GEO", "K12:S46", True)
    lastR = tb.Row + tb.Rows.Count - 1
    Set pa = gs.Range(gs.Cells(1, 1), gs.Cells(lastR, c2))
    If Replace(Replace(gs.PageSetup.PrintArea, "'" & gs.Name & "'!", ""), gs.Name & "!", "") <> pa.Address Then gs.PageSetup.PrintArea = pa.Address
    GeoRows = SH_GEO & ": " & tb.Rows.Count & " table rows" & IIf(want <> have, " (was " & have & ")", "") & "." & vbCrLf
End Function

' print area longer than 255 characters: through the sheet's Print_Area name (as ORTA_AYAK)
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
    If own Is Nothing Then Err.Raise vbObjectError + 24, , ws.Name & ": print-area name not found."
    own.RefersTo = pa
End Sub

Private Function NameOnSheet(nm As Name, ws As Worksheet) As Boolean
    On Error Resume Next
    NameOnSheet = (TypeOf nm.Parent Is Worksheet)
    If NameOnSheet Then NameOnSheet = (nm.Parent.Name = ws.Name)
    On Error GoTo 0
End Function

' ---- helpers --------------------------------------------------------------------------------
' foundation type in capitals without Turkish dots; "" when the workbook has no such name
Private Function TemelTipi(wb As Workbook) As String
    Dim nm As Variant, v As Variant, s As String
    For Each nm In Split(TYPE_NAMES, "|")
        v = Empty
        On Error Resume Next
        v = wb.Names(CStr(nm)).RefersToRange.Cells(1, 1).Value
        If Err.Number <> 0 Then v = Empty
        On Error GoTo 0
        If Not IsEmpty(v) Then
            s = UCase$(CellTxt(v))
            s = Replace(Replace(Replace(Replace(s, ChrW(304), "I"), ChrW(305), "I"), ChrW(220), "U"), ChrW(252), "U")
            TemelTipi = s
            Exit Function
        End If
    Next nm
End Function

Private Function SheetOf(wb As Workbook, ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetOf = wb.Worksheets(nm)
    On Error GoTo 0
End Function

Private Function SortByElem(c As Collection) As Collection
    Dim a() As Long, d As Object, it As Variant, n As Long, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    ReDim a(1 To c.Count)
    For Each it In c
        n = n + 1: a(n) = it(0): d(CStr(it(0))) = it
    Next it
    SortLongs a
    Set SortByElem = New Collection
    For i = 1 To n: SortByElem.Add d(CStr(a(i))): Next i
End Function

Private Function CellTxt(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    CellTxt = Trim$(CStr(v))
End Function

' "ENV_SER(max)" -> "ENV_SER"
Private Function BaseName(ByVal s As String) As String
    Dim p As Long
    BaseName = s
    If Right$(s, 1) <> ")" Then Exit Function
    p = InStrRev(s, "(")
    If p > 1 Then BaseName = Left$(s, p - 1)
End Function

Private Function NodeC(nd As Object, ByVal n As Long, ByVal c As String) As Double
    If Not nd.Exists(CStr(n)) Then Err.Raise vbObjectError + 14, , "Node " & n & " is not in the model."
    NodeC = CDbl(nd(CStr(n))(c))
End Function

Private Sub SortLongs(a() As Long)
    Dim i As Long, j As Long, x As Long
    For i = LBound(a) + 1 To UBound(a)
        x = a(i): j = i - 1
        Do While j >= LBound(a)
            If a(j) <= x Then Exit Do
            a(j + 1) = a(j): j = j - 1
        Loop
        a(j + 1) = x
    Next i
End Sub

Private Function ModelTag(ByVal pj As String) As String
    Dim root As Object, k As Variant, rec As Object
    On Error GoTo Quit
    Set root = ParseJson(pj)
    If root Is Nothing Then GoTo Quit
    If Not root.Exists("PJCF") Then GoTo Quit
    For Each k In root("PJCF").keys
        Set rec = root("PJCF")(k)
        If rec.Exists("FILE_NAME") Then ModelTag = CStr(rec("FILE_NAME"))
        If rec.Exists("COMMENT") Then ModelTag = ModelTag & IIf(ModelTag = "", "", " (") & CStr(rec("COMMENT")) & IIf(ModelTag = CStr(rec("COMMENT")), "", ")")
        Exit For
    Next k
Quit:
    If ModelTag = "" Then ModelTag = "model without project data"
End Function

Private Function TableGet(ByVal t As String) As Object
    Dim resp As String, root As Object, k As Variant
    If Not Api("GET", "db/" & t, "", resp) Then Exit Function
    Set root = ParseJson(resp)
    If root Is Nothing Then Exit Function
    For Each k In root.keys
        If UCase$(CStr(k)) = UCase$(t) Then Set TableGet = root(k): Exit Function
    Next k
End Function

Private Function Api(ByVal method As String, ByVal ep As String, ByVal body As String, ByRef resp As String) As Boolean
    Api = McApi(method, ep, body, resp)
    If McStatus() = 404 Then McForgetKey
End Function

Private Sub Logp(ByVal msgLine As String)
    If mScripted Then Exit Sub
    On Error Resume Next
    UiLog msgLine
End Sub

Private Function U(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, "\u")
    Do While p > 0
        s = Left$(s, p - 1) & ChrW(CLng("&H" & Mid$(s, p + 2, 4))) & Mid$(s, p + 6)
        p = InStr(p + 1, s, "\u")
    Loop
    U = s
End Function
