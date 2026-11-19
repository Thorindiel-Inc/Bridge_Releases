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
'  MAPI key: read by MIDAS_CORE (G_MAPI_KEY, INPUT!T3, MIDAS_MAPI_KEY, asked; memory only).
'  Test entries: MIDAS_KAZIK_RUN <report file> (no questions), MIDAS_KAZIK_PLAN (piles and envelopes
'  of the open model; nothing is read from the result tables or written), MIDAS_KAZIK_TEST_YAZ n,
'  <report file> (n made-up piles into KAZIK_UC without a model, then the pages; for the growth tests).
' ============================================================================

Private Const SCRIPT_ID As String = "midas-pile-loads"
Private Const SCRIPT_VERSION As String = "2026-11-19b"
Private Const SCRIPT_CHANGELOG As String = "Reads the envelopes the open model has (static or seismic abutment model of the model gate) and keeps the others from the last read when the piles are the same; New: pile-head envelope loads (ENV_SER / ENV_STR / ENV_DEP, max and min) of the open abutment model into KAZIK_UC; the 2A_KAZIK per-pile table, section 8.6.7 and GEO_KAZIK grow with the pile count, page numbers and lists follow"
Private Const SCRIPT_NEEDS As String = "midas-core"   ' UPDATE adds this module when the workbook lacks it
' macro-window row(s) of this module (PANEL reads them): order|group|title|macro|sheet|text|fields[|modules]
'@TOOL 84|MIDAS sonu\u00E7lar\u0131|Kaz\u0131k u\u00E7 y\u00FCkleri|MIDAS_KAZIK_OKU||Civil NX'te a\u00E7\u0131k kenar ayak modelinden kaz\u0131k ba\u015Flar\u0131n\u0131n zarf y\u00FCklerini (ENV_SER / ENV_STR / ENV_DEP, en b\u00FCy\u00FCk ve en k\u00FC\u00E7\u00FCk) KAZIK_UC sayfas\u0131na okur; rapor tablolar\u0131 (2A_KAZIK, 8.6.7, GEO_KAZIK) kaz\u0131k say\u0131s\u0131na g\u00F6re b\u00FCy\u00FCr. Modele bir \u015Fey yazmaz.|

Private Const MSG_TITLE As String = "MIDAS_KAZIK"
Private Const SH_UC As String = "KAZIK_UC"
Private Const SH_REP As String = "2A_KAZIK"
Private Const SH_GEO As String = "GEO_KAZIK"
Private Const PILE_GROUP As String = "KAZIK"
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
    s = RunRead(False)
    UiMsg s, IIf(Left$(s, 8) = "STOPPED:", vbExclamation, vbInformation), MSG_TITLE
End Sub

Public Sub MIDAS_KAZIK_RUN(ByVal reportFile As String)
    mScripted = True
    WriteText reportFile, RunRead(False)
End Sub

Public Function MIDAS_KAZIK_PLAN() As String
    mScripted = True
    MIDAS_KAZIK_PLAN = RunRead(True)
End Function

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
Private Function RunRead(ByVal planOnly As Boolean) As String
    Dim wb As Workbook, pj As String, tag As String, t As String, lc As Object, heads As Collection, src As String
    Dim keys As String, h As Variant, res As Object, s As String, k As Variant, envs As Object, miss As String, envUse As String
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
    Logp "Civil NX: piles (" & PILE_GROUP & ", elements, nodes)"
    Set heads = PileHeads(src)
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
            IIf(miss = "", "", "Not in this model: " & miss & " - their columns are kept from the last read when the piles are the same." & vbCrLf) & vbCrLf & _
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
' items Array(headElem, headNode, x, y, "I" | "J"), sorted by element; src = where the list came from.
' Today: the beam elements of structure group KAZIK (MIDAS_KENARAYAK), one chain per pile, head = the
' upper end of the chain. A new foundation builder replaces only this function.
Private Function PileHeads(ByRef src As String) As Collection
    Dim gr As Object, k As Variant, el As Variant, ev As Object, nd As Object, inSet As Object, byI As Object, jSet As Object
    Dim e As Variant, rec As Object, tops() As Long, n As Long, i As Long, cur As Long, nI As Long, nJ As Long, first As Long, steps As Long
    Dim zt As Double, zb As Double, out As Collection, nodes As Variant
    Set gr = TableGet("GRUP")
    If gr Is Nothing Then Err.Raise vbObjectError + 7, , "The open model has no structure groups (db/GRUP)."
    For Each k In gr.keys
        If UCase$(CStr(gr(k)("NAME"))) = PILE_GROUP Then
            If gr(k).Exists("E_LIST") Then el = gr(k)("E_LIST")
            Exit For
        End If
    Next k
    If IsEmpty(el) Then Err.Raise vbObjectError + 7, , "Structure group " & PILE_GROUP & " not found or without elements: not a piled abutment model of MIDAS_KENARAYAK."
    Set ev = TableGet("ELEM")
    Set nd = TableGet("NODE")
    If ev Is Nothing Or nd Is Nothing Then Err.Raise vbObjectError + 7, , "Elements or nodes could not be read."
    Set inSet = CreateObject("Scripting.Dictionary")
    Set byI = CreateObject("Scripting.Dictionary")
    Set jSet = CreateObject("Scripting.Dictionary")
    For Each e In el
        If Not ev.Exists(CStr(CLng(e))) Then Err.Raise vbObjectError + 7, , "Element " & CLng(e) & " of " & PILE_GROUP & " is not in the model."
        Set rec = ev(CStr(CLng(e)))
        nodes = rec("NODE")
        nI = CLng(NumOr0(nodes(1))): nJ = CLng(NumOr0(nodes(2)))
        If nI > 0 And nJ > 0 And CLng(NumOr0(nodes(3))) = 0 Then       ' line elements only
            inSet(CStr(CLng(e))) = Array(nI, nJ)
            If byI.Exists(CStr(nI)) Then Err.Raise vbObjectError + 7, , "Two pile elements start at node " & nI & ": the piles cannot be followed."
            byI(CStr(nI)) = CLng(e)
            jSet(CStr(nJ)) = True
        End If
    Next e
    If inSet.Count = 0 Then Err.Raise vbObjectError + 7, , "Structure group " & PILE_GROUP & " has no line elements."
    ReDim tops(1 To inSet.Count)
    For Each k In inSet.keys
        If Not jSet.Exists(CStr(inSet(k)(0))) Then n = n + 1: tops(n) = CLng(k)
    Next k
    If n = 0 Then Err.Raise vbObjectError + 7, , "No pile end found in " & PILE_GROUP & " (the elements form loops)."
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
    Set PileHeads = SortByElem(out)
    src = "structure group " & PILE_GROUP & ", " & n & " element chains, head = upper end"
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
        same = (CLng(NumOr0(McAnchor(uc, "M_KAZIK_ADET", "B3").Value)) = n)
        i = 0
        For Each h In heads
            i = i + 1
            If same Then same = (CLng(NumOr0(oldT(i, 2))) = CLng(h(0)))
        Next h
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

' ---- report pages ---------------------------------------------------------------------------
Private Function PagesRefresh(wb As Workbook) As String
    Dim uc As Worksheet, rp As Worksheet, gs As Worksheet, n As Long, lg As String, r As Variant
    Set uc = SheetOf(wb, SH_UC)
    If uc Is Nothing Then Err.Raise vbObjectError + 20, , "This workbook has no sheet " & SH_UC & "."
    n = CLng(NumOr0(McAnchor(uc, "M_KAZIK_ADET", "B3").Value))
    If n < 0 Or n > MAX_PILES Then Err.Raise vbObjectError + 20, , SH_UC & ": pile count " & n & " outside 0.." & MAX_PILES & "."
    Application.ScreenUpdating = False
    Set rp = SheetOf(wb, SH_REP)
    If rp Is Nothing Then
        lg = lg & "No sheet " & SH_REP & ": its tables were not resized." & vbCrLf
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
    Dim nPages As Long, k As Long, top As Long, f As String, pa As Range, addr As String, changed As Boolean, lg As String
    PageTops rp, t1, t2, t3
    have1 = (t2 - t1) \ PG: have2 = (t3 - t2) \ PG
    need1 = PagesFor(n, PER1): need2 = PagesFor(n, PER2)
    after = PagesAfter(rp, t3)
    If need2 <> have2 Then ChangePages rp, t2, t3, need2 - have2, True: changed = True
    If need1 <> have1 Then ChangePages rp, t1, t2, need1 - have1, False: changed = True
    PageTops rp, t1, t2, t3
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
    lg = SH_REP & ": per-pile table " & need1 & " page(s)" & IIf(need1 <> have1, " (was " & have1 & ")", "") & _
         ", 8.6.7 " & need2 & " page(s)" & IIf(need2 <> have2, " (was " & have2 & ")", "") & ", " & nPages & " pages, page numbers chained." & vbCrLf
    ReportPages = lg
End Function

Private Sub PageTops(rp As Worksheet, ByRef t1 As Long, ByRef t2 As Long, ByRef t3 As Long)
    t1 = McAnchor(rp, "M_KAZIK_SAYFA1", "A186:A231", True).Row
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
