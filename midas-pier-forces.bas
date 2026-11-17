Attribute VB_Name = "MIDAS_AYAK"
Option Explicit
' ============================================================================
'  MIDAS_AYAK - pier column forces of the model open in Civil NX into the data sheet of each pier
'  group of the 8'7 workbook (MIDAS_<pier runs>, made by ORTA_AYAK from SABLON_MIDAS, found by its group number
'  in the sheet-level name M_TESIR_NO; an older sheet MIDAS_<group number> is still found). Owner 2026-11-16.
'
'  Macro window (MAKROLAR) > MIDAS sonuclari > Ayak tesirleri, or MIDAS_AYAK_OKU.
'  Reads the model OPEN in Civil NX through the MAPI; it writes nothing to the model.
'  Model kind from its combinations (db/LCOM-GEN; single combinations, envelopes are not read):
'     static model   Servis * and STR-* (a moving load or an envelope combination inside gives a (max) and a
'                    (min) row; a combination that answers nothing as (CB) is read again as (CB:max) / (CB:min))
'     seismic model  TKDY-EQ-* (owner 2026-11-16: not H-TKDY, not ENV_DEP)
'  Run it once with each model open. A run replaces its own model's rows and keeps the other model's
'  rows when they were read for the same piers (stamps M_TESIR_STATIK, M_TESIR_DEPREM, M_TESIR_GRUP on
'  the data sheet); other rows (the reference rows of the template) are dropped.
'  Piers: MIDAS_KOPRU puts every pier element in structure group SG_KOLON and builds the piers in the
'  order P1, P2 ..., Kolon Sayisi columns each (G_AYAK_KOLON, empty = 1), every column from the cap
'  bottom down. A column = a chain of elements (J node of one = I node of the next); the chains in
'  element order are the columns of P1, P2 ... Each column length must equal the pier height of the
'  pier table (Yukseklik) within H_TOL, else the run stops: the model is not the one of this table.
'  Rows (owner 2026-11-13 / 2026-11-16, choice 1a): every pier of the group, each column, each
'  combination:
'     column base  J end of the lowest element, every pier
'     column top   I end of the highest element, only for piers with 2 or more columns (frame in Y)
'  Name <combination><(max)|(min)>_<pier>[K<column>]_<J|I>, e.g. 'Servis 1(max)_P3_J', 'TKDY-EQ-X-1_P3K2_I'.
'  Values (kN, kN.m): Pu = -Axial (compression +), My = Moment-y (long direction), Mz = Moment-z
'  (short), Vy = Shear-y (short), Vz = Shear-z (long), T = Torsion (owner 2026-11-16: My long, Mz
'  short, Fz long shear, Fy short shear; measured on the seismic model: EQ-X gives My and Vz).
'  Written to the block M_TESIR_BLOK of MIDAS_<group> (default CW26:DE3839: counter, No, Ad, Pu, My, Mz, Vy, Vz,
'  T) = the block that Tablo 63
'  'Ayak Tesirleri', the shear picks and the torsion pick of the group sheet read; static rows first,
'  then seismic (the EQ rows appended to the static rows, owner 2026-11-16). Tablo 63 grows with the rows:
'  ORTA_AYAK.SAYFALARI_YENILE_SESSIZ adds or removes its pages (ORTA_AYAK 2026-11-16e). Element list for later use (GSD) in M_TESIR_ELEMAN (DF6:DG400).
'  Then ORTA_AYAK.SAYFALARI_YENILE_SESSIZ (gates, print areas, page numbers, lists).
'  MAPI key: read by MIDAS_CORE (G_MAPI_KEY, INPUT!T3, MIDAS_MAPI_KEY, asked; memory only).
'  Test entries: MIDAS_AYAK_RUN <report file> (no questions), MIDAS_AYAK_PLAN (piers, elements,
'  combinations and requests of the open model; nothing is read from the result tables or written).
' ============================================================================

Private Const SCRIPT_ID As String = "midas-pier-forces"
Private Const SCRIPT_VERSION As String = "2026-11-16c"
Private Const SCRIPT_CHANGELOG As String = "Tablo 63 grows with the rows (ORTA_AYAK 2026-11-16e makes its pages), no row-limit note; combinations holding an envelope (Servis 4, STR-3, STR-4 through IVME/FREN) read as (max) / (min); data sheet found by its group number (M_TESIR_NO), named after the piers; new: column forces of every pier of each group (base, and top for 2 or more columns) from the open static or seismic model into the group's MIDAS_<group> sheet, which Tablo 63 reads"
Private Const SCRIPT_NEEDS As String = "midas-core"   ' UPDATE adds this module when the workbook lacks it
' macro-window row(s) of this module (PANEL reads them): order|group|title|macro|sheet|text|fields[|modules]
'@TOOL 85|MIDAS sonu\u00E7lar\u0131|Ayak tesirleri|MIDAS_AYAK_OKU||Civil NX'te a\u00E7\u0131k statik veya sismik modelden her ayak grubunun ayaklar\u0131n\u0131n kolon tesirlerini (taban; 2 ve daha \u00E7ok kolonda \u00FCst u\u00E7 da) grubun MIDAS_<grup> sayfas\u0131na okur; Tablo 63 bunlar\u0131 g\u00F6sterir. Her model i\u00E7in bir kez \u00E7al\u0131\u015Ft\u0131r\u0131n. Modele bir \u015Fey yazmaz.|

Private Const MSG_TITLE As String = "MIDAS_AYAK"
Private Const DATA_PFX As String = "MIDAS_"
' force block of a data sheet (sheet-level name M_TESIR_BLOK, default CW26:DE3839 = the rows its formulas
' B63:B3876 read): counter, No, Ad, Pu, My, Mz, Vy, Vz, T
Private Const M_BLOCK As String = "M_TESIR_BLOK"
Private Const H_TOL As Double = 0.05           ' column length against the pier height (m)
Private Const PIER_GROUP As String = "SG_KOLON"
Private Const N_TABLE As String = "G_AYAK_TABLOSU"
Private Const N_COLS As String = "G_AYAK_KOLON"
Private Const N_GROUP As String = "OA_GRUP_NO"
Private Const N_GTAB As String = "OA_GRUP_TABLO"
Private Const KIND_S As String = "S", KIND_D As String = "D"

Private mScripted As Boolean
Private mLog As String

' ============================================================================
Public Sub MIDAS_AYAK_OKU()
    Dim s As String
    mScripted = False
    s = RunRead(False)
    UiMsg s, IIf(Left$(s, 8) = "STOPPED:", vbExclamation, vbInformation), MSG_TITLE
End Sub

Public Sub MIDAS_AYAK_RUN(ByVal reportFile As String)
    mScripted = True
    WriteText reportFile, RunRead(False)
End Sub

Public Function MIDAS_AYAK_PLAN() As String
    mScripted = True
    MIDAS_AYAK_PLAN = RunRead(True)
End Function

' ============================================================================
Private Function RunRead(ByVal planOnly As Boolean) As String
    Dim wb As Workbook, pj As String, tag As String, t As Variant, g As Variant, kc As Variant, hasCols As Boolean
    Dim pRow As Object, i As Long, nm As String, nG As Long, k As Long, rc As Range
    Dim cols As Object, chains As Collection, pierOf As Object, lc As Object, combos As Collection
    Dim hasS As Boolean, hasD As Boolean, c As Variant, keys As String, elems As Object, res As Object
    Dim grpPiers() As Collection, used As Long, s As String, sumTxt As String, r As Variant, nc As Long, two As Object
    On Error GoTo Fail
    mLog = ""
    Set wb = ThisWorkbook

    ' ---- the workbook: pier table, groups, columns ----
    t = NmRange(wb, N_TABLE).Value
    g = NmRange(wb, N_GROUP).Value
    If UBound(g, 1) <> UBound(t, 1) Then Err.Raise vbObjectError + 1, , N_GROUP & " and " & N_TABLE & " differ in rows."
    nG = NmRange(wb, N_GTAB).Rows.Count
    On Error Resume Next
    Set rc = wb.Names(N_COLS).RefersToRange
    On Error GoTo Fail
    If Not rc Is Nothing Then
        kc = rc.Value
        hasCols = (UBound(kc, 1) = UBound(t, 1))
    End If
    Set pRow = CreateObject("Scripting.Dictionary")      ' pier number -> table row
    For i = 1 To UBound(t, 1)
        nm = UCase$(CellTxt(t(i, 1)))
        If Len(nm) > 1 And Left$(nm, 1) = "P" Then
            If Mid$(nm, 2) Like String$(Len(nm) - 1, "#") Then pRow(CStr(CLng(Mid$(nm, 2)))) = i
        End If
    Next i
    ReDim grpPiers(1 To nG)
    For k = 1 To nG: Set grpPiers(k) = New Collection: Next k
    For i = 1 To UBound(t, 1)
        k = GroupNo(g(i, 1), nG)
        If k > 0 Then
            nm = UCase$(CellTxt(t(i, 1)))
            If Len(nm) < 2 Or Left$(nm, 1) <> "P" Then Err.Raise vbObjectError + 2, , "'" & CellTxt(t(i, 1)) & "' has a group number but is not a pier P<n>."
            If Not (Mid$(nm, 2) Like String$(Len(nm) - 1, "#")) Then Err.Raise vbObjectError + 2, , "'" & CellTxt(t(i, 1)) & "' has a group number but is not a pier P<n>."
            nc = 1
            If hasCols Then nc = ColCount(kc(i, 1))
            grpPiers(k).Add Array(nm, CLng(Mid$(nm, 2)), nc, i)
            used = used + 1
        End If
    Next i
    If used = 0 Then Err.Raise vbObjectError + 3, , "No pier has a group number (INPUT, pier table, Grup No)."

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
    If lc Is Nothing Then Err.Raise vbObjectError + 4, , "The open model has no load combinations (db/LCOM-GEN)."
    Set combos = ReadCombos(lc, hasS, hasD)
    If combos.Count = 0 Then Err.Raise vbObjectError + 5, , "The open model (" & tag & ") has no Servis / STR or TKDY-EQ combination: open the static or the seismic MIDAS_KOPRU model."
    If UiStopped() Then RunRead = "STOPPED: by the user.": Exit Function
    Logp "Civil NX: pier columns (" & PIER_GROUP & ", elements, nodes)"
    Set chains = ReadChains()
    Set pierOf = MapPiers(chains, pRow, t, kc, hasCols, grpPiers, nG)
    If UiStopped() Then RunRead = "STOPPED: by the user.": Exit Function

    ' elements to read: base of every grouped pier column, top when the pier has 2 or more columns
    Set elems = CreateObject("Scripting.Dictionary")
    For k = 1 To nG
        For Each c In grpPiers(k)
            For i = 1 To c(2)
                r = pierOf(PcKey(c(1), i))
                elems(CStr(r(1))) = True
                If c(2) >= 2 Then elems(CStr(r(0))) = True
            Next i
        Next c
    Next k
    keys = JoinKeys(elems)

    If planOnly Then
        RunRead = PlanText(tag, combos, grpPiers, nG, pierOf, keys, wb)
        Exit Function
    End If

    If Not mScripted Then
        s = "Model: " & tag & vbCrLf & "Combinations: " & combos.Count & IIf(hasS, " static (Servis, STR)", "") & IIf(hasD, " seismic (TKDY-EQ)", "") & vbCrLf & _
            "Piers: " & used & " in " & GroupCount(grpPiers, nG) & " groups, " & elems.Count & " column ends." & vbCrLf & vbCrLf & _
            "The forces of this model replace its rows in the MIDAS_<group> sheets. Continue?"
        If UiMsg(s, vbOKCancel + vbQuestion, MSG_TITLE) <> vbOK Then RunRead = "STOPPED: nothing was read.": Exit Function
    End If

    ' ---- forces ----
    Set res = CreateObject("Scripting.Dictionary")
    Set two = ReadForces(keys, combos, res)
    If UiStopped() Then RunRead = "STOPPED: by the user (nothing was written).": Exit Function

    ' ---- write ----
    Application.ScreenUpdating = False
    For k = 1 To nG
        If grpPiers(k).Count > 0 Then sumTxt = sumTxt & WriteGroup(wb, k, grpPiers(k), combos, two, pierOf, res, hasS, hasD, tag) & vbCrLf
    Next k
    Application.Calculate
    On Error Resume Next
    r = Application.Run("'" & Replace(wb.Name, "'", "''") & "'!ORTA_AYAK.SAYFALARI_YENILE_SESSIZ")
    If Err.Number <> 0 Then
        mLog = mLog & "Pages not refreshed (ORTA_AYAK.SAYFALARI_YENILE_SESSIZ): " & Err.Description & vbCrLf
        Err.Clear
    ElseIf Left$(CStr(r), 5) = "FAIL:" Then
        mLog = mLog & "Pages: " & CStr(r) & vbCrLf
    Else
        mLog = mLog & "Pages, page numbers and lists refreshed." & vbCrLf
    End If
    On Error GoTo Fail
    Application.ScreenUpdating = True
    RunRead = "Done. Model: " & tag & vbCrLf & IIf(hasS, "static (Servis, STR) ", "") & IIf(hasD, "seismic (TKDY-EQ) ", "") & combos.Count & " combinations." & vbCrLf & sumTxt & mLog
    Exit Function
Fail:
    Application.ScreenUpdating = True
    RunRead = "STOPPED: " & Err.Description & IIf(Len(mLog) > 0, vbCrLf & mLog, "")
End Function

' single combinations of the static (Servis *, STR-*) and seismic (TKDY-EQ-*) kinds in model order:
' items Array(name, hasMovingLoad, kind)
Private Function ReadCombos(lc As Object, ByRef hasS As Boolean, ByRef hasD As Boolean) As Collection
    Dim byName As Object, ks As Variant, i As Long, rec As Object, nm As String, kind As String
    Set byName = CreateObject("Scripting.Dictionary")
    ks = SortedNumKeys(lc)
    For i = 1 To UBound(ks)
        Set rec = lc(ks(i))
        Set byName(CStr(rec("NAME"))) = rec
    Next i
    Set ReadCombos = New Collection
    For i = 1 To UBound(ks)
        Set rec = lc(ks(i))
        nm = CStr(rec("NAME"))
        kind = ""
        If CLng(NumOr0(rec("iTYPE"))) = 0 Then
            If nm Like "Servis *" Or nm Like "STR-*" Then kind = KIND_S: hasS = True
            If nm Like "TKDY-EQ-*" Then kind = KIND_D: hasD = True
        End If
        If Len(kind) > 0 Then ReadCombos.Add Array(nm, HasMV(byName, nm, 0), kind)
    Next i
End Function

' True when the combination's results come as (max) and (min): it holds a moving-load case or an envelope
' combination (iTYPE 1), directly or through another combination (measured 2026-11-16 on the static model:
' Servis 4, STR-3, STR-4 hold IVME/FREN, which holds the envelopes IVME and FREN; '(CB)' answers nothing)
Private Function HasMV(byName As Object, ByVal nm As String, ByVal depth As Long) As Boolean
    Dim rec As Object, it As Variant, j As Long, sub1 As String
    If depth > 10 Or Not byName.Exists(nm) Then Exit Function
    Set rec = byName(nm)
    If Not rec.Exists("vCOMB") Then Exit Function
    it = rec("vCOMB")
    If IsEmpty(it) Then Exit Function
    For j = 1 To UBound(it)
        Select Case UCase$(CStr(it(j)("ANAL")))
            Case "MV": HasMV = True: Exit Function
            Case "CB"
                sub1 = CStr(it(j)("LCNAME"))
                If byName.Exists(sub1) Then
                    If CLng(NumOr0(byName(sub1)("iTYPE"))) = 1 Then HasMV = True: Exit Function
                End If
                If HasMV(byName, sub1, depth + 1) Then HasMV = True: Exit Function
        End Select
    Next j
End Function

' columns of the model: chains of SG_KOLON elements from the top down, in element order;
' items Array(topElem, baseElem, length m, topNode, baseNode)
Private Function ReadChains() As Collection
    Dim gr As Object, k As Variant, el As Variant, ev As Object, nd As Object, inSet As Object, byI As Object, jSet As Object
    Dim e As Variant, tops() As Long, n As Long, i As Long, cur As Long, nI As Long, nJ As Long, first As Long, zt As Double, zb As Double, steps As Long
    Set gr = TableGet("GRUP")
    If gr Is Nothing Then Err.Raise vbObjectError + 6, , "The open model has no structure groups (db/GRUP)."
    For Each k In gr.keys
        If CStr(gr(k)("NAME")) = PIER_GROUP Then el = gr(k)("E_LIST"): Exit For
    Next k
    If IsEmpty(el) Then Err.Raise vbObjectError + 6, , "Structure group " & PIER_GROUP & " not found or empty: not a MIDAS_KOPRU model."
    Set ev = TableGet("ELEM")
    Set nd = TableGet("NODE")
    If ev Is Nothing Or nd Is Nothing Then Err.Raise vbObjectError + 6, , "Elements or nodes could not be read."
    Set inSet = CreateObject("Scripting.Dictionary")
    Set byI = CreateObject("Scripting.Dictionary")
    Set jSet = CreateObject("Scripting.Dictionary")
    For Each e In el
        If Not ev.Exists(CStr(CLng(e))) Then Err.Raise vbObjectError + 6, , "Element " & CLng(e) & " of " & PIER_GROUP & " is not in the model."
        nI = CLng(ev(CStr(CLng(e)))("NODE")(1)): nJ = CLng(ev(CStr(CLng(e)))("NODE")(2))
        inSet(CStr(CLng(e))) = Array(nI, nJ)
        If byI.Exists(CStr(nI)) Then Err.Raise vbObjectError + 6, , "Two pier elements start at node " & nI & ": the columns cannot be followed."
        byI(CStr(nI)) = CLng(e)
        jSet(CStr(nJ)) = True
    Next e
    ReDim tops(1 To inSet.Count)
    For Each k In inSet.keys
        If Not jSet.Exists(CStr(inSet(k)(0))) Then n = n + 1: tops(n) = CLng(k)
    Next k
    If n = 0 Then Err.Raise vbObjectError + 6, , "No column top found in " & PIER_GROUP & "."
    ReDim Preserve tops(1 To n)
    SortLongs tops
    Set ReadChains = New Collection
    For i = 1 To n
        first = tops(i): cur = first: steps = 0
        Do
            nJ = inSet(CStr(cur))(1)
            If Not byI.Exists(CStr(nJ)) Then Exit Do
            cur = byI(CStr(nJ))
            steps = steps + 1
            If steps > inSet.Count Then Err.Raise vbObjectError + 6, , "The pier elements from " & first & " form a loop."
        Loop
        zt = NodeZ(nd, inSet(CStr(first))(0))
        zb = NodeZ(nd, inSet(CStr(cur))(1))
        ReadChains.Add Array(first, cur, zt - zb, inSet(CStr(first))(0), inSet(CStr(cur))(1))
    Next i
End Function

' pier P<p> column cc -> its chain; the chains in element order are the columns of P1, P2 ...
' Stops when the counts or the heights do not match the pier table, or a grouped pier is not in the model.
Private Function MapPiers(chains As Collection, pRow As Object, t As Variant, kc As Variant, ByVal hasCols As Boolean, _
                          grpPiers() As Collection, ByVal nG As Long) As Object
    Dim idx As Long, p As Long, nc As Long, cc As Long, ch As Variant, h As Double, bad As String, nBad As Long, k As Long, c As Variant, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    idx = 1: p = 0
    Do While idx <= chains.Count
        p = p + 1
        If Not pRow.Exists(CStr(p)) Then Err.Raise vbObjectError + 7, , "The model has more pier columns than the pier table holds (no row P" & p & ")."
        nc = 1
        If hasCols Then nc = ColCount(kc(pRow(CStr(p)), 1))
        For cc = 1 To nc
            If idx > chains.Count Then Err.Raise vbObjectError + 7, , "P" & p & " has " & nc & " columns in the pier table, the model has fewer: not the model of this table."
            ch = chains(idx)
            h = NumOr0(t(pRow(CStr(p)), 2))
            If Abs(ch(2) - h) > H_TOL Then
                nBad = nBad + 1
                If nBad <= 6 Then bad = bad & "  P" & p & IIf(nc > 1, " K" & cc, "") & ": model " & Format$(ch(2), "0.00") & " m, table " & Format$(h, "0.00") & " m" & vbCrLf
            End If
            d(PcKey(p, cc)) = ch
            idx = idx + 1
        Next cc
    Loop
    If nBad > 0 Then Err.Raise vbObjectError + 8, , "Column lengths of the model differ from the pier heights of the table (" & nBad & " columns):" & vbCrLf & bad & "Open the model built from this pier table."
    For k = 1 To nG
        For Each c In grpPiers(k)
            If Not d.Exists(PcKey(c(1), 1)) Then Err.Raise vbObjectError + 9, , c(0) & " (group " & k & ") is not in the model: it has " & p & " piers."
        Next c
    Next k
    Set MapPiers = d
End Function

' Beam Force table of the elements for every combination: res("<elem>|<I|J>|<combination>|<suffix>") =
' Array(Pu, My, Mz, Vy, Vz, T); a (max) / (min) combination is read twice, (CB:max) and (CB:min); a combination
' read as (CB) that answers no row is read as (CB:max) / (CB:min) too. Returns the (max) / (min) combinations.
Private Function ReadForces(ByVal keys As String, combos As Collection, res As Object) As Object
    Dim plain As String, mv As String, c As Variant, got As Object, two As Object
    Set two = CreateObject("Scripting.Dictionary")
    For Each c In combos
        If c(1) Then
            mv = mv & "|" & c(0): two(CStr(c(0))) = True
        Else
            plain = plain & "|" & c(0)
        End If
    Next c
    If Len(plain) > 0 Then
        Logp "Civil NX: beam forces"
        Set got = ReadTable(keys, Mid$(plain, 2), "(CB)", "", res, True)
        For Each c In combos
            If Not c(1) Then
                If Not got.Exists(CStr(c(0))) Then
                    mv = mv & "|" & c(0): two(CStr(c(0))) = True
                    mLog = mLog & c(0) & ": no (CB) result, read as (max) / (min)." & vbCrLf
                End If
            End If
        Next c
    End If
    If Len(mv) > 0 Then
        Logp "Civil NX: beam forces (max)"
        ReadTable keys, Mid$(mv, 2), "(CB:max)", "(max)", res, False
        Logp "Civil NX: beam forces (min)"
        ReadTable keys, Mid$(mv, 2), "(CB:min)", "(min)", res, False
    End If
    Set ReadForces = two
End Function

' returns the combinations that answered at least one row; allowEmpty: an empty table is no error
Private Function ReadTable(ByVal keys As String, ByVal cbList As String, ByVal reqSfx As String, ByVal sfx As String, res As Object, ByVal allowEmpty As Boolean) As Object
    Dim nm As Variant, lcs As String, body As String, resp As String, root As Object, tb As Object, head As Variant, data As Variant
    Dim iE As Long, iL As Long, iP As Long, iA As Long, iVy As Long, iVz As Long, iT As Long, iMy As Long, iMz As Long
    Dim j As Long, r As Long, ld As String, part As String, want As Object, got As Object
    Set want = CreateObject("Scripting.Dictionary")
    Set got = CreateObject("Scripting.Dictionary")
    Set ReadTable = got
    For Each nm In Split(cbList, "|")
        lcs = lcs & IIf(lcs = "", "", ",") & QS(CStr(nm) & reqSfx)
        want(CStr(nm)) = True
    Next nm
    body = "{""Argument"":{""TABLE_NAME"":""F"",""TABLE_TYPE"":""BEAMFORCE"",""UNIT"":{""FORCE"":""KN"",""DIST"":""M""}," & _
           """NODE_ELEMS"":{""KEYS"":[" & keys & "]},""LOAD_CASE_NAMES"":[" & lcs & "],""PARTS"":[""PartI"",""PartJ""]}}"
    If Not Api("POST", "post/TABLE", body, resp) Then Err.Raise vbObjectError + 10, , "Beam forces " & reqSfx & ": " & Left$(resp, 300)
    Set root = ParseJson(resp)
    If root Is Nothing Then Err.Raise vbObjectError + 10, , "Beam forces " & reqSfx & ": no answer."
    If Not root.Exists("F") Then
        If allowEmpty Then Exit Function
        Err.Raise vbObjectError + 10, , "Beam forces " & reqSfx & ": no table (has the analysis been run?). " & Left$(resp, 200)
    End If
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
    If iE * iL * iP * iA * iVy * iVz * iT * iMy * iMz = 0 Then Err.Raise vbObjectError + 10, , "Beam forces " & reqSfx & ": unexpected columns."
    If IsEmpty(data) Then
        If allowEmpty Then Exit Function
        Err.Raise vbObjectError + 10, , "Beam forces " & reqSfx & ": empty table (has the analysis been run?)."
    End If
    For r = 1 To UBound(data)
        ld = CStr(data(r)(iL))
        If Not want.Exists(ld) Then ld = BaseName(ld)
        If want.Exists(ld) Then
            got(ld) = True
            part = Left$(CStr(data(r)(iP)), 1)
            res(CStr(CLng(Val(CStr(data(r)(iE))))) & "|" & part & "|" & ld & "|" & sfx) = Array( _
                -Val(CStr(data(r)(iA))), Val(CStr(data(r)(iMy))), Val(CStr(data(r)(iMz))), _
                Val(CStr(data(r)(iVy))), Val(CStr(data(r)(iVz))), Val(CStr(data(r)(iT))))
        End If
    Next r
End Function

' rows of group k into MIDAS_<k>; returns its summary line
Private Function WriteGroup(wb As Workbook, ByVal k As Long, piers As Collection, combos As Collection, two As Object, pierOf As Object, res As Object, _
                            ByVal hasS As Boolean, ByVal hasD As Boolean, ByVal tag As String) As String
    Dim ds As Worksheet, stS As Range, stD As Range, stG As Range, gName As String, keepS As Boolean, keepD As Boolean
    Dim oldV As Variant, i As Long, rowsS As Collection, rowsD As Collection, rw As Variant, n As Long, out() As Variant, j As Long
    Dim c As Variant, p As Variant, cc As Long, locs As Variant, lc As Variant, sfx As Variant, sfxs As Variant, ch As Variant, key As String, v As Variant
    Dim blk As Range, el As Range, used As Long
    Set ds = DataSheetOf(wb, k)
    If ds Is Nothing Then Err.Raise vbObjectError + 11, , "No data sheet for group " & k & ": run 'Ayak grubu sayfalari' (ORTA_AYAK) first."
    Set stS = McAnchor(ds, "M_TESIR_STATIK", "DG2", True)
    Set stD = McAnchor(ds, "M_TESIR_DEPREM", "DG3", True)
    Set stG = McAnchor(ds, "M_TESIR_GRUP", "DG4", True)
    Set blk = McAnchor(ds, M_BLOCK, "CW26:DE3839", True)
    Set el = McAnchor(ds, "M_TESIR_ELEMAN", "DF6:DG400", True)
    If blk.Columns.Count <> 9 Then Err.Raise vbObjectError + 11, , ds.Name & ": " & M_BLOCK & " must be 9 columns wide."
    If Len(CellTxt(stS.Offset(0, -1).Value)) = 0 Then stS.Offset(0, -1).Value = "Statik model"
    If Len(CellTxt(stD.Offset(0, -1).Value)) = 0 Then stD.Offset(0, -1).Value = "Deprem modeli"
    If Len(CellTxt(stG.Offset(0, -1).Value)) = 0 Then stG.Offset(0, -1).Value = "Ayaklar"
    gName = CellTxt(NmRange(wb, N_GTAB).Cells(k, 2).Value)
    ' rows of the other model stay when they were read for these piers
    keepS = (Not hasS) And Len(CellTxt(stS.Value)) > 0 And StrComp(CellTxt(stG.Value), gName, vbTextCompare) = 0
    keepD = (Not hasD) And Len(CellTxt(stD.Value)) > 0 And StrComp(CellTxt(stG.Value), gName, vbTextCompare) = 0
    Set rowsS = New Collection: Set rowsD = New Collection
    ' rows in use: down to the last name in the block's third column
    For i = blk.Rows.Count To 1 Step -1
        If Len(CellTxt(blk.Cells(i, 3).Value)) > 0 Then used = i: Exit For
    Next i
    If used > 0 Then
        oldV = blk.Resize(used).Value
        For i = 1 To UBound(oldV, 1)
            If Len(CellTxt(oldV(i, 3))) > 0 Then
                rw = Array(oldV(i, 3), oldV(i, 4), oldV(i, 5), oldV(i, 6), oldV(i, 7), oldV(i, 8), oldV(i, 9))
                If Left$(CellTxt(oldV(i, 3)), 5) = "TKDY-" Then
                    If keepD Then rowsD.Add rw
                Else
                    If keepS Then rowsS.Add rw
                End If
            End If
        Next i
    End If
    If Not hasS And Not keepS And Len(CellTxt(stS.Value)) > 0 Then mLog = mLog & ds.Name & ": static rows dropped (read for " & CellTxt(stG.Value) & "): read the static model again." & vbCrLf: stS.Value = ""
    If Not hasD And Not keepD And Len(CellTxt(stD.Value)) > 0 Then mLog = mLog & ds.Name & ": seismic rows dropped (read for " & CellTxt(stG.Value) & "): read the seismic model again." & vbCrLf: stD.Value = ""
    ' new rows of this model: combination, pier, column, base / top, (max) / (min)
    For Each c In combos
        If two.Exists(CStr(c(0))) Then sfxs = Array("(max)", "(min)") Else sfxs = Array("")
        For Each p In piers
            For cc = 1 To p(2)
                ch = pierOf(PcKey(p(1), cc))
                If p(2) >= 2 Then locs = Array("J", "I") Else locs = Array("J")
                For Each lc In locs
                    For Each sfx In sfxs
                        key = CStr(IIf(lc = "J", ch(1), ch(0))) & "|" & lc & "|" & c(0) & "|" & sfx
                        If Not res.Exists(key) Then Err.Raise vbObjectError + 12, , "No force for " & c(0) & sfx & " at " & p(0) & IIf(p(2) > 1, " K" & cc, "") & " (" & key & ")."
                        v = res(key)
                        rw = Array(c(0) & sfx & "_" & p(0) & IIf(p(2) > 1, "K" & cc, "") & "_" & lc, _
                                   Round(v(0), 2), Round(v(1), 2), Round(v(2), 2), Round(v(3), 2), Round(v(4), 2), Round(v(5), 2))
                        If c(2) = KIND_S Then rowsS.Add rw Else rowsD.Add rw
                    Next sfx
                Next lc
            Next cc
        Next p
    Next c
    n = rowsS.Count + rowsD.Count
    If n > blk.Rows.Count Then Err.Raise vbObjectError + 13, , ds.Name & ": " & n & " rows, the sheet holds " & blk.Rows.Count & "."
    If n > 0 Then
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
        blk.Resize(n).Value = out
    End If
    If used > n Then blk.Rows(n + 1).Resize(used - n).ClearContents
    If hasS Then stS.Value = tag & " | " & Format$(Now, "yyyy-mm-dd hh:nn")
    If hasD Then stD.Value = tag & " | " & Format$(Now, "yyyy-mm-dd hh:nn")
    stG.Value = gName
    ' element list (for GSD later)
    el.ClearContents
    el.Cells(1, 1).Value = "Elemanlar"
    i = 0
    For Each p In piers
        For cc = 1 To p(2)
            ch = pierOf(PcKey(p(1), cc))
            i = i + 1
            If i <= el.Rows.Count Then el.Cells(i, 2).Value = p(0) & IIf(p(2) > 1, " K" & cc, "") & ": " & ch(0) & "-" & ch(1) & ", base " & ch(1) & " J" & IIf(p(2) > 1, ", top " & ch(0) & " I", "")
        Next cc
    Next p
    WriteGroup = ds.Name & ": " & n & " rows (static " & rowsS.Count & ", seismic " & rowsD.Count & ")"
End Function

' ---- plan (test): what would be read ----
Private Function PlanText(ByVal tag As String, combos As Collection, grpPiers() As Collection, ByVal nG As Long, pierOf As Object, ByVal keys As String, wb As Workbook) As String
    Dim s As String, c As Variant, k As Long, p As Variant, cc As Long, ch As Variant, n As Long, perPier As Long
    s = "PLAN. Model: " & tag & vbCrLf & "Combinations:"
    For Each c In combos
        s = s & " " & c(0) & IIf(c(1), "(max/min)", "") & "[" & c(2) & "]"
        perPier = perPier + IIf(c(1), 2, 1)
    Next c
    s = s & vbCrLf & "Elements: " & keys & vbCrLf
    For k = 1 To nG
        If grpPiers(k).Count > 0 Then
            n = 0
            s = s & "Group " & k & " (" & CellTxt(NmRange(wb, N_GTAB).Cells(k, 2).Value) & "):"
            For Each p In grpPiers(k)
                For cc = 1 To p(2)
                    ch = pierOf(PcKey(p(1), cc))
                    s = s & " " & p(0) & IIf(p(2) > 1, "K" & cc, "") & "=" & ch(0) & ".." & ch(1) & "(" & Format$(ch(2), "0.00") & "m)"
                    n = n + perPier * IIf(p(2) >= 2, 2, 1)
                Next cc
            Next p
            s = s & " -> " & n & " rows" & vbCrLf
        End If
    Next k
    PlanText = s
End Function

' ---- helpers ----
Private Function PcKey(ByVal p As Long, ByVal cc As Long) As String
    PcKey = p & "|" & cc
End Function

Private Function GroupNo(ByVal v As Variant, ByVal nG As Long) As Long
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or Not IsNumeric(v) Then Exit Function
    If CDbl(v) = Int(CDbl(v)) And CDbl(v) >= 1 And CDbl(v) <= nG Then GroupNo = CLng(v)
End Function

Private Function GroupCount(grpPiers() As Collection, ByVal nG As Long) As Long
    Dim k As Long
    For k = 1 To nG
        If grpPiers(k).Count > 0 Then GroupCount = GroupCount + 1
    Next k
End Function

Private Function ColCount(ByVal v As Variant) As Long
    ColCount = 1
    If IsError(v) Then Exit Function
    If IsNumeric(v) And Not IsEmpty(v) Then
        If CDbl(v) >= 2 Then ColCount = CLng(v)
    End If
End Function

Private Function CellTxt(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    CellTxt = Trim$(CStr(v))
End Function

' "Servis 1(max)" -> "Servis 1"
Private Function BaseName(ByVal s As String) As String
    Dim p As Long
    BaseName = s
    If Right$(s, 1) <> ")" Then Exit Function
    p = InStrRev(s, "(")
    If p > 1 Then BaseName = Left$(s, p - 1)
End Function

Private Function NodeZ(nd As Object, ByVal n As Long) As Double
    If Not nd.Exists(CStr(n)) Then Err.Raise vbObjectError + 14, , "Node " & n & " is not in the model."
    NodeZ = CDbl(nd(CStr(n))("Z"))
End Function

Private Function JoinKeys(d As Object) As String
    Dim a() As Long, k As Variant, n As Long, i As Long
    ReDim a(1 To d.Count)
    For Each k In d.keys: n = n + 1: a(n) = CLng(k): Next k
    SortLongs a
    For i = 1 To n: JoinKeys = JoinKeys & IIf(i = 1, "", ",") & a(i): Next i
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

Private Function SortedNumKeys(d As Object) As Variant
    Dim a() As Long, k As Variant, n As Long, out() As String, i As Long
    ReDim a(1 To d.Count)
    For Each k In d.keys: n = n + 1: a(n) = CLng(Val(CStr(k))): Next k
    SortLongs a
    ReDim out(1 To n)
    For i = 1 To n: out(i) = CStr(a(i)): Next i
    SortedNumKeys = out
End Function

' the data sheet of group k: sheet-level name M_TESIR_NO = k (ORTA_AYAK 2026-11-16d), else an older sheet MIDAS_<k>
Private Function DataSheetOf(wb As Workbook, ByVal k As Long) As Worksheet
    Dim ws As Worksheet, nm As Name, old As Worksheet
    For Each ws In wb.Worksheets
        Set nm = Nothing
        On Error Resume Next
        Set nm = ws.Names("M_TESIR_NO")
        On Error GoTo 0
        If Not nm Is Nothing Then
            If nm.RefersTo = "=" & k Then Set DataSheetOf = ws: Exit Function
        ElseIf StrComp(ws.Name, DATA_PFX & k, vbTextCompare) = 0 Then
            Set old = ws
        End If
    Next ws
    Set DataSheetOf = old
End Function

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
