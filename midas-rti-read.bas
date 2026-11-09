Attribute VB_Name = "MIDAS_RTI"
Option Explicit
' ============================================================================
'  MIDAS_RTI - results of the RTI model (MIDAS_KOPRU, rail-track interaction) into
'  sheet RTI_SONUC of the main workbook; 3B_RTI K166, K167, J207:J211 read it.
'
'  Macro window (MAKROLAR) > MIDAS sonuclari > RTI sonuclari, or MIDAS_RTI_OKU.
'  Reads the model OPEN in Civil NX through the MAPI, read only: GET db/PJCF, GRUP,
'  NODE, ELEM, STAG and post/TABLE. Nothing is written to the model.
'
'  What it reads (groups made by MIDAS_KOPRU 2026-11-09b and later):
'     SG_RAY_t  rail elements and rail nodes of track t,  SG_IZ_t  track nodes,
'     DE_<axis>_L / _R  deck end node left / right of each axis (A1 only R, A2 only L).
'  Stages: every stage except TU+, TU-, END and LL_* is a braking / traction window.
'  Results are Summation(CS) at the last step of a stage: the per-type CS results are
'  not used (measured 2026-11-09: they mix after the window stages).
'
'  D5 / D6 rail stress (old 3B Excel rule, per rail element end):
'       window min + TU+ or TU- + LM71 min,  window max + LM71 max + TU+ or TU-
'       D5 = largest of the four sums (tension +), D6 = smallest (compression -);
'       Beam Stress table, axial column (N / A), kN/m2 / 1000 = N/mm2.
'  D7 dB   per window, deck end movement along the track: |left - right| at a pier,
'          |end| at an abutment (ground fixed); largest.
'  D8 dH   LM71 moving load, same difference from the envelopes (upper bound: the
'          max of one side against the min of the other).
'  D9 dV   same with DZ.
'  D10 du  per window, rail node minus its track node along the track; largest.
'  D11 dabs per window, deck end movement along the track; largest.
'  Along the track = horizontal DX, DY projected on the rail direction at the point.
'  Track type (G_HAT_TIPI): D10 and D11 only on ballasted track (owner 2026-11-09).
'  Column E = where, F = which window / case. B3 = model and time stamp.
'
'  MAPI key: INPUT!T3 (G_MAPI_KEY), else MIDAS_MAPI_KEY, else asked.
'  Tests: MIDAS_RTI_RECORD <folder> runs live and saves every answer to the folder;
'  MIDAS_RTI_TEST <folder> replays a saved run without Civil NX.
' ============================================================================

Private Const SCRIPT_ID As String = "midas-rti-read"
Private Const SCRIPT_VERSION As String = "2026-11-09a"
Private Const SCRIPT_CHANGELOG As String = "New: rail stress, deck end and rail-deck movements of the RTI model into sheet RTI_SONUC"

Private Const API_BASE As String = "https://moa-engineers.midasit.com:443/civil/"
Private Const MSG_TITLE As String = "MIDAS_RTI"
Private Const KEY_SHEET As String = "INPUT"
Private Const KEY_CELL As String = "T3"
Private Const KEY_NAME As String = "G_MAPI_KEY"
Private Const OUT_SHEET As String = "RTI_SONUC"
Private Const CHUNK_ROWS As Long = 30000
Private Const BIG As Double = 1E+300

Private mKey As String
Private http As Object
Private mRec As String, mPlay As String, mSeq As Long

' rail element ends: index k = (element index - 1) * 2 + 1 (end I) or + 2 (end J)
Private eId() As Long, eTrk() As Long, nEl As Long, elIx As Object
Private fMax() As Double, fMin() As Double, fMaxW() As String, fMinW() As String
Private tuP() As Double, tuN() As Double, lmMax() As Double, lmMin() As Double
' nodes read for displacements
Private nId() As Long, nN As Long, ndIx As Object
Private nX() As Double, nY() As Double, tX() As Double, tY() As Double
Private wNm() As String, nW As Long
Private nKind() As Long, pairIx() As Long   ' kind: t = rail node of track t, -t = track node, 0 = deck end
Private mDeGr As String, hasTUp As Boolean, hasTUn As Boolean
Private uX() As Double, uY() As Double

' ============================================================================
Public Sub MIDAS_RTI_OKU()
    Dim msg As String
    On Error GoTo Fail
    mRec = "": mPlay = "": mSeq = 0
    msg = RunAll(True)
    UiMsg msg, IIf(Left$(msg, 4) = "Done", vbInformation, vbExclamation), MSG_TITLE
    Exit Sub
Fail:
    Application.StatusBar = False
    UiMsg "STOPPED: " & Err.Description, vbExclamation, MSG_TITLE
End Sub

' live run that also saves every request and answer to folder (for the offline test)
Public Function MIDAS_RTI_RECORD(ByVal folder As String) As String
    On Error GoTo Fail
    mRec = folder: mPlay = "": mSeq = 0
    MIDAS_RTI_RECORD = RunAll(False)
    mRec = ""
    Exit Function
Fail:
    mRec = ""
    Application.StatusBar = False
    MIDAS_RTI_RECORD = "STOPPED: " & Err.Description
End Function

' offline run on a folder saved by MIDAS_RTI_RECORD
Public Function MIDAS_RTI_TEST(ByVal folder As String) As String
    On Error GoTo Fail
    mPlay = folder: mRec = "": mSeq = 0
    MIDAS_RTI_TEST = RunAll(False)
    mPlay = ""
    Exit Function
Fail:
    mPlay = ""
    Application.StatusBar = False
    MIDAS_RTI_TEST = "STOPPED: " & Err.Description
End Function

' ============================================================================
Private Function RunAll(ByVal ask As Boolean) As String
    Dim msg As String, pj As String, resp As String, ballast As Boolean, t0 As Double
    Dim ws As Worksheet, tag As String
    t0 = Timer
    If Not LayoutOk(msg) Then RunAll = msg: Exit Function
    ballast = IsBallast()
    If mPlay = "" Then
        If Not GetKey() Then RunAll = "No MAPI key: nothing was read.": Exit Function
    End If

    Note "model check"
    If Not Api("GET", "db/PJCF", "", pj) Then
        If Left$(pj, 10) = "HTTP error" Or pj = "" Then RunAll = "Civil NX does not answer (is it open, is the MAPI connected?)." & vbCrLf & Left$(pj, 200): Exit Function
        pj = ""
    End If
    If InStr(1, pj, " RTI", vbBinaryCompare) = 0 Or InStr(1, pj, "midas-bridge-build", vbBinaryCompare) = 0 Then
        If Not ask Then RunAll = "STOPPED: the open model is not a MIDAS_KOPRU RTI model (db/PJCF).": Exit Function
        If UiMsg("The model open in Civil NX is not a MIDAS_KOPRU RTI model." & vbCrLf & vbCrLf & _
                 "Read its results into " & OUT_SHEET & " anyway?", vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) <> vbYes Then
            RunAll = "Stopped: nothing was read.": Exit Function
        End If
    End If
    tag = ModelTag(pj)

    ReadGeometry
    ReadStages
    If ask Then
        If UiMsg("RTI results: " & nEl & " rail elements, " & (nN) & " nodes, " & nW & " windows." & vbCrLf & _
                 "Reading takes several minutes and keeps Civil NX busy. Continue?", vbYesNo + vbQuestion, MSG_TITLE) <> vbYes Then
            RunAll = "Stopped: nothing was written.": Exit Function
        End If
    End If
    ReadRailStress
    ReadWindowDisp
    msg = Evaluate_(ballast, tag)
    Application.StatusBar = False
    RunAll = msg & vbCrLf & "  Read time " & Format$((Timer - t0) / 60#, "0.0") & " min."
End Function

' ---- groups, nodes, rail direction ----
Private Sub ReadGeometry()
    Dim gr As String, nd As String, el As String, t As Long, i As Long, k As Long
    Dim eLs As Variant, rN As Variant, iz As Variant, de As Variant, nm As Variant
    Dim allN As Object, v As Variant, key As String
    Note "groups"
    If Not Api("GET", "db/GRUP", "", gr) Then Err.Raise vbObjectError + 10, , "No structure groups in the model. " & Left$(gr, 200)
    gr = Compact(gr)
    Set elIx = CreateObject("Scripting.Dictionary")
    Set ndIx = CreateObject("Scripting.Dictionary")
    Set allN = CreateObject("Scripting.Dictionary")
    nEl = 0: nN = 0
    ReDim eId(1 To 1): ReDim eTrk(1 To 1)
    For t = 1 To 2
        eLs = GroupList(gr, "SG_RAY_" & t, "E_LIST")
        If IsEmpty(eLs) Then
            If t = 1 Then Err.Raise vbObjectError + 10, , "Group SG_RAY_1 is missing: the model was built before MIDAS_KOPRU 2026-11-09b. Build it again."
            Exit For
        End If
        ReDim Preserve eId(1 To nEl + UBound(eLs) + 1): ReDim Preserve eTrk(1 To nEl + UBound(eLs) + 1)
        For i = 0 To UBound(eLs)
            If Not elIx.Exists(eLs(i)) Then
                nEl = nEl + 1: eId(nEl) = eLs(i): eTrk(nEl) = t: elIx(eLs(i)) = nEl
            End If
        Next i
        rN = GroupList(gr, "SG_RAY_" & t, "N_LIST")
        iz = GroupList(gr, "SG_IZ_" & t, "N_LIST")
        If IsEmpty(rN) Or IsEmpty(iz) Then Err.Raise vbObjectError + 10, , "Group SG_RAY_" & t & " or SG_IZ_" & t & " has no nodes."
        For i = 0 To UBound(rN)
            allN(rN(i)) = t
        Next i
        For i = 0 To UBound(iz)
            If Not allN.Exists(iz(i)) Then allN(iz(i)) = -t
        Next i
    Next t
    ReDim Preserve eId(1 To nEl): ReDim Preserve eTrk(1 To nEl)
    For Each nm In GroupNames(gr, "DE_")
        de = GroupList(gr, CStr(nm), "N_LIST")
        If Not IsEmpty(de) Then If Not allN.Exists(de(0)) Then allN(de(0)) = 0
    Next nm
    ReDim nId(1 To allN.Count): ReDim nKind(1 To allN.Count)
    For Each v In allN.keys
        nN = nN + 1: nId(nN) = v: nKind(nN) = allN(v): ndIx(v) = nN
    Next v
    mDeGr = gr

    Note "node coordinates"
    If Not Api("GET", "db/NODE", "", nd) Then Err.Raise vbObjectError + 11, , "Nodes could not be read. " & Left$(nd, 200)
    ReDim nX(1 To nN): ReDim nY(1 To nN): ReDim tX(1 To nN): ReDim tY(1 To nN)
    ReadNodes Compact(nd)

    Note "rail directions"
    If Not Api("GET", "db/ELEM", "", el) Then Err.Raise vbObjectError + 11, , "Elements could not be read. " & Left$(el, 200)
    RailDirections Compact(el)
    BuildPairs
End Sub

Private Sub ReadNodes(ByRef s As String)
    Dim p As Long, a As Long, id As Long, cnt As Long, kk As Long
    p = 1
    Do
        p = InStr(p, s, "{""X"":", vbBinaryCompare)
        If p = 0 Then Exit Do
        a = InStrRev(s, """", p - 3, vbBinaryCompare)
        id = CLng(Val(Mid$(s, a + 1, p - 3 - a)))
        If ndIx.Exists(id) Then
            kk = ndIx(id)
            nX(kk) = NumAfter(s, p, """X"":")
            nY(kk) = NumAfter(s, p, """Y"":")
            cnt = cnt + 1
        End If
        p = p + 5
    Loop
    If cnt <> nN Then Err.Raise vbObjectError + 11, , "Only " & cnt & " of " & nN & " group nodes were found in db/NODE."
End Sub

Private Sub RailDirections(ByRef s As String)
    Dim p As Long, a As Long, b As Long, id As Long, i1 As Long, i2 As Long, parts As Variant
    Dim k1 As Long, k2 As Long, dx As Double, dy As Double, L As Double, k As Long, nr As Long
    p = 1
    Do
        p = InStr(p, s, """NODE"":[", vbBinaryCompare)
        If p = 0 Then Exit Do
        a = InStrRev(s, """:{", p, vbBinaryCompare)
        b = InStrRev(s, """", a - 1, vbBinaryCompare)
        id = CLng(Val(Mid$(s, b + 1, a - b - 1)))
        If elIx.Exists(id) Then
            parts = Split(Mid$(s, p + 8, InStr(p + 8, s, "]", vbBinaryCompare) - p - 8), ",")
            i1 = CLng(Val(parts(0))): i2 = CLng(Val(parts(1)))
            If ndIx.Exists(i1) And ndIx.Exists(i2) Then
                k1 = ndIx(i1): k2 = ndIx(i2)
                dx = nX(k2) - nX(k1): dy = nY(k2) - nY(k1): L = Sqr(dx * dx + dy * dy)
                If L > 0 Then
                    tX(k1) = tX(k1) + dx / L: tY(k1) = tY(k1) + dy / L
                    tX(k2) = tX(k2) + dx / L: tY(k2) = tY(k2) + dy / L
                    nr = nr + 1
                End If
            End If
        End If
        p = p + 8
    Loop
    If nr <> nEl Then Err.Raise vbObjectError + 11, , "Only " & nr & " of " & nEl & " rail elements were found in db/ELEM."
    For k = 1 To nN
        L = Sqr(tX(k) * tX(k) + tY(k) * tY(k))
        If L > 0 Then tX(k) = tX(k) / L: tY(k) = tY(k) / L
    Next k
End Sub

Private Sub ReadStages()
    Dim s As String, p As Long, q As Long, nm As String
    Note "stages"
    If Not Api("GET", "db/STAG", "", s) Then Err.Raise vbObjectError + 12, , "No construction stages in the model. " & Left$(s, 200)
    s = Compact(s)
    nW = 0: ReDim wNm(1 To 1)
    hasTUp = False: hasTUn = False
    p = 1
    Do
        p = InStr(p, s, """NAME"":""", vbBinaryCompare)
        If p = 0 Then Exit Do
        q = InStr(p + 8, s, """", vbBinaryCompare)
        nm = Mid$(s, p + 8, q - p - 8)
        Select Case nm
            Case "TU+": hasTUp = True
            Case "TU-": hasTUn = True
            Case "END"
            Case Else
                If Left$(nm, 3) <> "LL_" Then
                    nW = nW + 1
                    ReDim Preserve wNm(1 To nW)
                    wNm(nW) = nm
                End If
        End Select
        p = q
    Loop
    If nW = 0 Then Err.Raise vbObjectError + 12, , "No braking / traction window stages in the model."
    If Not (hasTUp And hasTUn) Then Err.Raise vbObjectError + 12, , "Stage TU+ or TU- is missing."
End Sub

' ---- rail stress ----
Private Sub ReadRailStress()
    Dim i As Long, i0 As Long, chunk As Long, keys As String, resp As String, body As String
    Dim d As Variant, n As Long, r As Long, k As Long, v As Double, nm As String, steps As String, w As Long
    ReDim fMax(1 To 2 * nEl): ReDim fMin(1 To 2 * nEl): ReDim fMaxW(1 To 2 * nEl): ReDim fMinW(1 To 2 * nEl)
    ReDim tuP(1 To 2 * nEl): ReDim tuN(1 To 2 * nEl): ReDim lmMax(1 To 2 * nEl): ReDim lmMin(1 To 2 * nEl)
    For k = 1 To 2 * nEl: fMax(k) = -BIG: fMin(k) = BIG: Next k
    For w = 1 To nW: steps = steps & IIf(w = 1, "", ",") & """" & wNm(w) & ":001(last)""": Next w

    ' windows
    chunk = CHUNK_ROWS \ (2 * nW): If chunk < 1 Then chunk = 1
    For i0 = 1 To nEl Step chunk
        Note "rail stress, windows: elements " & i0 & "-" & Min_(i0 + chunk - 1, nEl) & " of " & nEl
        body = StressBody(i0, Min_(i0 + chunk - 1, nEl), """OPT_CS"":true,""LOAD_CASE_NAMES"":[""Summation(CS)""],""STAGE_STEP"":[" & steps & "]")
        If Not Api("POST", "post/TABLE", body, resp) Then Err.Raise vbObjectError + 13, , "Rail stress could not be read (has the model been analysed?). " & Left$(resp, 300)
        d = TableData(resp, Array("Elem", "Stage", "Part", "Axial"), n)
        For r = 1 To n
            k = EndIx(d(r, 0), d(r, 2))
            If k > 0 Then
                v = Val(d(r, 3))
                If v > fMax(k) Then fMax(k) = v: fMaxW(k) = d(r, 1)
                If v < fMin(k) Then fMin(k) = v: fMinW(k) = d(r, 1)
            End If
        Next r
    Next i0
    ' temperature
    chunk = CHUNK_ROWS \ 4
    For i0 = 1 To nEl Step chunk
        Note "rail stress, TU+ / TU-"
        body = StressBody(i0, Min_(i0 + chunk - 1, nEl), """OPT_CS"":true,""LOAD_CASE_NAMES"":[""Summation(CS)""],""STAGE_STEP"":[""TU+:001(last)"",""TU-:001(last)""]")
        If Not Api("POST", "post/TABLE", body, resp) Then Err.Raise vbObjectError + 13, , "Rail stress at TU+ / TU- could not be read. " & Left$(resp, 300)
        d = TableData(resp, Array("Elem", "Stage", "Part", "Axial"), n)
        For r = 1 To n
            k = EndIx(d(r, 0), d(r, 2))
            If k > 0 Then
                If d(r, 1) = "TU+" Then tuP(k) = Val(d(r, 3)) Else If d(r, 1) = "TU-" Then tuN(k) = Val(d(r, 3))
            End If
        Next r
    Next i0
    ' LM71 moving load
    For i0 = 1 To nEl Step chunk
        Note "rail stress, LM71"
        body = StressBody(i0, Min_(i0 + chunk - 1, nEl), """LOAD_CASE_NAMES"":[""LM71(MV:max)"",""LM71(MV:min)""]")
        If Not Api("POST", "post/TABLE", body, resp) Then Err.Raise vbObjectError + 13, , "No LM71 moving-load results (the nonlinear LM71 option is not read by this version). " & Left$(resp, 300)
        d = TableData(resp, Array("Elem", "Load", "Part", "Axial"), n)
        For r = 1 To n
            k = EndIx(d(r, 0), d(r, 2))
            If k > 0 Then
                If InStr(1, d(r, 1), "max", vbTextCompare) > 0 Then lmMax(k) = Val(d(r, 3)) Else lmMin(k) = Val(d(r, 3))
            End If
        Next r
    Next i0
End Sub

Private Function StressBody(ByVal i1 As Long, ByVal i2 As Long, ByVal tail As String) As String
    Dim i As Long, keys As String
    For i = i1 To i2: keys = keys & IIf(i = i1, "", ",") & eId(i): Next i
    StressBody = "{""Argument"":{""TABLE_NAME"":""T"",""TABLE_TYPE"":""BEAMSTRESS"",""UNIT"":{""FORCE"":""KN"",""DIST"":""M""},""STYLES"":{""FORMAT"":""Fixed"",""PLACE"":4}," & _
        """NODE_ELEMS"":{""KEYS"":[" & keys & "]},""PARTS"":[""PartI"",""PartJ""],""COMPONENTS"":[""Elem"",""Load"",""Stage"",""Part"",""Axial""]," & tail & "}}"
End Function

Private Function EndIx(ByVal elem As String, ByVal part As String) As Long
    Dim id As Long
    id = CLng(Val(elem))
    If Not elIx.Exists(id) Then Exit Function
    EndIx = (elIx(id) - 1) * 2 + IIf(Left$(part, 1) = "J", 2, 1)
End Function

' ---- displacements in the windows ----
Private Sub ReadWindowDisp()
    Dim i0 As Long, chunk As Long, i As Long, keys As String, resp As String, body As String, steps As String
    Dim d As Variant, n As Long, r As Long, k As Long, w As Long, wIx As Object
    Set wIx = CreateObject("Scripting.Dictionary")
    For w = 1 To nW: wIx(wNm(w)) = w: steps = steps & IIf(w = 1, "", ",") & """" & wNm(w) & ":001(last)""": Next w
    ReDim uX(1 To nN, 1 To nW): ReDim uY(1 To nN, 1 To nW)
    chunk = CHUNK_ROWS \ nW: If chunk < 1 Then chunk = 1
    For i0 = 1 To nN Step chunk
        Note "displacements, windows: nodes " & i0 & "-" & Min_(i0 + chunk - 1, nN) & " of " & nN
        keys = ""
        For i = i0 To Min_(i0 + chunk - 1, nN): keys = keys & IIf(i = i0, "", ",") & nId(i): Next i
        body = "{""Argument"":{""TABLE_NAME"":""T"",""TABLE_TYPE"":""DISPLACEMENTG"",""UNIT"":{""FORCE"":""KN"",""DIST"":""M""},""STYLES"":{""FORMAT"":""Scientific"",""PLACE"":6}," & _
            """NODE_ELEMS"":{""KEYS"":[" & keys & "]},""OPT_CS"":true,""LOAD_CASE_NAMES"":[""Summation(CS)""],""STAGE_STEP"":[" & steps & "]}}"
        If Not Api("POST", "post/TABLE", body, resp) Then Err.Raise vbObjectError + 14, , "Displacements could not be read. " & Left$(resp, 300)
        d = TableData(resp, Array("Node", "Stage", "DX", "DY"), n)
        For r = 1 To n
            k = ndIx(CLng(Val(d(r, 0))))
            If wIx.Exists(d(r, 1)) Then
                w = wIx(d(r, 1))
                uX(k, w) = Val(d(r, 2)): uY(k, w) = Val(d(r, 3))
            End If
        Next r
    Next i0
End Sub

' ---- the values ----
Private Function Evaluate_(ByVal ballast As Boolean, ByVal tag As String) As String
    Dim k As Long, s(1 To 4) As Double, j As Long, best As Double, worst As Double, bk As Long, bj As Long, wk As Long, wj As Long
    Dim ws As Worksheet, sTxt As Variant, w As Long, i As Long
    Dim dB As Double, dBw As String, dBe As String, dA As Double, dAw As String, dAe As String
    Dim dU As Double, dUw As String, dUe As String, v As Double, kn As Long, kt As Long
    Dim dH As Double, dHe As String, dV As Double, dVe As String
    Dim ax As Variant, kL As Long, kR As Long, ttx As Double, tty As Double
    Dim mvX() As Double, mvY() As Double, mvZ() As Double, deK As Object, lo As Double, hi As Double

    ' rail stress, four sums per element end
    best = -BIG: worst = BIG
    For k = 1 To 2 * nEl
        If fMax(k) = -BIG Then fMax(k) = 0: fMin(k) = 0
        s(1) = fMin(k) + tuP(k) + lmMin(k)
        s(2) = fMin(k) + tuN(k) + lmMin(k)
        s(3) = fMax(k) + lmMax(k) + tuP(k)
        s(4) = fMax(k) + lmMax(k) + tuN(k)
        For j = 1 To 4
            If s(j) > best Then best = s(j): bk = k: bj = j
            If s(j) < worst Then worst = s(j): wk = k: wj = j
        Next j
    Next k

    ' deck ends: label -> node index
    Set deK = DeckEnds()

    ' windows: dB, dabs at the deck ends; du at the rail nodes
    For w = 1 To nW
        For Each ax In deK.keys
            kL = deK(ax)(0): kR = deK(ax)(1)
            If kL > 0 Then ttx = tX(kL): tty = tY(kL) Else ttx = tX(kR): tty = tY(kR)
            v = 0
            If kL > 0 Then v = v + uX(kL, w) * ttx + uY(kL, w) * tty
            If kR > 0 Then v = v - (uX(kR, w) * ttx + uY(kR, w) * tty)
            If Abs(v) > dB Then dB = Abs(v): dBw = wNm(w): dBe = "Aks " & ax & IIf(kL = 0 Or kR = 0, " (kenar ayak)", "")
            If kL > 0 Then
                v = Abs(uX(kL, w) * ttx + uY(kL, w) * tty)
                If v > dA Then dA = v: dAw = wNm(w): dAe = "Aks " & ax & ", sol tabliye ucu"
            End If
            If kR > 0 Then
                v = Abs(uX(kR, w) * ttx + uY(kR, w) * tty)
                If v > dA Then dA = v: dAw = wNm(w): dAe = "Aks " & ax & ", sa" & ChrW(287) & " tabliye ucu"
            End If
        Next ax
        For kn = 1 To nN
            kt = pairIx(kn)
            If kt > 0 Then
                v = Abs((uX(kn, w) - uX(kt, w)) * tX(kn) + (uY(kn, w) - uY(kt, w)) * tY(kn))
                If v > dU Then dU = v: dUw = wNm(w): dUe = "Hat " & nKind(kn) & ", ray d" & ChrW(252) & ChrW(287) & ChrW(252) & "m" & ChrW(252) & " " & nId(kn)
            End If
        Next kn
    Next w

    ' LM71 envelopes at the deck ends: dH, dV
    ReadMvDeckEnds deK, mvX, mvY, mvZ
    For Each ax In deK.keys
        kL = deK(ax)(0): kR = deK(ax)(1)
        If kL > 0 Then ttx = tX(kL): tty = tY(kL) Else ttx = tX(kR): tty = tY(kR)
        ' along the track: hi / lo bounds of (left - right)
        hi = 0: lo = 0
        If kL > 0 Then
            hi = hi + Max_(ttx * mvX(kL, 1), ttx * mvX(kL, 2)) + Max_(tty * mvY(kL, 1), tty * mvY(kL, 2))
            lo = lo + Min2(ttx * mvX(kL, 1), ttx * mvX(kL, 2)) + Min2(tty * mvY(kL, 1), tty * mvY(kL, 2))
        End If
        If kR > 0 Then
            hi = hi - Min2(ttx * mvX(kR, 1), ttx * mvX(kR, 2)) - Min2(tty * mvY(kR, 1), tty * mvY(kR, 2))
            lo = lo - Max_(ttx * mvX(kR, 1), ttx * mvX(kR, 2)) - Max_(tty * mvY(kR, 1), tty * mvY(kR, 2))
        End If
        v = Max_(Abs(hi), Abs(lo))
        If v > dH Then dH = v: dHe = "Aks " & ax & IIf(kL = 0 Or kR = 0, " (kenar ayak)", "")
        hi = 0: lo = 0
        If kL > 0 Then hi = hi + mvZ(kL, 1): lo = lo + mvZ(kL, 2)
        If kR > 0 Then hi = hi - mvZ(kR, 2): lo = lo - mvZ(kR, 1)
        v = Max_(Abs(hi), Abs(lo))
        If v > dV Then dV = v: dVe = "Aks " & ax & IIf(kL = 0 Or kR = 0, " (kenar ayak)", "")
    Next ax

    ' write
    sTxt = Array("", "pencere min + TU+ + LM71 min", "pencere min + TU- + LM71 min", "pencere maks + LM71 maks + TU+", "pencere maks + LM71 maks + TU-")
    Set ws = ThisWorkbook.Worksheets(OUT_SHEET)
    ws.Range("D5:F11").ClearContents
    ws.Range("B3").Value = "MIDAS: " & tag & " / " & Format$(Now, "yyyy-mm-dd hh:nn") & " / " & nW & " pencere, " & nEl & " ray eleman" & ChrW(305)
    PutRow ws, 5, best / 1000#, EndText(bk), IIf(bj <= 2, fMinW(bk), fMaxW(bk)) & ": " & sTxt(bj)
    PutRow ws, 6, worst / 1000#, EndText(wk), IIf(wj <= 2, fMinW(wk), fMaxW(wk)) & ": " & sTxt(wj)
    PutRow ws, 7, dB * 1000#, dBe, dBw
    PutRow ws, 8, dH * 1000#, dHe, "LM71 zarf" & ChrW(305) & ", " & ChrW(252) & "st s" & ChrW(305) & "n" & ChrW(305) & "r"
    PutRow ws, 9, dV * 1000#, dVe, "LM71 zarf" & ChrW(305) & ", " & ChrW(252) & "st s" & ChrW(305) & "n" & ChrW(305) & "r"
    If ballast Then
        PutRow ws, 10, dU * 1000#, dUe, dUw
        PutRow ws, 11, dA * 1000#, dAe, dAw
    End If
    Application.Calculate

    Evaluate_ = "Done: RTI results written to " & OUT_SHEET & " (" & nW & " windows, " & nEl & " rail elements, " & deK.Count & " axes)." & vbCrLf & _
        "  Rail stress  max " & Format$(best / 1000#, "0.00") & " / min " & Format$(worst / 1000#, "0.00") & " N/mm2" & vbCrLf & _
        "  dB " & Format$(dB * 1000#, "0.000") & " mm, dH " & Format$(dH * 1000#, "0.000") & " mm, dV " & Format$(dV * 1000#, "0.000") & " mm" & vbCrLf & _
        IIf(ballast, "  du " & Format$(dU * 1000#, "0.000") & " mm, dabs " & Format$(dA * 1000#, "0.000") & " mm", "  Slab track: du and dabs are not checked (left empty).")
End Function

Private Sub PutRow(ws As Worksheet, ByVal r As Long, ByVal v As Double, ByVal where As String, ByVal what As String)
    ws.Cells(r, 4).Value = v
    ws.Cells(r, 5).Value = where
    ws.Cells(r, 6).Value = what
End Sub

Private Function EndText(ByVal k As Long) As String
    Dim e As Long
    If k = 0 Then Exit Function
    e = (k + 1) \ 2
    EndText = "Hat " & eTrk(e) & ", eleman " & eId(e) & IIf(k Mod 2 = 1, " (I ucu)", " (J ucu)")
End Function

' axis label -> Array(left node index, right node index), 0 when that side has no deck
Private Function DeckEnds() As Object
    Dim d As Object, nm As Variant, ax As String, de As Variant, a As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For Each nm In GroupNames(mDeGr, "DE_")
        de = GroupList(mDeGr, CStr(nm), "N_LIST")
        If Not IsEmpty(de) Then
            ax = Mid$(nm, 4, Len(nm) - 5)
            If Not d.Exists(ax) Then d(ax) = Array(0, 0)
            a = d(ax)
            If Right$(nm, 2) = "_L" Then a(0) = ndIx(de(0)) Else a(1) = ndIx(de(0))
            d(ax) = a
        End If
    Next nm
    ' deck end direction: the nearest rail node of track 1
    DeckDirections d
    Set DeckEnds = d
End Function

Private Sub DeckDirections(d As Object)
    Dim ax As Variant, j As Long, k As Long, kb As Long, i As Long, dd As Double, bd As Double
    For Each ax In d.keys
        For j = 0 To 1
            k = d(ax)(j)
            If k > 0 Then
                bd = BIG: kb = 0
                For i = 1 To nN
                    If nKind(i) = 1 Then
                        dd = (nX(i) - nX(k)) ^ 2 + (nY(i) - nY(k)) ^ 2
                        If dd < bd Then bd = dd: kb = i
                    End If
                Next i
                If kb > 0 Then tX(k) = tX(kb): tY(k) = tY(kb)
            End If
        Next j
    Next ax
End Sub

' rail node index -> track node index of the same track at the same plan point (0 when none)
Private Sub BuildPairs()
    Dim xy As Object, i As Long, key As String, n As Long
    Set xy = CreateObject("Scripting.Dictionary")
    ReDim pairIx(1 To nN)
    For i = 1 To nN
        If nKind(i) < 0 Then xy(CStr(-nKind(i)) & "|" & PlanKey(i)) = i
    Next i
    For i = 1 To nN
        If nKind(i) > 0 Then
            key = CStr(nKind(i)) & "|" & PlanKey(i)
            If xy.Exists(key) Then pairIx(i) = xy(key): n = n + 1
        End If
    Next i
    If n = 0 Then Err.Raise vbObjectError + 11, , "No rail node has a track node at the same plan point."
End Sub

Private Function PlanKey(ByVal i As Long) As String
    PlanKey = Format$(Round(nX(i) * 1000#, 0), "0") & "_" & Format$(Round(nY(i) * 1000#, 0), "0")
End Function

Private Sub ReadMvDeckEnds(deK As Object, mvX() As Double, mvY() As Double, mvZ() As Double)
    Dim ax As Variant, j As Long, k As Long, keys As String, resp As String, body As String
    Dim d As Variant, n As Long, r As Long, c As Long
    ReDim mvX(1 To nN, 1 To 2): ReDim mvY(1 To nN, 1 To 2): ReDim mvZ(1 To nN, 1 To 2)
    For Each ax In deK.keys
        For j = 0 To 1
            k = deK(ax)(j)
            If k > 0 Then keys = keys & IIf(keys = "", "", ",") & nId(k)
        Next j
    Next ax
    Note "displacements, LM71"
    body = "{""Argument"":{""TABLE_NAME"":""T"",""TABLE_TYPE"":""DISPLACEMENTG"",""UNIT"":{""FORCE"":""KN"",""DIST"":""M""},""STYLES"":{""FORMAT"":""Scientific"",""PLACE"":6}," & _
        """NODE_ELEMS"":{""KEYS"":[" & keys & "]},""LOAD_CASE_NAMES"":[""LM71(MV:max)"",""LM71(MV:min)""]}}"
    If Not Api("POST", "post/TABLE", body, resp) Then Err.Raise vbObjectError + 15, , "LM71 displacements could not be read. " & Left$(resp, 300)
    d = TableData(resp, Array("Node", "Load", "DX", "DY", "DZ"), n)
    For r = 1 To n
        k = ndIx(CLng(Val(d(r, 0))))
        c = IIf(InStr(1, d(r, 1), "max", vbTextCompare) > 0, 1, 2)
        mvX(k, c) = Val(d(r, 2)): mvY(k, c) = Val(d(r, 3)): mvZ(k, c) = Val(d(r, 4))
    Next r
End Sub

' ---- table and JSON helpers ----
' DATA of a post/TABLE answer: String(1..rows, 0..UBound(want)); every DATA value is a quoted string
Private Function TableData(ByRef resp As String, ByVal want As Variant, ByRef nRows As Long) As Variant
    Dim pH As Long, pE As Long, pD As Long, head As Variant, nC As Long, colTo() As Long, j As Long, c As Long
    Dim p As Long, q1 As Long, q2 As Long, r As Long, out() As String, found As Boolean
    nRows = 0
    pH = InStr(1, resp, """HEAD"":[", vbBinaryCompare)
    If pH = 0 Then Err.Raise vbObjectError + 16, , "The answer has no table. " & Left$(resp, 300)
    pE = InStr(pH, resp, "]", vbBinaryCompare)
    head = Split(Replace(Replace(Mid$(resp, pH + 8, pE - pH - 8), """", ""), " ", ""), ",")
    nC = UBound(head) + 1
    ReDim colTo(0 To nC - 1)
    For c = 0 To nC - 1: colTo(c) = -1: Next c
    For j = 0 To UBound(want)
        found = False
        For c = 0 To nC - 1
            If StrComp(head(c), want(j), vbTextCompare) = 0 Then colTo(c) = j: found = True: Exit For
        Next c
        If Not found Then Err.Raise vbObjectError + 16, , "The table has no column " & want(j) & " (" & Join(head, ", ") & ")."
    Next j
    pD = InStr(pE, resp, """DATA"":[", vbBinaryCompare)
    If pD = 0 Then TableData = Empty: Exit Function
    p = pD + 8
    If Mid$(resp, p, 1) = "]" Then TableData = Empty: Exit Function
    nRows = 1
    q1 = p
    Do
        q1 = InStr(q1, resp, "],[", vbBinaryCompare)
        If q1 = 0 Then Exit Do
        nRows = nRows + 1: q1 = q1 + 3
    Loop
    ReDim out(1 To nRows, 0 To UBound(want))
    For r = 1 To nRows
        For c = 0 To nC - 1
            q1 = InStr(p, resp, """", vbBinaryCompare)
            q2 = InStr(q1 + 1, resp, """", vbBinaryCompare)
            If q1 = 0 Or q2 = 0 Then Err.Raise vbObjectError + 16, , "Table row " & r & " is cut short."
            If colTo(c) >= 0 Then out(r, colTo(c)) = Mid$(resp, q1 + 1, q2 - q1 - 1)
            p = q2 + 1
        Next c
    Next r
    TableData = out
End Function

' numbers of key in the record of group name; Empty when missing
Private Function GroupList(ByRef s As String, ByVal nm As String, ByVal key As String) As Variant
    Dim p As Long, pn As Long, pk As Long, pe As Long, parts As Variant, a() As Long, i As Long
    p = InStr(1, s, """NAME"":""" & nm & """", vbBinaryCompare)
    If p = 0 Then Exit Function
    pn = InStr(p + 1, s, """NAME"":""", vbBinaryCompare)
    If pn = 0 Then pn = Len(s) + 1
    pk = InStr(p, s, """" & key & """:[", vbBinaryCompare)
    If pk = 0 Or pk > pn Then Exit Function
    pe = InStr(pk, s, "]", vbBinaryCompare)
    If pe = pk + Len(key) + 4 Then Exit Function
    parts = Split(Mid$(s, pk + Len(key) + 4, pe - pk - Len(key) - 4), ",")
    ReDim a(0 To UBound(parts))
    For i = 0 To UBound(parts): a(i) = CLng(Val(parts(i))): Next i
    GroupList = a
End Function

Private Function GroupNames(ByRef s As String, ByVal prefix As String) As Collection
    Dim c As New Collection, p As Long, q As Long
    p = 1
    Do
        p = InStr(p, s, """NAME"":""" & prefix, vbBinaryCompare)
        If p = 0 Then Exit Do
        q = InStr(p + 8, s, """", vbBinaryCompare)
        c.Add Mid$(s, p + 8, q - p - 8)
        p = q
    Loop
    Set GroupNames = c
End Function

Private Function NumAfter(ByRef s As String, ByVal p As Long, ByVal key As String) As Double
    Dim a As Long, b As Long
    a = InStr(p, s, key, vbBinaryCompare)
    If a = 0 Then Exit Function
    a = a + Len(key)
    b = a
    Do While InStr(1, "+-0123456789.eE", Mid$(s, b, 1), vbBinaryCompare) > 0: b = b + 1: Loop
    NumAfter = Val(Mid$(s, a, b - a))
End Function

' JSON without line breaks and blanks around the separators
Private Function Compact(ByVal s As String) As String
    s = Replace(Replace(Replace(s, vbCr, ""), vbLf, ""), vbTab, "")
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0: s = Replace(s, "  ", " "): Loop
    s = Replace(Replace(Replace(Replace(s, """: ", """:"), ", """, ","""), "{ ", "{"), " }", "}")
    s = Replace(Replace(Replace(Replace(s, "[ ", "["), " ]", "]"), ", ", ","), " :", ":")
    Compact = s
End Function

' ---- checks ----
Private Function LayoutOk(ByRef msg As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(OUT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then msg = "STOPPED: no sheet " & OUT_SHEET & " in this workbook.": Exit Function
    If InStr(1, CStr(ws.Range("B5").Value), "Ray gerilmesi", vbTextCompare) <> 1 Or CStr(ws.Range("B11").Value) <> ChrW(&H3B4) & "abs" Then
        msg = "STOPPED: sheet " & OUT_SHEET & " has another layout (B5 'Ray gerilmesi...', B11 '" & ChrW(&H3B4) & "abs' expected). Nothing was written."
        Exit Function
    End If
    LayoutOk = True
End Function

Private Function IsBallast() As Boolean
    Dim v As String
    On Error Resume Next
    v = CStr(ThisWorkbook.Names("G_HAT_TIPI").RefersToRange.Cells(1, 1).Value)
    On Error GoTo 0
    IsBallast = (InStr(1, v, "Balast", vbTextCompare) = 1)
End Function

Private Function ModelTag(ByVal pj As String) As String
    Dim p As Long, q As Long
    p = InStr(1, pj, """COMMENT"":", vbBinaryCompare)
    If p > 0 Then
        p = InStr(p + 10, pj, """", vbBinaryCompare)
        q = InStr(p + 1, pj, """", vbBinaryCompare)
        If p > 0 And q > p Then ModelTag = Mid$(pj, p + 1, q - p - 1)
    End If
    If ModelTag = "" Then ModelTag = "model without project data"
End Function

Private Sub Note(ByVal txt As String)
    Application.StatusBar = "MIDAS_RTI: " & txt
    If mPlay = "" And mRec = "" Then UiLog "Civil NX: " & txt
    DoEvents
End Sub

Private Function Min_(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then Min_ = a Else Min_ = b
End Function

Private Function Max_(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then Max_ = a Else Max_ = b
End Function

Private Function Min2(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then Min2 = a Else Min2 = b
End Function

' ---- API (live, record or replay) ----
Private Function Api(ByVal method As String, ByVal ep As String, ByVal body As String, ByRef resp As String) As Boolean
    Dim f As String, req As String
    mSeq = mSeq + 1
    req = method & " " & ep & vbLf & body
    If mPlay <> "" Then
        f = mPlay & "\" & Format$(mSeq, "000")
        If Dir$(f & ".req") = "" Then Err.Raise vbObjectError + 20, , "Test folder has no answer " & mSeq & " (" & method & " " & ep & ")."
        If ReadUtf8(f & ".req") <> req Then Err.Raise vbObjectError + 20, , "Test answer " & mSeq & " was saved for another request (" & method & " " & ep & ")."
        resp = ReadUtf8(f & ".json")
        Api = (InStr(1, resp, """error""", vbTextCompare) = 0)
        Exit Function
    End If
    On Error GoTo Bad
    If http Is Nothing Then Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open method, API_BASE & ep, False
    http.SetTimeouts 30000, 60000, 900000, 900000
    http.SetRequestHeader "MAPI-Key", mKey
    http.SetRequestHeader "Content-Type", "application/json"
    If method = "GET" Then http.Send Else http.Send body
    resp = http.ResponseText
    If http.Status = 404 Then mKey = ""
    Api = (http.Status >= 200 And http.Status < 300 And InStr(1, resp, """error""", vbTextCompare) = 0)
    If mRec <> "" Then
        f = mRec & "\" & Format$(mSeq, "000")
        WriteUtf8 f & ".req", req
        WriteUtf8 f & ".json", resp
    End If
    Exit Function
Bad:
    resp = "HTTP error: " & Err.Description
    Set http = Nothing
    Api = False
End Function

Private Function GetKey() As Boolean
    Dim cellKey As String
    On Error Resume Next
    cellKey = Trim$(CStr(ThisWorkbook.Names(KEY_NAME).RefersToRange.Cells(1, 1).Value))
    If cellKey = "" Then cellKey = Trim$(CStr(ThisWorkbook.Worksheets(KEY_SHEET).Range(KEY_CELL).Value))
    On Error GoTo 0
    If cellKey <> "" Then mKey = cellKey
    If mKey = "" Then mKey = Environ$("MIDAS_MAPI_KEY")
    If mKey = "" And mRec = "" Then mKey = UiInputS("MIDAS MAPI key (kept in memory for this Excel session only):", MSG_TITLE)
    GetKey = (mKey <> "")
End Function

Private Function ReadUtf8(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText
    st.Close
End Function

Private Sub WriteUtf8(ByVal path As String, ByVal txt As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText txt
    st.SaveToFile path, 2
    st.Close
End Sub
