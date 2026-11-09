Attribute VB_Name = "MIDAS_RTI"
Option Explicit
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal h As LongPtr, ByVal c As Long) As LongPtr
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowTextW Lib "user32" (ByVal h As LongPtr, ByVal s As LongPtr, ByVal n As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal h As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal h As LongPtr, ByVal n As Long) As Long
' ============================================================================
'  MIDAS_RTI - results of the RTI model (MIDAS_KOPRU, rail-track interaction) into
'  sheet RTI_SONUC of the main workbook; 3B_RTI K166, K167, J207:J211 read it.
'
'  Macro window (MAKROLAR) > MIDAS sonuclari > RTI sonuclari, or MIDAS_RTI_OKU.
'  Reads the model OPEN in Civil NX through the MAPI, read only: GET db/PJCF, GRUP,
'  NODE, ELEM, STAG and post/TABLE. The results are not changed; for the pictures it changes the
'  view and, while the result pictures are taken, the unit system (N, mm; set back after).
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
'  Window field 'islem': read + pictures (default), pictures only (no read: the windows and places
'  come from RTI_SONUC E / F, whose B3 stamp must name the open model), read only. Pictures: 10
'  into 3B_RTI_Y, see PICTURES below; field 'res' = picture width (4000 / 3000 / 2000 px or screen).
'  MAPI key: INPUT!T3 (G_MAPI_KEY), else MIDAS_MAPI_KEY, else asked.
'  Tests: MIDAS_RTI_RECORD <folder> runs live and saves every answer to the folder;
'  MIDAS_RTI_TEST <folder> replays a saved run without Civil NX; with pictures True it only lists
'  the picture requests (picture_plan.txt in the folder).
' ============================================================================

Private Const SCRIPT_ID As String = "midas-rti-read"
Private Const SCRIPT_VERSION As String = "2026-11-09e"
Private Const SCRIPT_CHANGELOG As String = "Pictures: LM71 strip over the whole bridge with Civil NX zoom fit and a diagram scale field (olcek); LM71 values MinMax Only, Min & Max, Limit Scale 0.01 % (as the GUI); stage strips become labelled places for pictures pasted by hand (Civil NX ignores STAGE_NAME in a capture), no displacement page, LM71 diagram at half scale, later runs keep pasted pictures; window field Islem: read + pictures, pictures only (windows from RTI_SONUC), read only; pictures into 3B_RTI_Y in the same run: model, rail stress strips (N/mm2) and deck displacement strips (mm) at the governing windows; a read without a valid answer is sent again, up to 3 times; reads shown step by step in the macro window log; new: rail stress, deck end and rail-deck movements of the RTI model into sheet RTI_SONUC"

Private Const API_BASE As String = "https://moa-engineers.midasit.com:443/civil/"
Private Const MSG_TITLE As String = "MIDAS_RTI"
Private Const KEY_SHEET As String = "INPUT"
Private Const KEY_CELL As String = "T3"
Private Const KEY_NAME As String = "G_MAPI_KEY"
Private Const OUT_SHEET As String = "RTI_SONUC"
Private Const CHUNK_ROWS As Long = 30000
Private Const BIG As Double = 1E+300
' pictures (3B_RTI_Y)
Private Const FIG_SHEET As String = "3B_RTI_Y"
Private Const PAGE_ROWS As String = "29,61,93"      ' caption rows, pages 1-3 (owner removed the displacement page)
Private Const CAM_H As Double = 30                  ' model pictures: camera as MIDAS_SEKIL
Private Const CAM_V As Double = 30
Private Const CAM_V_STRIP As Double = 0             ' result strips: side view
Private Const REG_SPANS As Double = 2               ' close-ups: +- this many spans (owner: a)
Private Const CS_CASE As String = "Summation"
Private Const MV_CASE As String = "LM71"
Private Const MV_MINMAX As String = "All"
Private Const MV_STAGE As String = "PostCS"
Private Const N_COLOR As Long = 12
Private Const WAIT_S As Long = 30
Private Const FILL As Double = 0.9
Private Const K_FIT As Double = 0.627               ' zoom fit height / 3D diagonal (MIDAS_SEKIL, measured)
Private Const ZOOM_RATE As Double = 0.0203
Private Const RES_W As String = "4000,3000,2000,0"  ' window field res; 0 = screen size
Private Const PIC_PREFIX As String = "RTI_RESIM_"
Private Const LBL_PREFIX As String = "RTI_ETIKET_"
Private Const LIMIT_PCT As Double = 0.01             ' values: MinMax Only, Min & Max, Limit Scale 0.01 % (owner's GUI)
Private Const SLOT_PREFIX As String = "RTI_YER_"     ' place for a stage picture pasted by hand
Private Const FRAME_TAG As String = "RTI_FRAME"      ' frame of a page, kept in the alternative text

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
' pictures: all nodes and elements, the governing places, camera state
Private gnX() As Double, gnY() As Double, gnZ() As Double, gnN As Long, gnIx As Object
Private aeId() As Long, aeA() As Long, aeB() As Long, aeN As Long, aeIx As Object
Private rgIdx() As Long, rgN As Long, mDeK As Object
Private mAxDeg As Double, mX0 As Double, mY0 As Double, mSpan As Double
Private gTenE As Long, gTenWin As String, gCmpE As Long, gCmpWin As String
Private gDbAx As String, gDbWin As String, gDhAx As String, gDaAx As String, gDaWin As String
Private mPics As Boolean, mOnly As Boolean, mResW As Long, mLastMode As String, mAspect As Double, mPlanLog As String
Private mWm As Double, mHm As Double, mD As Double
Private mNoLimit As Boolean   ' Civil NX refused LIMIT_SCALE as a decimal: sent without it
Private mDiag As Double       ' diagram scale of the rail stress strip (window field olcek, default 1)

' ============================================================================
Public Sub MIDAS_RTI_OKU()
    Dim msg As String
    On Error GoTo Fail
    mRec = "": mPlay = "": mSeq = 0
    ' field islem: 1 read + pictures, 2 pictures only (windows from RTI_SONUC), 3 read only
    mPics = (UiIdx("islem", 1) <> 3): mOnly = (UiIdx("islem", 1) = 2)
    mResW = ResWidth(UiIdx("res", 1))
    mDiag = UiNum("olcek", 1)
    If mDiag <= 0 Then mDiag = 1
    msg = RunAll(True)
    UiMsg msg, IIf(Left$(msg, 4) = "Done", vbInformation, vbExclamation), MSG_TITLE
    Exit Sub
Fail:
    Application.StatusBar = False
    UiMsg "STOPPED: " & Err.Description, vbExclamation, MSG_TITLE
End Sub

' live run that also saves every request and answer to folder (for the offline test)
Public Function MIDAS_RTI_RECORD(ByVal folder As String, Optional ByVal withPics As Boolean = False, Optional ByVal resW As Long = 4000, Optional ByVal onlyPics As Boolean = False) As String
    On Error GoTo Fail
    mRec = folder: mPlay = "": mSeq = 0: mPics = withPics Or onlyPics: mOnly = onlyPics: mResW = resW: mDiag = 1
    MIDAS_RTI_RECORD = RunAll(False)
    mRec = ""
    Exit Function
Fail:
    mRec = ""
    Application.StatusBar = False
    MIDAS_RTI_RECORD = "STOPPED: " & Err.Description
End Function

' offline run on a folder saved by MIDAS_RTI_RECORD (pictures: only the requests are listed)
Public Function MIDAS_RTI_TEST(ByVal folder As String, Optional ByVal withPics As Boolean = False, Optional ByVal resW As Long = 4000, Optional ByVal onlyPics As Boolean = False) As String
    On Error GoTo Fail
    mPlay = folder: mRec = "": mSeq = 0: mPics = withPics Or onlyPics: mOnly = onlyPics: mResW = resW: mDiag = 1
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

    Note 1, "read model data (db/PJCF)"
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
    If mOnly Then RunAll = PicturesOnly(ask, pj): Application.StatusBar = False: Exit Function
    ReadStages
    If ask Then
        If UiMsg("RTI results: " & nEl & " rail elements, " & (nN) & " nodes, " & nW & " windows." & vbCrLf & _
                 "Reading takes several minutes and keeps Civil NX busy." & IIf(mPics, vbCrLf & "Then 3 pictures go into " & FIG_SHEET & " (and 4 labelled places for the stage pictures pasted by hand): Civil NX shows each view and uses N, mm for the LM71 picture (set back after).", "") & vbCrLf & "Continue?", vbYesNo + vbQuestion, MSG_TITLE) <> vbYes Then
            RunAll = "Stopped: nothing was written.": Exit Function
        End If
    End If
    ReadRailStress
    ReadWindowDisp
    msg = Evaluate_(ballast, tag)
    If mPics Then msg = msg & vbCrLf & Pictures()
    Application.StatusBar = False
    RunAll = msg & vbCrLf & "  Read time " & Format$((Timer - t0) / 60#, "0.0") & " min."
End Function

' ---- groups, nodes, rail direction ----
Private Sub ReadGeometry()
    Dim gr As String, nd As String, el As String, t As Long, i As Long, k As Long
    Dim eLs As Variant, rN As Variant, iz As Variant, de As Variant, nm As Variant
    Dim allN As Object, v As Variant, key As String
    Note 3, "read groups (db/GRUP)"
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

    Note 5, "read nodes (db/NODE)"
    If Not Api("GET", "db/NODE", "", nd) Then Err.Raise vbObjectError + 11, , "Nodes could not be read. " & Left$(nd, 200)
    ReDim nX(1 To nN): ReDim nY(1 To nN): ReDim tX(1 To nN): ReDim tY(1 To nN)
    ReadNodes Compact(nd)

    Note 7, "read elements (db/ELEM)"
    If Not Api("GET", "db/ELEM", "", el) Then Err.Raise vbObjectError + 11, , "Elements could not be read. " & Left$(el, 200)
    RailDirections Compact(el)
    BuildPairs
End Sub

' all nodes (kept for the pictures); the group nodes also into nX / nY
Private Sub ReadNodes(ByRef s As String)
    Dim p As Long, a As Long, id As Long, cnt As Long, kk As Long, nAll As Long
    p = 1
    Do
        p = InStr(p, s, "{""X"":", vbBinaryCompare)
        If p = 0 Then Exit Do
        nAll = nAll + 1: p = p + 5
    Loop
    gnN = 0
    Set gnIx = CreateObject("Scripting.Dictionary")
    ReDim gnX(1 To IIf(nAll > 0, nAll, 1)): ReDim gnY(1 To IIf(nAll > 0, nAll, 1)): ReDim gnZ(1 To IIf(nAll > 0, nAll, 1))
    p = 1
    Do
        p = InStr(p, s, "{""X"":", vbBinaryCompare)
        If p = 0 Then Exit Do
        a = InStrRev(s, """", p - 3, vbBinaryCompare)
        id = CLng(Val(Mid$(s, a + 1, p - 3 - a)))
        gnN = gnN + 1
        gnX(gnN) = NumAfter(s, p, """X"":"): gnY(gnN) = NumAfter(s, p, """Y"":"): gnZ(gnN) = NumAfter(s, p, """Z"":")
        gnIx(id) = gnN
        If ndIx.Exists(id) Then
            kk = ndIx(id)
            nX(kk) = gnX(gnN): nY(kk) = gnY(gnN)
            cnt = cnt + 1
        End If
        p = p + 5
    Loop
    If cnt <> nN Then Err.Raise vbObjectError + 11, , "Only " & cnt & " of " & nN & " group nodes were found in db/NODE."
End Sub

' all elements (kept for the pictures); rail direction at every rail node
Private Sub RailDirections(ByRef s As String)
    Dim p As Long, a As Long, b As Long, id As Long, i1 As Long, i2 As Long, parts As Variant, nAll As Long
    Dim k1 As Long, k2 As Long, dx As Double, dy As Double, L As Double, k As Long, nr As Long
    p = 1
    Do
        p = InStr(p, s, """NODE"":[", vbBinaryCompare)
        If p = 0 Then Exit Do
        nAll = nAll + 1: p = p + 8
    Loop
    aeN = 0
    Set aeIx = CreateObject("Scripting.Dictionary")
    ReDim aeId(1 To IIf(nAll > 0, nAll, 1)): ReDim aeA(1 To IIf(nAll > 0, nAll, 1)): ReDim aeB(1 To IIf(nAll > 0, nAll, 1))
    p = 1
    Do
        p = InStr(p, s, """NODE"":[", vbBinaryCompare)
        If p = 0 Then Exit Do
        a = InStrRev(s, """:{", p, vbBinaryCompare)
        b = InStrRev(s, """", a - 1, vbBinaryCompare)
        id = CLng(Val(Mid$(s, b + 1, a - b - 1)))
        parts = Split(Mid$(s, p + 8, InStr(p + 8, s, "]", vbBinaryCompare) - p - 8), ",")
        i1 = CLng(Val(parts(0))): i2 = CLng(Val(parts(1)))
        aeN = aeN + 1: aeId(aeN) = id: aeA(aeN) = i1: aeB(aeN) = i2: aeIx(id) = aeN
        If elIx.Exists(id) Then
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
    Note 9, "read stages (db/STAG)"
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
        Note 10 + 50 * (i0 - 1) / nEl, "read rail stress, windows: elements " & i0 & "-" & Min_(i0 + chunk - 1, nEl) & " / " & nEl
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
        Note 62, "read rail stress TU+ / TU-: elements " & i0 & "-" & Min_(i0 + chunk - 1, nEl)
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
        Note 66, "read rail stress LM71: elements " & i0 & "-" & Min_(i0 + chunk - 1, nEl)
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
        Note 70 + 25 * (i0 - 1) / nN, "read displacements, windows: nodes " & i0 & "-" & Min_(i0 + chunk - 1, nN) & " / " & nN
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
    gTenE = 0: gCmpE = 0: gTenWin = "": gCmpWin = "": gDbAx = "": gDhAx = "": gDaAx = ""
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
    Set mDeK = deK
    If bk > 0 Then gTenE = (bk + 1) \ 2: gTenWin = IIf(bj <= 2, fMinW(bk), fMaxW(bk))
    If wk > 0 Then gCmpE = (wk + 1) \ 2: gCmpWin = IIf(wj <= 2, fMinW(wk), fMaxW(wk))

    ' windows: dB, dabs at the deck ends; du at the rail nodes
    For w = 1 To nW
        For Each ax In deK.keys
            kL = deK(ax)(0): kR = deK(ax)(1)
            If kL > 0 Then ttx = tX(kL): tty = tY(kL) Else ttx = tX(kR): tty = tY(kR)
            v = 0
            If kL > 0 Then v = v + uX(kL, w) * ttx + uY(kL, w) * tty
            If kR > 0 Then v = v - (uX(kR, w) * ttx + uY(kR, w) * tty)
            If Abs(v) > dB Then dB = Abs(v): dBw = wNm(w): dBe = "Aks " & ax & IIf(kL = 0 Or kR = 0, " (kenar ayak)", ""): gDbAx = CStr(ax)
            If kL > 0 Then
                v = Abs(uX(kL, w) * ttx + uY(kL, w) * tty)
                If v > dA Then dA = v: dAw = wNm(w): dAe = "Aks " & ax & ", sol tabliye ucu": gDaAx = CStr(ax)
            End If
            If kR > 0 Then
                v = Abs(uX(kR, w) * ttx + uY(kR, w) * tty)
                If v > dA Then dA = v: dAw = wNm(w): dAe = "Aks " & ax & ", sa" & ChrW(287) & " tabliye ucu": gDaAx = CStr(ax)
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
        If v > dH Then dH = v: dHe = "Aks " & ax & IIf(kL = 0 Or kR = 0, " (kenar ayak)", ""): gDhAx = CStr(ax)
        hi = 0: lo = 0
        If kL > 0 Then hi = hi + mvZ(kL, 1): lo = lo + mvZ(kL, 2)
        If kR > 0 Then hi = hi - mvZ(kR, 2): lo = lo - mvZ(kR, 1)
        v = Max_(Abs(hi), Abs(lo))
        If v > dV Then dV = v: dVe = "Aks " & ax & IIf(kL = 0 Or kR = 0, " (kenar ayak)", "")
    Next ax

    gDbWin = dBw: gDaWin = dAw
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
    Note 97, "read deck end displacements LM71"
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

' progress: status bar, and the macro window bar and log when the window runs it
Private Sub Note(ByVal pct As Double, ByVal txt As String)
    Application.StatusBar = "MIDAS_RTI: " & txt
    If mPlay = "" And mRec = "" Then UiProgress pct, txt
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
    Dim f As String, req As String, tries As Long, side As Boolean
    ' picture calls (view/, db/UNIT) are not numbered or recorded; a replay only lists them
    side = (Left$(ep, 5) = "view/" Or ep = "db/UNIT")
    req = method & " " & ep & vbLf & body
    If side And mPlay <> "" Then
        mPlanLog = mPlanLog & vbLf & "  " & method & " " & ep & " " & body
        If method = "GET" Then resp = "{""UNIT"":{""1"":{""FORCE"":""KN"",""DIST"":""M"",""HEAT"":""KJ"",""TEMPER"":""C""}}}" Else resp = "{""message"":""dry run""}"
        Api = True
        Exit Function
    End If
    If Not side Then mSeq = mSeq + 1
    If mPlay <> "" Then
        f = mPlay & "\" & Format$(mSeq, "000")
        If Dir$(f & ".req") = "" Then Err.Raise vbObjectError + 20, , "Test folder has no answer " & mSeq & " (" & method & " " & ep & ")."
        If ReadUtf8(f & ".req") <> req Then Err.Raise vbObjectError + 20, , "Test answer " & mSeq & " was saved for another request (" & method & " " & ep & ")."
        resp = ReadUtf8(f & ".json")
        Api = (InStr(1, resp, """error""", vbTextCompare) = 0)
        Exit Function
    End If
    ' the MAPI goes through the MIDAS relay server: a large answer can arrive broken
    ' ('invalid or unrecognized response', owner 2026-11-09), so a read without a valid
    ' answer is sent again, up to 3 times
Again:
    On Error GoTo Bad
    If http Is Nothing Then Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open method, API_BASE & ep, False
    http.SetTimeouts 30000, 60000, 900000, 900000
    http.SetRequestHeader "MAPI-Key", mKey
    http.SetRequestHeader "Content-Type", "application/json"
    If method = "GET" Then http.Send Else http.Send body
    resp = http.ResponseText
    If http.Status >= 500 And tries < 2 Then
        tries = tries + 1
        RetryNote ep, "HTTP " & http.Status, tries
        GoTo Again
    End If
    If http.Status = 404 Then mKey = ""
    Api = (http.Status >= 200 And http.Status < 300 And InStr(1, resp, """error""", vbTextCompare) = 0)
    If mRec <> "" And Not side Then
        f = mRec & "\" & Format$(mSeq, "000")
        WriteUtf8 f & ".req", req
        WriteUtf8 f & ".json", resp
    End If
    Exit Function
Bad:
    resp = "HTTP error: " & Err.Description
    Set http = Nothing
    If tries < 2 Then
        tries = tries + 1
        RetryNote ep, Err.Description, tries
        Resume Again
    End If
    Api = False
End Function

Private Sub RetryNote(ByVal ep As String, ByVal why As String, ByVal tries As Long)
    If mPlay = "" And mRec = "" Then UiLog "! " & ep & ": no valid answer (" & why & "), sent again (" & (tries + 1) & " / 3)"
    Application.Wait Now + TimeSerial(0, 0, 3 * tries)
End Sub

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

' ============================================================================
'  PICTURES in 3B_RTI_Y (owner 2026-11-09: one macro with the result pull)
'  page 1  model, line view with supports, whole bridge
'  page 2  model, solid, +-REG_SPANS spans around the axis of the largest dB
'  page 3  5 strips of rail axial stress Sax (N/mm2): LM71 (captured: the whole rail of the track
'          of the largest tension, Civil NX zoom fit, diagram scale from field olcek), then the
'          window of the largest tension, the window of the
'          largest compression, TU+, TU-: Civil NX ignores STAGE_NAME in a capture (owner's run
'          2026-11-09), so these four get a labelled place for a picture pasted by hand
'  (the displacement page was removed by the owner 2026-11-09)
'  A strip shows +-REG_SPANS spans (ACTIVE elements), side view; a label names its case.
'  The strips fill the frame of the page (kept in the alternative text of the macro's shapes); on
'  the first run the old pictures go only after all the new ones are in; later runs replace only
'  the macro's own shapes, so pictures pasted by hand stay. Units N, mm for the result pictures,
'  set back after (as the Culvert macros).
' ============================================================================
Private Function Pictures() As String
    Dim ws As Worksheet, rows As Variant, p As Long, res As String, rep As String, nOk As Long
    Dim uOld As String, uSet As Boolean, folder As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FIG_SHEET)
    On Error GoTo Fail
    If ws Is Nothing Then Pictures = "  Pictures: no sheet " & FIG_SHEET & " in this workbook, none taken.": Exit Function
    folder = Environ$("TEMP") & "\Bridge_RTI_Y"
    If mPlay = "" Then
        If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
    End If
    mLastMode = "": mAspect = 0: mPlanLog = "": mNoLimit = False
    PicGeometry
    rows = Split(PAGE_ROWS, ",")
    For p = 1 To 3
        If p = 3 Then
            uSet = UnitsTo(uOld)
            If Not uSet Then rep = rep & vbCrLf & "  ! the unit system could not be set to N, mm: the result pictures show kN, m"
        End If
        res = PicturePage(ws, p, CLng(rows(p - 1)), folder)
        If res = "" Then nOk = nOk + 1 Else rep = rep & vbCrLf & "  ! " & FIG_SHEET & " page " & p & ": " & res
    Next p
Done:
    On Error Resume Next
    If uSet Then UnitsBack uOld, rep
    ActiveAll
    If mPlay = "" Then MaximizeCivil
    If mPlay <> "" Then WriteUtf8 mPlay & "\picture_plan.txt", mPlanLog
    Pictures = "  Pictures in " & FIG_SHEET & ": " & nOk & " of 3 pages " & IIf(mPlay = "", "replaced", "planned (dry run, list in picture_plan.txt)") & rep
    Exit Function
Fail:
    rep = rep & vbCrLf & "  ! STOPPED: " & Err.Description
    Resume Done
End Function

' one page: captures its pictures, then puts them into the frame of the page; page 3 strips 2-5
' (stage results) get a labelled place to paste the picture by hand; when a capture fails the page
' is left as it was; "" = done
Private Function PicturePage(ws As Worksheet, ByVal page As Long, ByVal capRow As Long, ByVal folder As String) As String
    Dim fr As Variant, n As Long, i As Long, f() As String, lbl() As String, res As String, man() As Boolean
    fr = PageFrame(ws, page, capRow)
    If IsEmpty(fr) Then PicturePage = "no picture on the page (rows " & capRow - 27 & "-" & capRow - 1 & "), nothing captured": Exit Function
    n = Choose(page, 1, 1, 5)
    ReDim f(1 To n): ReDim lbl(1 To n): ReDim man(1 To n)
    For i = 1 To n
        man(i) = (page = 3 And i > 1)
        If man(i) Then
            lbl(i) = SlotText(i)
            mPlanLog = mPlanLog & vbLf & "page " & page & ", strip " & i & " / " & n & ": by hand: " & Replace(lbl(i), vbLf, " / ")
        Else
            Note 97, "picture: page " & page & IIf(n > 1, ", strip " & i & " / " & n, "")
            f(i) = folder & "\RTI_Y_" & page & "_" & i & ".png"
            res = CaptureStrip(page, i, n, fr, f(i), lbl(i))
            If res <> "" Then PicturePage = IIf(n > 1, "strip " & i & ": ", "") & res & "; the page was not changed": Exit Function
        End If
    Next i
    If mPlay <> "" Then Exit Function
    PicturePage = PlaceStrips(ws, page, capRow, fr, f, lbl, man, n)
End Function

' one capture; "" = the image file is there (dry run: the request is only listed)
Private Function CaptureStrip(ByVal page As Long, ByVal i As Long, ByVal n As Long, fr As Variant, ByVal f As String, ByRef lbl As String) As String
    Dim kind As String, cType As String, stg As String, cx As Double, cy As Double, trk As Long, axl As String
    Dim eL As String, nL As String, axDeg As Double, hor As Double, vert As Double, cnt As Long
    Dim frameAsp As Double, asp As Double, w As Long, h As Long, zoom As Long, z2 As Long, body As String, resp As String
    Dim pass As Long, pw As Long, ph As Long, pMode As String, bb As String
    bb = "en b" & ChrW(252) & "y" & ChrW(252) & "k "
    Select Case page
        Case 1
            kind = "model"
        Case 2
            kind = "solid": axl = gDbAx
            If Not AxisPoint(axl, cx, cy) Then CaptureStrip = "no deck end axis for the close-up": Exit Function
        Case 3
            kind = "stress"
            Select Case i
                Case 1: cType = "MV": ElemPoint gTenE, cx, cy, trk: lbl = "LM71"
                Case 2: cType = "CS": stg = gTenWin: ElemPoint gTenE, cx, cy, trk: lbl = "Fren + ivme, " & stg & " (" & bb & ChrW(231) & "ekme)"
                Case 3: cType = "CS": stg = gCmpWin: ElemPoint gCmpE, cx, cy, trk: lbl = "Fren + ivme, " & stg & " (" & bb & "bas" & ChrW(305) & "n" & ChrW(231) & ")"
                Case 4: cType = "CS": stg = "TU+": ElemPoint gTenE, cx, cy, trk: lbl = "TU+"
                Case Else: cType = "CS": stg = "TU-": ElemPoint gCmpE, cx, cy, trk: lbl = "TU-"
            End Select
            If trk = 0 Then CaptureStrip = "no governing rail element": Exit Function
            lbl = lbl & ", hat " & trk & ", Sax (N/mm" & ChrW(178) & ")"
        Case Else
            kind = "disp"
            Select Case i
                Case 1: cType = "MV": axl = gDhAx: lbl = "LM71"
                Case 2: cType = "CS": stg = gDbWin: axl = gDbAx: lbl = "Fren + ivme, " & stg & " (" & bb & ChrW(948) & "B)"
                Case Else: cType = "CS": stg = gDaWin: axl = gDaAx: lbl = "Fren + ivme, " & stg & " (" & bb & ChrW(948) & "abs)"
            End Select
            lbl = lbl & ", aks " & axl & ", DX (mm)"
            If Not AxisPoint(axl, cx, cy) Then CaptureStrip = "no deck end axis " & axl: Exit Function
    End Select
    If cType = "CS" And stg = "" Then CaptureStrip = "no governing window": Exit Function
    ' active part, its direction and extents, camera
    frameAsp = fr(2) / (fr(3) / n)
    If kind = "model" Then
        UseAllNodes
        axDeg = mAxDeg: cnt = aeN
    Else
        cnt = PicRegion(kind, cx, cy, trk, eL, nL, axDeg)
        If cnt = 0 Then CaptureStrip = "no elements in the region": Exit Function
    End If
    If kind = "model" Or kind = "solid" Then
        hor = CAM_H - axDeg: vert = CAM_V: pMode = "pre"
    Else
        hor = -axDeg: vert = CAM_V_STRIP: pMode = "post"
    End If
    hor = hor - 360 * Int(hor / 360)
    PicExtents hor, vert
    If mResW > 0 Then
        w = mResW: h = CLng(mResW / frameAsp): asp = w / h
    ElseIf mAspect > 0 Then
        asp = mAspect
    Else
        asp = frameAsp
    End If
    zoom = ZoomFor(asp, frameAsp)
    ' side view of the flat rail: Civil NX's own zoom fit (the computed zoom 181 showed about 65 m
    ' instead of the planned 120 m in the owner's run 2026-11-09)
    If kind = "stress" Then zoom = 100
    mPlanLog = mPlanLog & vbLf & "page " & page & ", strip " & i & " / " & n & ": " & kind & " " & cType & " " & stg & ", " & cnt & " elements, direction " & JNum(Round(axDeg, 2)) & " deg, frame " & JNum(Round(frameAsp, 2)) & ":1, zoom " & zoom & IIf(w > 0, ", " & w & " x " & h & " px", "") & ", label: " & lbl
    ' Civil NX drops ZOOM_LEVEL on the capture that switches between model and result view
    ' (MIDAS_SEKIL, measured 2026-11-09): that capture is sent once more
    If pMode <> mLastMode Then
        body = CapBody(kind, cType, stg, Replace(f, "\", "\\"), w, h, zoom, hor, vert, eL, nL)
        If Not Api("POST", "view/CAPTURE", body, resp) Then
            ' the schema types LIMIT_SCALE as an integer: a refused capture is sent once more without it
            If InStr(body, "LIMIT_SCALE") = 0 Then CaptureStrip = "Civil NX: " & Left$(resp, 200): Exit Function
            mNoLimit = True
            If mPlay = "" And mRec = "" Then UiLog "! Civil NX refused LIMIT_SCALE " & JNum(LIMIT_PCT) & ", values shown without it: " & Left$(resp, 120)
            body = CapBody(kind, cType, stg, Replace(f, "\", "\\"), w, h, zoom, hor, vert, eL, nL)
            If Not Api("POST", "view/CAPTURE", body, resp) Then CaptureStrip = "Civil NX: " & Left$(resp, 200): Exit Function
        End If
        If mPlay = "" Then
            Call WaitForFile(f, WAIT_S)
            If mResW > 0 Then MaximizeCivil
        End If
        mLastMode = pMode
    End If
    For pass = 1 To 2
        If mPlay = "" Then
            If Len(Dir(f)) > 0 Then
                On Error Resume Next
                Kill f
                On Error GoTo 0
                If Len(Dir(f)) > 0 Then CaptureStrip = "cannot delete the old image " & f: Exit Function
            End If
        End If
        body = CapBody(kind, cType, stg, Replace(f, "\", "\\"), w, h, zoom, hor, vert, eL, nL)
        If Not Api("POST", "view/CAPTURE", body, resp) Then
            ' the schema types LIMIT_SCALE as an integer: a refused capture is sent once more without it
            If InStr(body, "LIMIT_SCALE") = 0 Then CaptureStrip = "Civil NX: " & Left$(resp, 200): Exit Function
            mNoLimit = True
            If mPlay = "" And mRec = "" Then UiLog "! Civil NX refused LIMIT_SCALE " & JNum(LIMIT_PCT) & ", values shown without it: " & Left$(resp, 120)
            body = CapBody(kind, cType, stg, Replace(f, "\", "\\"), w, h, zoom, hor, vert, eL, nL)
            If Not Api("POST", "view/CAPTURE", body, resp) Then CaptureStrip = "Civil NX: " & Left$(resp, 200): Exit Function
        End If
        If mPlay <> "" Then Exit For
        If Not WaitForFile(f, WAIT_S) Then
            If mResW > 0 Then MaximizeCivil
            CaptureStrip = "Civil NX accepted the request, but no image appeared": Exit Function
        End If
        If mResW > 0 Then MaximizeCivil
        If mResW > 0 Or mAspect > 0 Then Exit For
        ' screen size: the first image tells the picture shape
        If Not PngSize(f, pw, ph) Then Exit For
        mAspect = pw / ph
        z2 = ZoomFor(mAspect, frameAsp)
        If kind = "stress" Then z2 = 100
        If z2 = zoom Then Exit For
        zoom = z2
    Next pass
End Function

' request body of one capture (manual pages Capture, Active, Beam Stresses Diagram, Displacement Contour)
Private Function CapBody(ByVal kind As String, ByVal cType As String, ByVal stg As String, ByVal pathJ As String, ByVal w As Long, ByVal h As Long, ByVal zoom As Long, ByVal hor As Double, ByVal vert As Double, ByVal eL As String, ByVal nL As String) As String
    Dim b As String
    b = "{""Argument"":{""EXPORT_PATH"":""" & pathJ & ""","
    If w > 0 And h > 0 Then b = b & """WIDTH"":" & w & ",""HEIGHT"":" & h & ","
    b = b & """ZOOM_LEVEL"":" & zoom & ",""ANGLE"":{""HORIZONTAL"":" & JNum(hor) & ",""VERTICAL"":" & JNum(vert) & "},"
    If eL = "" Then
        b = b & """ACTIVE"":{""ACTIVE_MODE"":""All""},"
    Else
        b = b & """ACTIVE"":{""ACTIVE_MODE"":""Active"",""N_LIST"":[" & nL & "],""E_LIST"":[" & eL & "]},"
    End If
    Select Case kind
        Case "model"
            ' elastic link symbols cover a long bridge (MIDAS_SEKIL), so supports only
            b = b & """SET_MODE"":""pre"",""SET_HIDDEN"":false,""DISPLAY"":{""BOUNDARY"":{""SUPPORT"":true,""ELASTIC_LINK"":false}}"
        Case "solid"
            b = b & """SET_MODE"":""pre"",""SET_HIDDEN"":true,""DISPLAY"":{""BOUNDARY"":{""SUPPORT"":false,""ELASTIC_LINK"":false}}"
        Case Else
            b = b & """STAGE_NAME"":" & QJ(IIf(cType = "MV", MV_STAGE, stg)) & ","
            b = b & """SET_MODE"":""post"",""SET_HIDDEN"":" & IIf(kind = "disp", "true", "false") & ","
            b = b & """DISPLAY"":{""BOUNDARY"":{""SUPPORT"":false,""ELASTIC_LINK"":false}},"
            b = b & """RESULT_GRAPHIC"":{""CURRENT_MODE"":" & IIf(kind = "stress", """BeamStressesDiagram""", """DisplacementContour""") & ","
            If cType = "MV" Then
                b = b & """LOAD_CASE_COMB"":{""TYPE"":""MV"",""NAME"":" & QJ(MV_CASE) & ",""MINMAX"":" & QJ(MV_MINMAX) & "},"
            Else
                b = b & """LOAD_CASE_COMB"":{""TYPE"":""CS"",""NAME"":" & QJ(CS_CASE) & "},"
            End If
            If kind = "stress" Then
                b = b & """COMPONENTS"":{""PART"":""Total"",""COMP"":""Sax""},""DISPLAY_OPTIONS"":{""SCALE"":" & JNum(mDiag) & ",""FILL"":""Solid""},"
            Else
                b = b & """COMPONENTS"":{""COMP"":""DX""},"
            End If
            b = b & """TYPE_OF_DISPLAY"":{""CONTOUR"":{""OPT_CHECK"":true,""NUM_OF_COLOR"":" & N_COLOR & ",""COLOR_TYPE"":""rgb""},"
            If kind = "disp" Then b = b & """DEFORM"":{""OPT_CHECK"":false},"
            b = b & """VALUES"":{""OPT_CHECK"":true,""DECIMAL_PT"":" & IIf(kind = "stress", 1, 2) & ",""MINMAX_ONLY"":{""MAXMIN"":""Min&Max""" & IIf(mNoLimit, "", ",""LIMIT_SCALE"":" & JNum(LIMIT_PCT)) & "}},"
            b = b & """LEGEND"":{""OPT_CHECK"":false}}}"
    End Select
    CapBody = b & "}}"
End Function

' ---- geometry for the pictures ----
' bridge direction (principal axis of all nodes), station origin and the span length
' (median distance of consecutive axes along the bridge)
Private Sub PicGeometry()
    Dim i As Long, j As Long, n As Long, key As Variant, cx As Double, cy As Double, t As Double
    Dim st() As Double, d() As Double
    UseAllNodes
    mX0 = 0: mY0 = 0
    For i = 1 To gnN
        mX0 = mX0 + gnX(i): mY0 = mY0 + gnY(i)
    Next i
    If gnN > 0 Then mX0 = mX0 / gnN: mY0 = mY0 / gnN
    mAxDeg = AxisOfSet(0)
    mSpan = 30
    If mDeK Is Nothing Then Exit Sub
    If mDeK.Count < 2 Then Exit Sub
    ReDim st(1 To mDeK.Count)
    For Each key In mDeK.keys
        If AxisPoint(CStr(key), cx, cy) Then n = n + 1: st(n) = Sta(cx, cy)
    Next key
    If n < 2 Then Exit Sub
    For i = 1 To n - 1
        For j = i + 1 To n
            If st(j) < st(i) Then t = st(i): st(i) = st(j): st(j) = t
        Next j
    Next i
    ReDim d(1 To n - 1)
    For i = 1 To n - 1: d(i) = st(i + 1) - st(i): Next i
    For i = 1 To n - 2
        For j = i + 1 To n - 1
            If d(j) < d(i) Then t = d(i): d(i) = d(j): d(j) = t
        Next j
    Next i
    mSpan = d(n \ 2)
    If mSpan <= 0 Then mSpan = 30
End Sub

Private Sub UseAllNodes()
    Dim i As Long
    rgN = gnN
    ReDim rgIdx(1 To IIf(gnN > 0, gnN, 1))
    For i = 1 To gnN: rgIdx(i) = i: Next i
End Sub

' elements within +-REG_SPANS spans of the point along the bridge (stress: the rail of track trk,
' disp: all but the rail, solid: all); their nodes become the active set; returns the count
Private Function PicRegion(ByVal kind As String, ByVal cx As Double, ByVal cy As Double, ByVal trk As Long, ByRef eL As String, ByRef nL As String, ByRef axDeg As Double) As Long
    Dim s0 As Double, rr As Double, i As Long, ka As Long, kb As Long, sm As Double, take As Boolean, t As Long
    Dim seen As Object, key As Variant
    s0 = Sta(cx, cy): rr = REG_SPANS * mSpan
    If kind = "stress" Then rr = BIG          ' rail stress strip: the whole rail of the track (owner 2026-11-09)
    Set seen = CreateObject("Scripting.Dictionary")
    eL = "": nL = ""
    For i = 1 To aeN
        If gnIx.Exists(aeA(i)) And gnIx.Exists(aeB(i)) Then
            ka = gnIx(aeA(i)): kb = gnIx(aeB(i))
            sm = (Sta(gnX(ka), gnY(ka)) + Sta(gnX(kb), gnY(kb))) / 2
            If Abs(sm - s0) <= rr Then
                t = 0
                If elIx.Exists(aeId(i)) Then t = eTrk(elIx(aeId(i)))
                Select Case kind
                    Case "stress": take = (t = trk)
                    Case "disp": take = (t = 0)
                    Case Else: take = True
                End Select
                If take Then
                    eL = eL & IIf(eL = "", "", ",") & aeId(i)
                    PicRegion = PicRegion + 1
                    If Not seen.Exists(ka) Then seen(ka) = 1: nL = nL & IIf(nL = "", "", ",") & aeA(i)
                    If Not seen.Exists(kb) Then seen(kb) = 1: nL = nL & IIf(nL = "", "", ",") & aeB(i)
                End If
            End If
        End If
    Next i
    rgN = 0
    ReDim rgIdx(1 To IIf(seen.Count > 0, seen.Count, 1))
    For Each key In seen.keys
        rgN = rgN + 1: rgIdx(rgN) = key
    Next key
    axDeg = AxisOfSet(mAxDeg)
End Function

' principal axis of the active nodes in degrees from +X, turned to within 90 deg of ref
Private Function AxisOfSet(ByVal ref As Double) As Double
    Dim i As Long, k As Long, mx As Double, my As Double, sxx As Double, syy As Double, sxy As Double, a As Double
    AxisOfSet = ref
    If rgN < 2 Then Exit Function
    For i = 1 To rgN
        k = rgIdx(i): mx = mx + gnX(k): my = my + gnY(k)
    Next i
    mx = mx / rgN: my = my / rgN
    For i = 1 To rgN
        k = rgIdx(i)
        sxx = sxx + (gnX(k) - mx) ^ 2: syy = syy + (gnY(k) - my) ^ 2: sxy = sxy + (gnX(k) - mx) * (gnY(k) - my)
    Next i
    a = 0.5 * Atan2(2 * sxy, sxx - syy) * 45 / Atn(1)
    If Cos((a - ref) * Atn(1) / 45) < 0 Then a = a + 180
    AxisOfSet = a
End Function

Private Function Sta(ByVal x As Double, ByVal y As Double) As Double
    Dim r As Double
    r = mAxDeg * Atn(1) / 45
    Sta = (x - mX0) * Cos(r) + (y - mY0) * Sin(r)
End Function

Private Function AxisPoint(ByVal lab As String, ByRef cx As Double, ByRef cy As Double) As Boolean
    Dim a As Variant, c As Long, j As Long
    If mDeK Is Nothing Or lab = "" Then Exit Function
    If Not mDeK.Exists(lab) Then Exit Function
    a = mDeK(lab): cx = 0: cy = 0
    For j = 0 To 1
        If a(j) > 0 Then cx = cx + nX(a(j)): cy = cy + nY(a(j)): c = c + 1
    Next j
    If c = 0 Then Exit Function
    cx = cx / c: cy = cy / c
    AxisPoint = True
End Function

' middle of rail element e (index into eId) and its track
Private Function ElemPoint(ByVal e As Long, ByRef cx As Double, ByRef cy As Double, ByRef trk As Long) As Boolean
    Dim k As Long, ka As Long, kb As Long
    trk = 0
    If e < 1 Or e > nEl Then Exit Function
    If Not aeIx.Exists(eId(e)) Then Exit Function
    k = aeIx(eId(e))
    If Not (gnIx.Exists(aeA(k)) And gnIx.Exists(aeB(k))) Then Exit Function
    ka = gnIx(aeA(k)): kb = gnIx(aeB(k))
    cx = (gnX(ka) + gnX(kb)) / 2: cy = (gnY(ka) + gnY(kb)) / 2
    trk = eTrk(e)
    ElemPoint = True
End Function

' projected width and height of the active nodes for the camera, and their 3D diagonal
' (camera at plan bearing 270 - HORIZONTAL from +X, elevation VERTICAL; MIDAS_SEKIL)
Private Sub PicExtents(ByVal hor As Double, ByVal vert As Double)
    Dim i As Long, k As Long, dg As Double, bt As Double, ev As Double
    Dim rx As Double, ry As Double, ux As Double, uy As Double, uz As Double, pr As Double, pu As Double
    Dim r0 As Double, r1 As Double, u0 As Double, u1 As Double
    Dim x0 As Double, x1 As Double, y0 As Double, y1 As Double, z0 As Double, z1 As Double
    mWm = 0: mHm = 0: mD = 0
    If rgN < 1 Then Exit Sub
    dg = Atn(1) / 45
    bt = (270 - hor) * dg: ev = vert * dg
    rx = -Sin(bt): ry = Cos(bt)
    ux = -Sin(ev) * Cos(bt): uy = -Sin(ev) * Sin(bt): uz = Cos(ev)
    For i = 1 To rgN
        k = rgIdx(i)
        pr = gnX(k) * rx + gnY(k) * ry
        pu = gnX(k) * ux + gnY(k) * uy + gnZ(k) * uz
        If i = 1 Then
            r0 = pr: r1 = pr: u0 = pu: u1 = pu
            x0 = gnX(k): x1 = x0: y0 = gnY(k): y1 = y0: z0 = gnZ(k): z1 = z0
        Else
            If pr < r0 Then r0 = pr
            If pr > r1 Then r1 = pr
            If pu < u0 Then u0 = pu
            If pu > u1 Then u1 = pu
            If gnX(k) < x0 Then x0 = gnX(k)
            If gnX(k) > x1 Then x1 = gnX(k)
            If gnY(k) < y0 Then y0 = gnY(k)
            If gnY(k) > y1 Then y1 = gnY(k)
            If gnZ(k) < z0 Then z0 = gnZ(k)
            If gnZ(k) > z1 Then z1 = gnZ(k)
        End If
    Next i
    mWm = r1 - r0: mHm = u1 - u0
    mD = Sqr((x1 - x0) ^ 2 + (y1 - y0) ^ 2 + (z1 - z0) ^ 2)
End Sub

' zoom at which the active part fills FILL of the kept part of the image (MIDAS_SEKIL); a flat
' part (rail in a side view, height 0) is fitted by its width alone; 100 when unknown
Private Function ZoomFor(ByVal aspect As Double, ByVal frameAsp As Double) As Long
    Dim fw As Double, fh As Double, s As Double, zz As Double
    ZoomFor = 100
    If mD <= 0 Or mWm <= 0 Or mHm < 0 Or aspect <= 0 Or frameAsp <= 0 Then Exit Function
    If aspect >= frameAsp Then
        fw = mWm / (K_FIT * mD * frameAsp): fh = mHm / (K_FIT * mD)
    Else
        fw = mWm / (K_FIT * mD * aspect): fh = mHm / (K_FIT * mD) * frameAsp / aspect
    End If
    s = FILL / fw
    If fh > 0 Then
        If FILL / fh < s Then s = FILL / fh
    End If
    zz = 100 + Log(s) / ZOOM_RATE
    If zz < 25 Then zz = 25
    If zz > 200 Then zz = 200
    ZoomFor = CLng(zz)
End Function

' ---- the sheet ----
' frame of a page: kept by the macro in the alternative text of its shapes (RTI_FRAME|l|t|w|h);
' first run: the box around the pictures and picture groups in rows capRow-27 .. capRow-1; Empty
' when there is none
Private Function PageFrame(ws As Worksheet, ByVal page As Long, ByVal capRow As Long) As Variant
    Dim shp As Shape, r As Long, l0 As Double, t0 As Double, l1 As Double, t1 As Double, n As Long, v As Variant
    For Each shp In ws.Shapes
        If IsOurs(shp.Name, page) Then
            v = Split(shp.AlternativeText, "|")
            If UBound(v) = 4 Then
                If v(0) = FRAME_TAG Then PageFrame = Array(Val(v(1)), Val(v(2)), Val(v(3)), Val(v(4))): Exit Function
            End If
        End If
    Next shp
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Or shp.Type = msoGroup Then
            r = 0
            On Error Resume Next
            r = shp.TopLeftCell.Row
            On Error GoTo 0
            If r >= capRow - 27 And r < capRow Then
                If n = 0 Then
                    l0 = shp.Left: t0 = shp.Top: l1 = shp.Left + shp.Width: t1 = shp.Top + shp.Height
                Else
                    If shp.Left < l0 Then l0 = shp.Left
                    If shp.Top < t0 Then t0 = shp.Top
                    If shp.Left + shp.Width > l1 Then l1 = shp.Left + shp.Width
                    If shp.Top + shp.Height > t1 Then t1 = shp.Top + shp.Height
                End If
                n = n + 1
            End If
        End If
    Next shp
    If n > 0 Then PageFrame = Array(l0, t0, l1 - l0, t1 - t0)
End Function

Private Function IsOurs(ByVal nm As String, ByVal page As Long) As Boolean
    Dim pre As Variant
    For Each pre In Array(PIC_PREFIX, LBL_PREFIX, SLOT_PREFIX)
        If Left$(nm, Len(pre & page & "_")) = pre & page & "_" Then IsOurs = True: Exit Function
    Next pre
End Function

Private Function FrameTag(fr As Variant) As String
    FrameTag = FRAME_TAG & "|" & Trim$(Str$(fr(0))) & "|" & Trim$(Str$(fr(1))) & "|" & Trim$(Str$(fr(2))) & "|" & Trim$(Str$(fr(3)))
End Function

' the page's new pictures and hand-paste places stacked in the frame, then the old ones deleted:
' the macro's own shapes always, the other pictures of the page only on the first run (the original
' layout); a picture pasted by hand over a place is kept and that place is not drawn again
Private Function PlaceStrips(ws As Worksheet, ByVal page As Long, ByVal capRow As Long, fr As Variant, f() As String, lbl() As String, man() As Boolean, ByVal n As Long) As String
    Dim i As Long, shp As Shape, newS() As Shape, hs As Double, old As New Collection, plc As Long, r As Long
    Dim firstRun As Boolean, tg As String
    hs = fr(3) / n
    plc = 1
    tg = FrameTag(fr)
    firstRun = True
    For Each shp In ws.Shapes
        If IsOurs(shp.Name, page) Then firstRun = False: Exit For
    Next shp
    For Each shp In ws.Shapes
        r = 0
        On Error Resume Next
        r = shp.TopLeftCell.Row
        On Error GoTo 0
        If IsOurs(shp.Name, page) Then
            old.Add shp
            If Left$(shp.Name, Len(PIC_PREFIX)) = PIC_PREFIX Then plc = shp.Placement
        ElseIf firstRun And (shp.Type = msoPicture Or shp.Type = msoLinkedPicture Or shp.Type = msoGroup) And r >= capRow - 27 And r < capRow Then
            old.Add shp: plc = shp.Placement
        End If
    Next shp
    ReDim newS(1 To n)
    On Error GoTo Bad
    For i = 1 To n
        If Not man(i) Then
            Set newS(i) = ws.Shapes.AddPicture(f(i), msoFalse, msoTrue, fr(0), fr(1) + (i - 1) * hs, -1, -1)
            FitInto newS(i), fr(0), fr(1) + (i - 1) * hs, fr(2), hs
            newS(i).Placement = plc
            newS(i).AlternativeText = tg
        End If
    Next i
    On Error GoTo 0
    For Each shp In old
        shp.Delete
    Next shp
    For i = 1 To n
        If man(i) Then
            If Not HandPicture(ws, page, fr(0) + fr(2) / 2, fr(1) + (i - 0.5) * hs) Then AddSlot ws, page, i, lbl(i), fr(0), fr(1) + (i - 1) * hs, fr(2), hs, tg
        Else
            newS(i).Name = PIC_PREFIX & page & "_" & i
            If n > 1 Then AddLabel ws, page, i, lbl(i), fr(0), fr(1) + (i - 1) * hs
        End If
    Next i
    Exit Function
Bad:
    PlaceStrips = "captured, but a picture could not be put in (" & Err.Description & "); the old one was kept"
    On Error Resume Next
    For i = 1 To n
        If Not newS(i) Is Nothing Then newS(i).Delete
    Next i
End Function

' a picture the macro did not make covering the point (pasted by hand)
Private Function HandPicture(ws As Worksheet, ByVal page As Long, ByVal px As Double, ByVal py As Double) As Boolean
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Not IsOurs(shp.Name, page) Then
            If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Or shp.Type = msoGroup Then
                If px >= shp.Left And px <= shp.Left + shp.Width And py >= shp.Top And py <= shp.Top + shp.Height Then HandPicture = True: Exit Function
            End If
        End If
    Next shp
End Function

' dashed box with the text of what to capture by hand
Private Sub AddSlot(ws As Worksheet, ByVal page As Long, ByVal i As Long, ByVal txt As String, ByVal bl As Double, ByVal bt As Double, ByVal bw As Double, ByVal bh As Double, ByVal tg As String)
    Dim sh As Shape
    Set sh = ws.Shapes.AddShape(1, bl, bt, bw, bh)
    With sh
        .Name = SLOT_PREFIX & page & "_" & i
        .Placement = 1
        .AlternativeText = tg
        .Fill.ForeColor.RGB = RGB(248, 248, 248)
        .Line.ForeColor.RGB = RGB(150, 150, 150): .Line.Weight = 0.75: .Line.DashStyle = 4
        With .TextFrame2
            .WordWrap = msoTrue
            .VerticalAnchor = 3
            .TextRange.Text = txt
            .TextRange.ParagraphFormat.Alignment = 2
            .TextRange.Font.Size = 8
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Fill.ForeColor.RGB = RGB(70, 70, 70)
        End With
    End With
End Sub

' text of a place for a stage picture pasted by hand (page 3, strips 2-5)
Private Function SlotText(ByVal i As Long) As String
    Dim e As Long, stg As String, what As String, cx As Double, cy As Double, trk As Long, axn As String, bb As String
    bb = "en b" & ChrW(252) & "y" & ChrW(252) & "k "
    Select Case i
        Case 2: e = gTenE: stg = gTenWin: what = "Fren + ivme, " & stg & " (" & bb & ChrW(231) & "ekme)"
        Case 3: e = gCmpE: stg = gCmpWin: what = "Fren + ivme, " & stg & " (" & bb & "bas" & ChrW(305) & "n" & ChrW(231) & ")"
        Case 4: e = gTenE: stg = "TU+": what = "TU+"
        Case Else: e = gCmpE: stg = "TU-": what = "TU-"
    End Select
    ElemPoint e, cx, cy, trk
    axn = AxisNear(cx, cy)        ' the governing rail element is near this axis
    SlotText = what & ", hat " & trk & ", Sax (N/mm" & ChrW(178) & ")" & vbLf & _
        "Elle: Civil NX'te a" & ChrW(351) & "ama " & stg & ", CS Summation, Beam Stresses Diagram Sax, hat " & trk & " ray" & ChrW(305) & ", t" & ChrW(252) & "m k" & ChrW(246) & "pr" & ChrW(252) & " (belirleyici eleman aks " & axn & " yak" & ChrW(305) & "n" & ChrW(305) & "nda)" & _
        "; resmi bu kutunun yerine yap" & ChrW(305) & ChrW(351) & "t" & ChrW(305) & "r" & ChrW(305) & "n."
End Function

' label of the axis nearest to the point
Private Function AxisNear(ByVal px As Double, ByVal py As Double) As String
    Dim key As Variant, ax2 As Double, ay2 As Double, dd As Double, bd As Double
    bd = BIG
    If mDeK Is Nothing Then Exit Function
    For Each key In mDeK.keys
        If AxisPoint(CStr(key), ax2, ay2) Then
            dd = (ax2 - px) ^ 2 + (ay2 - py) ^ 2
            If dd < bd Then bd = dd: AxisNear = CStr(key)
        End If
    Next key
End Function

' crop the image (centred) to the shape of the box, then give it the box
Private Sub FitInto(shp As Shape, ByVal bl As Double, ByVal bt As Double, ByVal bw As Double, ByVal bh As Double)
    Dim w0 As Double, h0 As Double, fa As Double, c As Double
    w0 = shp.Width: h0 = shp.Height: fa = bw / bh
    If w0 / h0 > fa Then
        c = (w0 - h0 * fa) / 2
        shp.PictureFormat.CropLeft = c: shp.PictureFormat.CropRight = c
    Else
        c = (h0 - w0 / fa) / 2
        shp.PictureFormat.CropTop = c: shp.PictureFormat.CropBottom = c
    End If
    shp.LockAspectRatio = msoFalse
    shp.Width = bw: shp.Height = bh: shp.Left = bl: shp.Top = bt
    shp.LockAspectRatio = msoTrue
End Sub

Private Sub AddLabel(ws As Worksheet, ByVal page As Long, ByVal i As Long, ByVal txt As String, ByVal bl As Double, ByVal bt As Double)
    Dim tb As Shape
    Set tb = ws.Shapes.AddTextbox(1, bl + 2, bt + 2, 300, 12)
    With tb
        .Name = LBL_PREFIX & page & "_" & i
        .Placement = 1
        .Line.Visible = msoFalse
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Fill.Transparency = 0.2
        With .TextFrame2
            .WordWrap = msoFalse
            .AutoSize = 1
            .MarginLeft = 2: .MarginRight = 2: .MarginTop = 1: .MarginBottom = 1
            .TextRange.Text = txt
            .TextRange.Font.Size = 7
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        End With
    End With
End Sub

' ---- Civil NX state: units, active elements, window ----
' unit system to N, mm for the result pictures; the old record is kept for UnitsBack
Private Function UnitsTo(ByRef old As String) As Boolean
    Dim resp As String, body As String
    If Not Api("GET", "db/UNIT", "", resp) Then Exit Function
    If JStr(resp, "FORCE") = "" Then Exit Function
    old = resp
    body = "{""Assign"":{""1"":{""FORCE"":""N"",""DIST"":""MM"",""HEAT"":" & QJ(JStr(old, "HEAT")) & ",""TEMPER"":" & QJ(JStr(old, "TEMPER")) & "}}}"
    UnitsTo = Api("PUT", "db/UNIT", body, resp)
End Function

Private Sub UnitsBack(ByVal old As String, ByRef rep As String)
    Dim body As String, resp As String, ok As Boolean
    body = "{""Assign"":{""1"":{""FORCE"":" & QJ(JStr(old, "FORCE")) & ",""DIST"":" & QJ(JStr(old, "DIST")) & ",""HEAT"":" & QJ(JStr(old, "HEAT")) & ",""TEMPER"":" & QJ(JStr(old, "TEMPER")) & "}}}"
    ok = Api("PUT", "db/UNIT", body, resp)
    If ok And mPlay = "" Then
        ok = Api("GET", "db/UNIT", "", resp)
        If ok Then ok = (JStr(resp, "FORCE") = JStr(old, "FORCE") And JStr(resp, "DIST") = JStr(old, "DIST"))
    End If
    If Not ok Then rep = rep & vbCrLf & "  ! the unit system was not set back: set it to " & JStr(old, "FORCE") & ", " & JStr(old, "DIST") & " in Civil NX"
End Sub

Private Sub ActiveAll()
    Dim resp As String
    Call Api("POST", "view/ACTIVE", "{""Argument"":{""ACTIVE_MODE"":""All""}}", resp)
End Sub

Private Sub MaximizeCivil()
    Dim hw As LongPtr, buf As String, n As Long
    On Error Resume Next
    hw = GetWindow(GetDesktopWindow(), 5)          ' GW_CHILD: first top-level window
    Do While hw <> 0
        If IsWindowVisible(hw) <> 0 Then
            buf = String$(512, vbNullChar)
            n = GetWindowTextW(hw, StrPtr(buf), 512)
            If n > 0 Then
                If InStr(1, Left$(buf, n), "MIDAS CIVIL NX", vbTextCompare) = 1 Then ShowWindow hw, 3: Exit Sub
            End If
        End If
        hw = GetWindow(hw, 2)                       ' GW_HWNDNEXT
    Loop
End Sub

Private Function ResWidth(ByVal choice As Long) As Long
    Dim v As Variant
    v = Split(RES_W, ",")
    If choice >= 1 And choice <= UBound(v) + 1 Then ResWidth = CLng(v(choice - 1))
End Function

Private Function WaitForFile(ByVal f As String, ByVal secs As Long) As Boolean
    Dim t0 As Single, n1 As Long, n2 As Long
    t0 = Timer
    Do
        If Len(Dir(f)) > 0 Then
            n1 = FileLen(f)
            Pause 0.5
            n2 = FileLen(f)
            If n1 > 0 And n1 = n2 Then WaitForFile = True: Exit Function
        End If
        If Timer < t0 Then t0 = t0 - 86400!
        If Timer - t0 > secs Then Exit Function
        Pause 0.2
    Loop
End Function

Private Sub Pause(ByVal s As Single)
    Dim t0 As Single
    t0 = Timer
    Do While Timer - t0 < s And Timer >= t0
        DoEvents
    Loop
End Sub

' width and height of a PNG file from its IHDR header
Private Function PngSize(ByVal f As String, ByRef w As Long, ByRef h As Long) As Boolean
    Dim fn As Integer, b(0 To 23) As Byte
    On Error GoTo Bad
    fn = FreeFile
    Open f For Binary Access Read As #fn
    Get #fn, 1, b
    Close #fn
    If b(1) <> 80 Or b(2) <> 78 Or b(3) <> 71 Then Exit Function
    w = ((CLng(b(16)) * 256 + b(17)) * 256 + b(18)) * 256 + b(19)
    h = ((CLng(b(20)) * 256 + b(21)) * 256 + b(22)) * 256 + b(23)
    PngSize = (w > 0 And h > 0)
    Exit Function
Bad:
    On Error Resume Next
    Close #fn
End Function

Private Function Atan2(ByVal y As Double, ByVal x As Double) As Double
    Const PI As Double = 3.14159265358979
    If x > 0 Then
        Atan2 = Atn(y / x)
    ElseIf x < 0 Then
        Atan2 = Atn(y / x) + IIf(y >= 0, PI, -PI)
    Else
        Atan2 = IIf(y > 0, PI / 2, IIf(y < 0, -PI / 2, 0))
    End If
End Function

Private Function JNum(ByVal v As Double) As String
    Dim s As String
    s = Trim$(Str$(Round(v, 6)))
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    JNum = s
End Function

Private Function QJ(ByVal s As String) As String
    QJ = """" & Replace(s, """", "\""") & """"
End Function

' value of a flat string field of a JSON answer (first occurrence)
Private Function JStr(ByVal s As String, ByVal key As String) As String
    Dim p As Long, q As Long
    p = InStr(1, s, """" & key & """", vbBinaryCompare)
    If p = 0 Then Exit Function
    p = InStr(p + Len(key) + 2, s, """", vbBinaryCompare)
    If p = 0 Then Exit Function
    q = InStr(p + 1, s, """", vbBinaryCompare)
    If q > p Then JStr = Mid$(s, p + 1, q - p - 1)
End Function

' pictures without the read: the governing windows and places come from RTI_SONUC (E / F)
Private Function PicturesOnly(ByVal ask As Boolean, ByVal pj As String) As String
    Dim ws As Worksheet, msg As String, stamp As String, cm As String, r As String
    Set ws = ThisWorkbook.Worksheets(OUT_SHEET)
    If Not FromSheet(ws, msg) Then PicturesOnly = "STOPPED: " & msg: Exit Function
    stamp = CStr(ws.Range("B3").Value)
    cm = JStr(pj, "COMMENT")
    If cm = "" Or InStr(1, stamp, cm, vbTextCompare) = 0 Then
        msg = OUT_SHEET & " was not read from the model open in Civil NX:" & vbCrLf & "  " & OUT_SHEET & " B3: " & IIf(stamp = "", "(no stamp)", stamp) & vbCrLf & "  Civil NX: " & IIf(cm = "", "(no project comment)", cm)
        If Not ask Then PicturesOnly = "STOPPED: " & msg: Exit Function
        If UiMsg(msg & vbCrLf & vbCrLf & "Take the pictures with these windows anyway?", vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) <> vbYes Then PicturesOnly = "Stopped: nothing was captured.": Exit Function
    End If
    If ask Then
        If UiMsg("Pictures only: 3 pictures and 4 places for stage pictures go into " & FIG_SHEET & ", windows and places from " & OUT_SHEET & " (" & gTenWin & ", " & gCmpWin & ", " & gDbWin & ")." & vbCrLf & _
                 "Civil NX shows each view and uses N, mm for the result pictures (set back after). Continue?", vbYesNo + vbQuestion, MSG_TITLE) <> vbYes Then PicturesOnly = "Stopped: nothing was captured.": Exit Function
    End If
    Set mDeK = DeckEnds()
    r = Pictures()
    PicturesOnly = IIf(InStr(r, "!") > 0, "Finished with problems: ", "Done: ") & "pictures from the windows in " & OUT_SHEET & "." & vbCrLf & r
End Function

' the governing places written by the last read: rail elements (E5, E6), windows (F5, F6, F7, F11),
' axes (E7, E8, E11)
Private Function FromSheet(ws As Worksheet, ByRef msg As String) As Boolean
    Dim id As Long
    gTenE = 0: gCmpE = 0
    id = CLng(Val(After(CStr(ws.Range("E5").Value), "eleman ")))
    If elIx.Exists(id) Then gTenE = elIx(id)
    id = CLng(Val(After(CStr(ws.Range("E6").Value), "eleman ")))
    If elIx.Exists(id) Then gCmpE = elIx(id)
    gTenWin = WinOf(ws.Range("F5").Value): gCmpWin = WinOf(ws.Range("F6").Value)
    gDbAx = AxOf(ws.Range("E7").Value): gDbWin = WinOf(ws.Range("F7").Value)
    gDhAx = AxOf(ws.Range("E8").Value)
    gDaAx = AxOf(ws.Range("E11").Value): gDaWin = WinOf(ws.Range("F11").Value)
    If gTenE = 0 Or gCmpE = 0 Or gTenWin = "" Or gCmpWin = "" Or gDbAx = "" Then
        msg = OUT_SHEET & " holds no results of this model (E5:F7 empty or other rail elements): run it with " & ChrW(39) & "Oku + resimle" & ChrW(39) & " first. Nothing was captured."
        Exit Function
    End If
    FromSheet = True
End Function

Private Function After(ByVal s As String, ByVal key As String) As String
    Dim p As Long
    p = InStr(1, s, key, vbTextCompare)
    If p > 0 Then After = Mid$(s, p + Len(key))
End Function

' window name of an F cell ('F_SON_P1: pencere maks + ...' -> F_SON_P1)
Private Function WinOf(ByVal v As Variant) As String
    Dim t As String
    t = Trim$(CStr(v))
    If InStr(t, ":") > 0 Then t = Left$(t, InStr(t, ":") - 1)
    WinOf = Trim$(t)
End Function

' axis of an E cell ('Aks P1, sol tabliye ucu' -> P1)
Private Function AxOf(ByVal v As Variant) As String
    Dim t As String, i As Long
    t = After(CStr(v), "Aks ")
    For i = 1 To Len(t)
        If InStr(1, " ,(", Mid$(t, i, 1)) > 0 Then t = Left$(t, i - 1): Exit For
    Next i
    AxOf = Trim$(t)
End Function
