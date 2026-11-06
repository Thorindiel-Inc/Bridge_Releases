Attribute VB_Name = "DXF_OKU"
Option Explicit
' ============================================================================
'  DXF_OKU - reads the plan of a bridge from the drawn centerline of a plan DXF into the
'  1-3A INPUT table G_AKS_XY (X, Y of every axis in drawing coordinates, A1 first).
'  Only the centerline layer is asked. The centerline may be one LWPOLYLINE (bulges = arcs)
'  or LINE / ARC / LWPOLYLINE pieces that join end to end (TOL_JOIN). SPLINE, ELLIPSE, old
'  POLYLINE and a closed polyline on that layer are refused. Paper space is skipped.
'  Axes = the points where the drawn axis lines cross the centerline (MIDAS_KOPRU makes each
'  span a straight chord between its axes). A1 given by hand: axes at the span lengths along the line.
'  A1: found from the lines of every other layer that cross the centerline - the one run of
'  span count + 1 crossings whose spacing along the line matches the span list (TOL_SPAN),
'  in either direction of the line. If there is no such run, or more than one, the distance
'  along the line from its first point to A1 is asked (negative = from its last point,
'  against the drawing direction).
'  Texts A1, A2, P1, P2 ... near the line are checked against the axis order (warning only).
'  Button macro: DXF_OKU_OLUSTUR. Scripted: DXF_OKU_RUN path, layer, reportFile[, a1Dist]
'  (layer "" = best guess, a1Dist "" = find A1). Nothing is written if a check fails.
' ============================================================================

Private Const SCRIPT_ID As String = "dxf-alignment"
Private Const SCRIPT_VERSION As String = "2026-11-06c"
Private Const SCRIPT_CHANGELOG As String = "Span radius in the column next to G_AKS_XY (centrifugal load of MIDAS_KOPRU); G_AKS_XY default position A224:B264; axes at the drawn crossing points instead of the span list, axis names A1 / A2 / P1.. checked; centerline mode: axis X / Y from the drawn centerline into G_AKS_XY, only its layer asked, A1 from the lines crossing it (else asked), label mode removed; first version: alignment labels and axis stations from a plan DXF into INPUT, layers asked"
Private Const R_STRAIGHT As Double = 50000#     ' a span flatter than this counts as straight (m)

Private Const MSG_TITLE As String = "DXF_OKU"
Private Const TOL_SPAN As Double = 0.02      ' m, crossing lines against the span list
Private Const TOL_JOIN As Double = 0.01      ' m, end points of centerline pieces
Private Const TXT_NEAR As Double = 50#       ' m, station text to the centerline
Private Const SAMPLE_DS As Double = 1#       ' m, sample spacing along the line for station texts
Private Const PI As Double = 3.14159265358979

Private dl() As String, nDl As Long, reSt As Object, reNm As Object
' entity being read
Private eTyp As String, eLay As String, ePaper As Boolean, eZ As Double, eFlag As Long, eClosed As Boolean, eTxt As String
Private nV As Long, vX() As Double, vY() As Double, vB() As Double
Private eX2 As Double, eY2 As Double, eR As Double, eA1 As Double, eA2 As Double
' layers: entity count, longest open entity
Private layN As Object, layLen As Object
' centerline pieces, then the joined chain (chB(i) = bulge of segment i to i + 1, chS = length to point i)
Private nPc As Long, pcFrom() As Long, pcTo() As Long, nQ As Long, qX() As Double, qY() As Double, qB() As Double
Private nBad As Long, badTyp As String
Private nCh As Long, chX() As Double, chY() As Double, chB() As Double, chS() As Double, totLen As Double, nArc As Long
Private bbX0 As Double, bbY0 As Double, bbX1 As Double, bbY1 As Double
Private sgX0() As Double, sgY0() As Double, sgX1() As Double, sgY1() As Double
' crossings (distance along the chain), station texts, spans, result
Private nCr As Long, crS() As Double, nMt As Long, mtS() As Double, mtD() As Long, mtW() As Double
Private nSm As Long, smS() As Double, smX() As Double, smY() As Double, nTx As Long, txS() As Double, txKm() As Double
Private nSp As Long, sp() As Double, cum() As Double
Private axX() As Double, axY() As Double, axS() As Double, a1S As Double, a1Dir As Long, a1How As String, axR() As Double
Private nNm As Long, nameS() As Double, nameT() As String

Public Sub DXF_OKU_OLUSTUR()
    Dim v As Variant, msg As String, lay As String
    On Error GoTo Fail
    v = Application.GetOpenFilename("DXF (*.dxf),*.dxf", , "Plan DXF of the bridge")
    If VarType(v) = vbBoolean Then Exit Sub
    Application.StatusBar = "DXF_OKU: reading the drawing..."
    ReadDxf CStr(v)
    LayerStats
    Application.StatusBar = False
    lay = AskLayer()
    If lay = "" Then GoTo Quit
    Application.StatusBar = "DXF_OKU: reading the centerline..."
    msg = Process(lay, "", True)
    Application.StatusBar = False
    If msg = "" Then GoTo Quit
    If MsgBox(Left$(msg, 850) & vbCrLf & vbCrLf & "Write this to INPUT?", vbYesNo + vbQuestion, MSG_TITLE) = vbYes Then
        WriteTable
        MsgBox "Written to INPUT (G_AKS_XY, " & nSp + 1 & " axes).", vbInformation, MSG_TITLE
    End If
Quit:
    Application.StatusBar = False
    Erase dl
    Exit Sub
Fail:
    Application.StatusBar = False
    Erase dl
    MsgBox "Stopped: " & Err.Description, vbCritical, MSG_TITLE
End Sub

Public Sub DXF_OKU_RUN(ByVal path As String, ByVal layer As String, ByVal reportFile As String, Optional ByVal a1Dist As String = "")
    Dim msg As String, f As Integer
    On Error GoTo Fail
    ReadDxf path
    If layer = "" Then
        LayerStats
        layer = BestLayer()
    End If
    msg = "Layer: " & layer & vbCrLf & Process(layer, a1Dist, False)
    WriteTable
Done:
    Erase dl
    f = FreeFile
    Open reportFile For Output As #f
    Print #f, msg;
    Close #f
    Exit Sub
Fail:
    msg = "FAIL: " & Err.Description
    Resume Done
End Sub

' ---- reading ----------------------------------------------------------------
Private Sub ReadDxf(ByVal path As String)
    Dim st As Object, all As String, fso As Object, tmp As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    tmp = fso.GetSpecialFolder(2) & "\dxf_oku_" & Int(Timer * 100) & ".dxf"
    fso.CopyFile path, tmp, True
    On Error GoTo CleanTmp
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.LoadFromFile tmp
    all = st.ReadText
    st.Close
    On Error Resume Next
    fso.DeleteFile tmp, True
    On Error GoTo 0
    dl = Split(Replace(all, vbCr, ""), vbLf)
    all = ""
    nDl = UBound(dl)
    Exit Sub
CleanTmp:
    Dim d As String: d = Err.Description
    On Error Resume Next
    st.Close
    fso.DeleteFile tmp, True
    On Error GoTo 0
    Err.Raise vbObjectError + 10, , d
End Sub

' mode 1 = layer list, 2 = centerline pieces of layer lay, 3 = crossings of every other layer + station texts
Private Sub Scan(ByVal mode As Long, ByVal lay As String)
    Dim i As Long, code As String, v As String, inEnt As Boolean
    ReDim vX(1 To 64): ReDim vY(1 To 64): ReDim vB(1 To 64)
    eTyp = ""
    For i = 0 To nDl - 1 Step 2
        code = Trim$(dl(i)): v = Trim$(dl(i + 1))
        If Not inEnt Then
            If code = "2" And v = "ENTITIES" Then inEnt = True
        ElseIf code = "0" Then
            If eTyp <> "" Then EndEntity mode, lay
            If v = "ENDSEC" Then Exit For
            StartEntity v, mode
        ElseIf eTyp <> "" Then
            Select Case code
                Case "8": eLay = v
                Case "1", "3": eTxt = eTxt & v
                Case "67": ePaper = (Val(v) = 1)
                Case "70": eFlag = Val(v)
                Case "230": eZ = Val(v)
                Case "10": AddV Val(v)
                Case "20": If nV > 0 Then vY(nV) = Val(v)
                Case "42": If nV > 0 Then vB(nV) = Val(v)
                Case "11": eX2 = Val(v)
                Case "21": eY2 = Val(v)
                Case "40": eR = Val(v)
                Case "50": eA1 = Val(v)
                Case "51": eA2 = Val(v)
            End Select
        End If
    Next i
End Sub

Private Sub StartEntity(ByVal v As String, ByVal mode As Long)
    eTyp = ""
    Select Case v
        Case "LINE", "ARC", "LWPOLYLINE", "SPLINE", "POLYLINE", "ELLIPSE": eTyp = v
        Case "TEXT", "MTEXT", "ATTRIB": If mode = 3 Then eTyp = v
    End Select
    eLay = "": ePaper = False: eZ = 1: eFlag = 0: nV = 0: eClosed = False: eTxt = ""
    eX2 = 0: eY2 = 0: eR = 0: eA1 = 0: eA2 = 0
End Sub

Private Sub AddV(ByVal x As Double)
    nV = nV + 1
    If nV > UBound(vX) Then
        ReDim Preserve vX(1 To nV * 2): ReDim Preserve vY(1 To nV * 2): ReDim Preserve vB(1 To nV * 2)
    End If
    vX(nV) = x: vY(nV) = 0: vB(nV) = 0
End Sub

Private Function OnLay(ByVal lay As String) As Boolean
    OnLay = (StrComp(eLay, lay, vbTextCompare) = 0)
End Function

Private Sub EndEntity(ByVal mode As Long, ByVal lay As String)
    Dim i As Long, L As Double
    If ePaper Then Exit Sub
    If eTyp = "TEXT" Or eTyp = "MTEXT" Or eTyp = "ATTRIB" Then
        AddStationText
        Exit Sub
    End If
    If eTyp = "SPLINE" Or eTyp = "POLYLINE" Or eTyp = "ELLIPSE" Then
        If mode = 2 Then
            If OnLay(lay) Then nBad = nBad + 1: badTyp = eTyp
        End If
        Exit Sub
    End If
    If Not Normalize() Then Exit Sub
    Select Case mode
        Case 1
            For i = 1 To nV - 1
                L = L + SegLen(vX(i), vY(i), vX(i + 1), vY(i + 1), vB(i))
            Next i
            layN(eLay) = layN(eLay) + 1
            If Not eClosed And L > layLen(eLay) Then layLen(eLay) = L
        Case 2
            If OnLay(lay) Then
                If eClosed Then
                    nBad = nBad + 1: badTyp = "closed LWPOLYLINE"
                Else
                    AddPiece
                End If
            End If
        Case 3
            If Not OnLay(lay) Then Crossings
    End Select
End Sub

' the entity as a run of points 1..nV with bulges (vB(i) = segment i to i + 1), world X / Y
Private Function Normalize() As Boolean
    Dim i As Long, sw As Double, cx0 As Double, cy0 As Double
    Select Case eTyp
        Case "LINE"
            If nV < 1 Then Exit Function
            If nV < 2 Then AddV 0
            nV = 2: vX(2) = eX2: vY(2) = eY2: vB(1) = 0: vB(2) = 0
        Case "ARC"
            If nV < 1 Or eR <= 0 Then Exit Function
            cx0 = vX(1): cy0 = vY(1)
            sw = eA2 - eA1
            Do While sw <= 0: sw = sw + 360: Loop
            Do While sw > 360: sw = sw - 360: Loop
            If nV < 2 Then AddV 0
            nV = 2
            vX(1) = cx0 + eR * Cos(eA1 * PI / 180): vY(1) = cy0 + eR * Sin(eA1 * PI / 180)
            vX(2) = cx0 + eR * Cos(eA2 * PI / 180): vY(2) = cy0 + eR * Sin(eA2 * PI / 180)
            vB(1) = Tan(sw * PI / 180 / 4): vB(2) = 0
            If eZ < 0 Then vX(1) = -vX(1): vX(2) = -vX(2): vB(1) = -vB(1)
        Case "LWPOLYLINE"
            If nV < 2 Then Exit Function
            If eZ < 0 Then
                For i = 1 To nV
                    vX(i) = -vX(i): vB(i) = -vB(i)
                Next i
            End If
            If (eFlag And 1) = 1 Then
                eClosed = True
                AddV vX(1)
                vY(nV) = vY(1)
            Else
                vB(nV) = 0
            End If
    End Select
    Normalize = True
End Function

' ---- layers -----------------------------------------------------------------
Private Sub LayerStats()
    Set layN = CreateObject("Scripting.Dictionary")
    Set layLen = CreateObject("Scripting.Dictionary")
    layN.CompareMode = vbTextCompare: layLen.CompareMode = vbTextCompare
    Scan 1, ""
End Sub

Private Function IsPlanCand(ByVal s As String) As Boolean
    Dim u As String: u = UCase$(s)
    IsPlanCand = (InStr(u, "PRF") = 0 And InStr(u, "PAF") = 0 And InStr(u, "CER") = 0 And InStr(u, "ANTET") = 0)
End Function

Private Function BestLayer() As String
    Dim k As Variant, b As Double, p As Long
    For p = 1 To 2
        For Each k In layLen.Keys
            If (p = 2 Or (IsPlanCand(CStr(k)) And layN(k) = 1)) And layLen(k) > b Then b = layLen(k): BestLayer = CStr(k)
        Next k
        If BestLayer <> "" Then Exit Function
    Next p
End Function

Private Function AskLayer() As String
    Dim nm() As String, ln() As Double, sc() As Double, n As Long, k As Variant, i As Long, j As Long, s As String, t As Double, u As String
    n = layN.Count
    If n = 0 Then Err.Raise vbObjectError + 1, , "No LINE, ARC or LWPOLYLINE in the ENTITIES section of the drawing."
    ReDim nm(1 To n): ReDim ln(1 To n): ReDim sc(1 To n)
    For Each k In layN.Keys
        i = i + 1: nm(i) = CStr(k): ln(i) = layLen(k)
        sc(i) = ln(i) + IIf(IsPlanCand(nm(i)), 1000000#, 0#) + IIf(layN(nm(i)) = 1, 100000#, 0#)
    Next k
    For i = 1 To IIf(n < 12, n, 12)
        For j = i + 1 To n
            If sc(j) > sc(i) Then
                t = sc(i): sc(i) = sc(j): sc(j) = t
                t = ln(i): ln(i) = ln(j): ln(j) = t
                u = nm(i): nm(i) = nm(j): nm(j) = u
            End If
        Next j
        s = s & "  " & Left$(nm(i), 40) & "  (" & layN(nm(i)) & ", longest " & Format(ln(i), "0") & " m)" & vbCrLf
    Next i
    AskLayer = InputBox("Layer of the bridge centerline (the axis line of the deck)." & vbCrLf & "Layers:" & vbCrLf & s, MSG_TITLE, BestLayer())
End Function

' ---- centerline ---------------------------------------------------------------
Private Sub AddPiece()
    Dim i As Long
    nPc = nPc + 1
    If nPc > UBound(pcFrom) Then ReDim Preserve pcFrom(1 To nPc * 2): ReDim Preserve pcTo(1 To nPc * 2)
    If nQ + nV > UBound(qX) Then
        ReDim Preserve qX(1 To (nQ + nV) * 2): ReDim Preserve qY(1 To (nQ + nV) * 2): ReDim Preserve qB(1 To (nQ + nV) * 2)
    End If
    pcFrom(nPc) = nQ + 1
    For i = 1 To nV
        nQ = nQ + 1
        qX(nQ) = vX(i): qY(nQ) = vY(i): qB(nQ) = vB(i)
    Next i
    qB(nQ) = 0
    pcTo(nPc) = nQ
End Sub

Private Sub ReadPieces(ByVal lay As String)
    nPc = 0: nQ = 0: nBad = 0: badTyp = ""
    ReDim pcFrom(1 To 16): ReDim pcTo(1 To 16)
    ReDim qX(1 To 256): ReDim qY(1 To 256): ReDim qB(1 To 256)
    Scan 2, lay
    If nBad > 0 Then Err.Raise vbObjectError + 2, , "Layer " & lay & " holds " & nBad & " " & badTyp & " (not read). Convert the centerline to an open polyline (PEDIT / EXPLODE) and try again."
    If nPc = 0 Then Err.Raise vbObjectError + 3, , "No LINE, ARC or LWPOLYLINE on layer " & lay & " (model space)."
End Sub

Private Function Near(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Boolean
    Near = ((x1 - x2) ^ 2 + (y1 - y2) ^ 2 <= TOL_JOIN ^ 2)
End Function

' piece k as points 1..m with bulges, reversed if rev
Private Sub GetPiece(ByVal k As Long, ByVal rev As Boolean, px() As Double, py() As Double, pb() As Double, ByRef m As Long)
    Dim j As Long, a As Long
    a = pcFrom(k): m = pcTo(k) - a + 1
    ReDim px(1 To m): ReDim py(1 To m): ReDim pb(1 To m)
    For j = 1 To m
        If rev Then
            px(j) = qX(pcTo(k) - j + 1): py(j) = qY(pcTo(k) - j + 1)
            If j < m Then pb(j) = -qB(a + m - j - 1)
        Else
            px(j) = qX(a + j - 1): py(j) = qY(a + j - 1)
            If j < m Then pb(j) = qB(a + j - 1)
        End If
    Next j
    pb(m) = 0
End Sub

Private Sub AttachPiece(ByVal k As Long, ByVal rev As Boolean, ByVal atHead As Boolean)
    Dim px() As Double, py() As Double, pb() As Double, m As Long, j As Long, i As Long
    GetPiece k, rev, px, py, pb, m
    If nCh = 0 Then
        For j = 1 To m
            chX(j) = px(j): chY(j) = py(j): chB(j) = pb(j)
        Next j
        nCh = m
    ElseIf Not atHead Then
        chB(nCh) = pb(1)
        For j = 2 To m
            nCh = nCh + 1
            chX(nCh) = px(j): chY(nCh) = py(j): chB(nCh) = pb(j)
        Next j
    Else
        For i = nCh To 1 Step -1
            chX(i + m - 1) = chX(i): chY(i + m - 1) = chY(i): chB(i + m - 1) = chB(i)
        Next i
        For j = 1 To m - 1
            chX(j) = px(j): chY(j) = py(j): chB(j) = pb(j)
        Next j
        nCh = nCh + m - 1
    End If
End Sub

Private Sub BuildChain(ByVal lay As String)
    Dim used() As Boolean, k As Long, nLeft As Long, found As Boolean, i As Long
    Dim xs As Double, ys As Double, xe As Double, ye As Double, ox As Double, oy As Double, r As Double, th As Double, sg As Double
    ReDim used(1 To nPc)
    ReDim chX(1 To nQ): ReDim chY(1 To nQ): ReDim chB(1 To nQ)
    nCh = 0
    AttachPiece 1, False, False
    used(1) = True: nLeft = nPc - 1
    Do While nLeft > 0
        found = False
        For k = 1 To nPc
            If Not used(k) Then
                xs = qX(pcFrom(k)): ys = qY(pcFrom(k)): xe = qX(pcTo(k)): ye = qY(pcTo(k))
                If Near(xs, ys, chX(nCh), chY(nCh)) Then
                    AttachPiece k, False, False: found = True
                ElseIf Near(xe, ye, chX(nCh), chY(nCh)) Then
                    AttachPiece k, True, False: found = True
                ElseIf Near(xe, ye, chX(1), chY(1)) Then
                    AttachPiece k, False, True: found = True
                ElseIf Near(xs, ys, chX(1), chY(1)) Then
                    AttachPiece k, True, True: found = True
                End If
                If found Then
                    used(k) = True: nLeft = nLeft - 1
                    Exit For
                End If
            End If
        Next k
        If Not found Then Err.Raise vbObjectError + 4, , "The centerline on layer " & lay & " is in " & nPc & " pieces that do not join end to end (" & nLeft & " left over, gap over " & TOL_JOIN & " m)."
    Loop
    ' lengths, bounding boxes
    ReDim chS(1 To nCh)
    ReDim sgX0(1 To nCh): ReDim sgY0(1 To nCh): ReDim sgX1(1 To nCh): ReDim sgY1(1 To nCh)
    nArc = 0: chS(1) = 0
    bbX0 = chX(1): bbX1 = chX(1): bbY0 = chY(1): bbY1 = chY(1)
    For i = 1 To nCh - 1
        chS(i + 1) = chS(i) + SegLen(chX(i), chY(i), chX(i + 1), chY(i + 1), chB(i))
        sgX0(i) = IIf(chX(i) < chX(i + 1), chX(i), chX(i + 1)): sgX1(i) = IIf(chX(i) > chX(i + 1), chX(i), chX(i + 1))
        sgY0(i) = IIf(chY(i) < chY(i + 1), chY(i), chY(i + 1)): sgY1(i) = IIf(chY(i) > chY(i + 1), chY(i), chY(i + 1))
        If Abs(chB(i)) > 0.000000001 Then
            nArc = nArc + 1
            sg = Abs(chB(i)) * Sqr((chX(i + 1) - chX(i)) ^ 2 + (chY(i + 1) - chY(i)) ^ 2) / 2
            If Abs(chB(i)) > 1 Then ArcOf i, ox, oy, r, th: sg = 2 * r
            sgX0(i) = sgX0(i) - sg: sgX1(i) = sgX1(i) + sg: sgY0(i) = sgY0(i) - sg: sgY1(i) = sgY1(i) + sg
        End If
        If sgX0(i) < bbX0 Then bbX0 = sgX0(i)
        If sgX1(i) > bbX1 Then bbX1 = sgX1(i)
        If sgY0(i) < bbY0 Then bbY0 = sgY0(i)
        If sgY1(i) > bbY1 Then bbY1 = sgY1(i)
    Next i
    totLen = chS(nCh)
    If totLen <= 0 Then Err.Raise vbObjectError + 5, , "The centerline on layer " & lay & " has no length."
End Sub

Private Function SegLen(ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double, ByVal b As Double) As Double
    Dim c As Double, th As Double
    c = Sqr((x1 - x0) ^ 2 + (y1 - y0) ^ 2)
    If Abs(b) < 0.000000001 Then
        SegLen = c
    Else
        th = 4 * Atn(Abs(b))
        SegLen = c * th / (2 * Sin(th / 2))
    End If
End Function

' arc of chain segment i: centre, radius, signed sweep (+ = counter-clockwise)
Private Sub ArcOf(ByVal i As Long, ByRef ox As Double, ByRef oy As Double, ByRef r As Double, ByRef th As Double)
    Dim dx As Double, dy As Double, b As Double, f As Double
    dx = chX(i + 1) - chX(i): dy = chY(i + 1) - chY(i): b = chB(i)
    th = 4 * Atn(b)
    f = (1 - b * b) / (4 * b)
    ox = (chX(i) + chX(i + 1)) / 2 - dy * f
    oy = (chY(i) + chY(i + 1)) / 2 + dx * f
    r = Sqr((chX(i) - ox) ^ 2 + (chY(i) - oy) ^ 2)
End Sub

Private Sub PointAt(ByVal s As Double, ByRef x As Double, ByRef y As Double)
    Dim i As Long, t As Double, L As Double, ox As Double, oy As Double, r As Double, th As Double, a As Double
    If s < 0 Then s = 0
    If s > totLen Then s = totLen
    i = 1
    Do While i < nCh - 1 And chS(i + 1) < s: i = i + 1: Loop
    t = s - chS(i): L = chS(i + 1) - chS(i)
    If Abs(chB(i)) < 0.000000001 Then
        x = chX(i) + (chX(i + 1) - chX(i)) * t / L
        y = chY(i) + (chY(i + 1) - chY(i)) * t / L
    Else
        ArcOf i, ox, oy, r, th
        a = th * t / L
        x = ox + (chX(i) - ox) * Cos(a) - (chY(i) - oy) * Sin(a)
        y = oy + (chX(i) - ox) * Sin(a) + (chY(i) - oy) * Cos(a)
    End If
End Sub

' ---- crossings ------------------------------------------------------------------
Private Sub Crossings()
    Dim j As Long, i As Long, x0 As Double, y0 As Double, x1 As Double, y1 As Double
    For j = 1 To nV - 1
        If Abs(vB(j)) < 0.000000001 Then
            x0 = vX(j): y0 = vY(j): x1 = vX(j + 1): y1 = vY(j + 1)
            If Not ((x0 < bbX0 And x1 < bbX0) Or (x0 > bbX1 And x1 > bbX1) Or (y0 < bbY0 And y1 < bbY0) Or (y0 > bbY1 And y1 > bbY1)) Then
                For i = 1 To nCh - 1
                    If Not ((x0 < sgX0(i) And x1 < sgX0(i)) Or (x0 > sgX1(i) And x1 > sgX1(i)) Or (y0 < sgY0(i) And y1 < sgY0(i)) Or (y0 > sgY1(i) And y1 > sgY1(i))) Then
                        CrossSeg i, x0, y0, x1, y1
                    End If
                Next i
            End If
        End If
    Next j
End Sub

Private Sub CrossSeg(ByVal i As Long, ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double)
    Dim rx As Double, ry As Double, dx As Double, dy As Double, den As Double, t As Double, u As Double
    Dim ox As Double, oy As Double, r As Double, th As Double, fx As Double, fy As Double, a As Double, b As Double, c As Double
    Dim disc As Double, k As Long, px As Double, py As Double, ang As Double
    dx = x1 - x0: dy = y1 - y0
    If Abs(chB(i)) < 0.000000001 Then
        rx = chX(i + 1) - chX(i): ry = chY(i + 1) - chY(i)
        den = rx * dy - ry * dx
        If Abs(den) < 0.000000000001 Then Exit Sub
        t = ((x0 - chX(i)) * dy - (y0 - chY(i)) * dx) / den
        u = ((x0 - chX(i)) * ry - (y0 - chY(i)) * rx) / den
        If t >= 0 And t <= 1 And u >= 0 And u <= 1 Then AddCr chS(i) + t * (chS(i + 1) - chS(i))
    Else
        ArcOf i, ox, oy, r, th
        fx = x0 - ox: fy = y0 - oy
        a = dx * dx + dy * dy: b = 2 * (fx * dx + fy * dy): c = fx * fx + fy * fy - r * r
        If a <= 0 Then Exit Sub
        disc = b * b - 4 * a * c
        If disc < 0 Then Exit Sub
        For k = -1 To 1 Step 2
            u = (-b + k * Sqr(disc)) / (2 * a)
            If u >= 0 And u <= 1 Then
                px = x0 + u * dx - ox: py = y0 + u * dy - oy
                ang = Atn2((chX(i) - ox) * py - (chY(i) - oy) * px, (chX(i) - ox) * px + (chY(i) - oy) * py)
                If th < 0 Then ang = -ang
                If ang < 0 Then ang = ang + 2 * PI
                If ang <= Abs(th) + 0.000000001 Then AddCr chS(i) + r * ang
            End If
            If disc = 0 Then Exit For
        Next k
    End If
End Sub

Private Sub AddCr(ByVal s As Double)
    nCr = nCr + 1
    If nCr > UBound(crS) Then ReDim Preserve crS(1 To nCr * 2)
    crS(nCr) = s
End Sub

Private Sub SortCr(ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Double, t As Double
    Do While lo < hi
        i = lo: j = hi: p = crS((lo + hi) \ 2)
        Do While i <= j
            Do While crS(i) < p: i = i + 1: Loop
            Do While crS(j) > p: j = j - 1: Loop
            If i <= j Then
                t = crS(i): crS(i) = crS(j): crS(j) = t
                i = i + 1: j = j - 1
            End If
        Loop
        If j - lo < hi - i Then
            If lo < j Then SortCr lo, j
            lo = i
        Else
            If i < hi Then SortCr i, hi
            hi = j
        End If
    Loop
End Sub

Private Function NearestCr(ByVal s As Double) As Double
    NearestCr = Abs(NearestCrS(s) - s)
End Function

' station of the crossing nearest to s
Private Function NearestCrS(ByVal s As Double) As Double
    Dim lo As Long, hi As Long, m As Long
    lo = 1: hi = nCr
    Do While hi - lo > 1
        m = (lo + hi) \ 2
        If crS(m) < s Then lo = m Else hi = m
    Loop
    NearestCrS = crS(lo)
    If Abs(crS(hi) - s) < Abs(crS(lo) - s) Then NearestCrS = crS(hi)
End Function

' runs of crossings matching the span list; returns how many distinct runs (mtS / mtD / mtW, best first in a1S / a1Dir)
Private Function FindA1(ByRef worst As Double) As Long
    Dim i As Long, k As Long, dirn As Long, s0 As Double, e As Double, w As Double, d As Double, j As Long, hit As Long
    nMt = 0: ReDim mtS(1 To 8): ReDim mtD(1 To 8): ReDim mtW(1 To 8)
    worst = 1E+30
    If nCr < nSp + 1 Then Exit Function
    For dirn = 1 To -1 Step -2
        For i = 1 To nCr
            s0 = crS(i): e = s0 + dirn * cum(nSp)
            If e >= -TOL_SPAN And e <= totLen + TOL_SPAN Then
                w = 0
                For k = 1 To nSp
                    d = NearestCr(s0 + dirn * cum(k))
                    If d > w Then w = d
                    If w > TOL_SPAN Then Exit For
                Next k
                If w <= TOL_SPAN Then
                    hit = 0
                    For j = 1 To nMt
                        If mtD(j) = dirn And Abs(mtS(j) - s0) < 0.5 Then hit = j
                    Next j
                    If hit = 0 And nMt < 8 Then nMt = nMt + 1: hit = nMt: mtW(hit) = 1E+30
                    If hit > 0 Then
                        If w < mtW(hit) Then mtW(hit) = w: mtS(hit) = s0: mtD(hit) = dirn
                    End If
                    If w < worst Then worst = w: a1S = s0: a1Dir = dirn
                End If
            End If
        Next i
    Next dirn
    FindA1 = nMt
End Function

' station texts (101+848.373) near the line, at the distance along the line of the nearest sample point
Private Sub AddStationText()
    Dim m As Object, k As Long, b As Long, d As Double, dB As Double, isNm As Boolean
    If nV < 1 Then Exit Sub
    isNm = reNm.Test(Trim$(eTxt))
    If Not isNm And Not reSt.Test(eTxt) Then Exit Sub
    If vX(1) < bbX0 - TXT_NEAR Or vX(1) > bbX1 + TXT_NEAR Or vY(1) < bbY0 - TXT_NEAR Or vY(1) > bbY1 + TXT_NEAR Then Exit Sub
    dB = 1E+30
    For k = 1 To nSm
        d = (smX(k) - vX(1)) ^ 2 + (smY(k) - vY(1)) ^ 2
        If d < dB Then dB = d: b = k
    Next k
    If dB > TXT_NEAR ^ 2 Then Exit Sub
    If isNm Then
        nNm = nNm + 1
        If nNm > UBound(nameS) Then ReDim Preserve nameS(1 To nNm * 2): ReDim Preserve nameT(1 To nNm * 2)
        nameS(nNm) = smS(b): nameT(nNm) = UCase$(Trim$(eTxt))
        Exit Sub
    End If
    Set m = reSt.Execute(eTxt)(0)
    nTx = nTx + 1
    If nTx > UBound(txS) Then ReDim Preserve txS(1 To nTx * 2): ReDim Preserve txKm(1 To nTx * 2)
    txS(nTx) = smS(b): txKm(nTx) = Val(m.SubMatches(0)) * 1000# + Val(m.SubMatches(1))
End Sub

' +1 = km grow along the line, -1 = against, 0 = not known (neighbouring texts compared)
Private Function KmDirection(ByRef votes As String) As Long
    Dim i As Long, j As Long, t As Double, f As Long, r As Long, ds As Double, dk As Double
    For i = 1 To nTx - 1
        For j = i + 1 To nTx
            If txS(j) < txS(i) Then
                t = txS(i): txS(i) = txS(j): txS(j) = t
                t = txKm(i): txKm(i) = txKm(j): txKm(j) = t
            End If
        Next j
    Next i
    For i = 1 To nTx - 1
        ds = txS(i + 1) - txS(i): dk = txKm(i + 1) - txKm(i)
        If ds > 1 Then
            If Abs(dk - ds) < 2 Then f = f + 1
            If Abs(dk + ds) < 2 Then r = r + 1
        End If
    Next i
    votes = nTx & " station texts near the line, " & f & " neighbour pairs along it, " & r & " against"
    If f >= 2 And f >= 2 * r Then KmDirection = 1
    If r >= 2 And r >= 2 * f Then KmDirection = -1
End Function

' sample points every SAMPLE_DS along the line (for the station texts)
Private Sub Samples()
    Dim k As Long, n As Long
    n = Int(totLen / SAMPLE_DS) + 1
    nSm = n + 1
    ReDim smS(1 To nSm): ReDim smX(1 To nSm): ReDim smY(1 To nSm)
    For k = 1 To nSm
        smS(k) = IIf(k = nSm, totLen, (k - 1) * SAMPLE_DS)
        If smS(k) > totLen Then smS(k) = totLen
        PointAt smS(k), smX(k), smY(k)
    Next k
End Sub

' ---- process --------------------------------------------------------------------
Private Sub ReadSpans()
    Dim v As Variant, r As Range, i As Long, k As Long
    nSp = 0
    ReDim sp(1 To 120)
    For Each v In Array("G_ACIKLIK_LISTE_1", "G_ACIKLIK_LISTE_2")
        Set r = NamedRange(CStr(v))
        For i = 1 To r.Cells.Count
            If IsEmpty(r.Cells(i).Value) Or Not IsNumeric(r.Cells(i).Value) Then GoTo Done
            nSp = nSp + 1: sp(nSp) = CDbl(r.Cells(i).Value)
        Next i
    Next v
Done:
    If nSp = 0 Then Err.Raise vbObjectError + 6, , "The span list (G_ACIKLIK_LISTE_1) is empty."
    ReDim cum(0 To nSp)
    For k = 1 To nSp
        cum(k) = cum(k - 1) + sp(k)
    Next k
End Sub

Private Sub SetA1(ByVal d As Double)
    If d >= 0 Then
        a1S = d: a1Dir = 1
    Else
        a1S = totLen + d: a1Dir = -1
    End If
End Sub

Private Function Process(ByVal lay As String, ByVal a1Dist As String, ByVal ask As Boolean) As String
    Dim t As Range, w As Double, nM As Long, why As String, v As String, k As Long, e As Double
    Dim cl As Double, off As Double, mx As Double, s As String, drawn As Boolean, dv As Double, wS As Double
    Set t = NamedRange("G_AKS_XY")
    ReadSpans
    If t.Rows.Count < nSp + 1 Then Err.Raise vbObjectError + 7, , "G_AKS_XY has " & t.Rows.Count & " rows, the bridge has " & nSp + 1 & " axes."
    ReadPieces lay
    BuildChain lay
    a1How = "": nNm = 0
    If Trim$(a1Dist) <> "" Then
        SetA1 Val(Replace(a1Dist, ",", "."))
        a1How = "given"
    Else
        Samples
        Set reSt = CreateObject("VBScript.RegExp")
        reSt.Pattern = "(?:^|[^0-9])(\d{1,3})\+(\d{3}\.\d{3})(?:$|[^0-9])": reSt.IgnoreCase = False: reSt.Global = False
        Set reNm = CreateObject("VBScript.RegExp")
        reNm.Pattern = "^(A[12]|P[0-9]+)$": reNm.IgnoreCase = True: reNm.Global = False
        ReDim nameS(1 To 128): ReDim nameT(1 To 128)
        nCr = 0: ReDim crS(1 To 1024)
        nTx = 0: ReDim txS(1 To 128): ReDim txKm(1 To 128)
        Scan 3, lay
        If nCr > 1 Then SortCr 1, nCr
        nM = FindA1(w)
        Dim kmD As Long, votes As String, pick As Long
        If nM > 1 Then
            kmD = KmDirection(votes)
            If kmD <> 0 Then
                For k = 1 To nM
                    If mtD(k) = kmD Then
                        If pick = 0 Then pick = k Else pick = -1
                    End If
                Next k
                If pick > 0 Then
                    a1S = mtS(pick): a1Dir = mtD(pick): w = mtW(pick): nM = 1
                End If
            End If
        End If
        If nM = 1 Then
            drawn = True
            a1How = "from the lines crossing the centerline (" & nSp + 1 & " crossings match the span list, worst " & Format(w, "0.000") & " m" & IIf(pick > 0, "; direction from " & votes, "") & ")"
        Else
            If nM = 0 Then
                why = "no " & nSp + 1 & " lines cross the centerline at the span spacing"
            Else
                why = nM & " different runs of crossing lines match the span list"
            End If
            If Not ask Then Err.Raise vbObjectError + 8, , "A1 not found: " & why & ". Give the distance from the first point of the line to A1 (a1Dist)."
            v = InputBox("A1 could not be found: " & why & "." & vbCrLf & vbCrLf & "Centerline: " & Format(totLen, "0.000") & " m, first point " & Format(chX(1), "0.000") & " / " & Format(chY(1), "0.000") & "." & vbCrLf & _
                "Distance along the centerline from its first point to A1 (m)." & vbCrLf & "Negative = from its last point, against the drawing direction.", MSG_TITLE)
            If Trim$(v) = "" Then Exit Function
            SetA1 Val(Replace(v, ",", "."))
            a1How = "given"
        End If
    End If
    e = a1S + a1Dir * cum(nSp)
    If a1S < -0.001 Or a1S > totLen + 0.001 Or e < -0.001 Or e > totLen + 0.001 Then Err.Raise vbObjectError + 9, , _
        "The bridge (" & Format(cum(nSp), "0.000") & " m) runs past the end of the centerline (" & Format(totLen, "0.000") & " m, A1 at " & Format(a1S, "0.000") & " m " & IIf(a1Dir = 1, "along", "against") & " the line)."
    ReDim axX(0 To nSp): ReDim axY(0 To nSp): ReDim axS(0 To nSp)
    For k = 0 To nSp
        axS(k) = a1S + a1Dir * cum(k)
        If drawn Then axS(k) = NearestCrS(axS(k))
        PointAt axS(k), axX(k), axY(k)
        If k > 0 Then
            dv = Abs(Abs(axS(k) - axS(k - 1)) - sp(k))
            If dv > wS Then wS = dv
        End If
    Next k
    ReDim axR(1 To nSp)
    Dim nCv As Long
    For k = 1 To nSp
        axR(k) = SpanRadius(k)
        If axR(k) <> 0 Then nCv = nCv + 1
    Next k
    cl = Sqr((axX(nSp) - axX(0)) ^ 2 + (axY(nSp) - axY(0)) ^ 2)
    For k = 1 To nSp - 1
        off = Abs((axX(nSp) - axX(0)) * (axY(k) - axY(0)) - (axY(nSp) - axY(0)) * (axX(k) - axX(0))) / cl
        If off > mx Then mx = off
    Next k
    s = SCRIPT_ID & " " & SCRIPT_VERSION & vbCrLf
    s = s & "Centerline: " & nPc & " piece(s), " & nCh - 1 & " segments (" & nArc & " arcs), " & Format(totLen, "0.000") & " m" & vbCrLf
    s = s & "A1 " & a1How & ": " & Format(a1S, "0.000") & " m from the first point, axes " & IIf(a1Dir = 1, "along", "against") & " the drawing direction" & vbCrLf
    s = s & nSp + 1 & " axes " & IIf(drawn, "at the drawn crossing points, worst span against the list " & Format(wS, "0.000") & " m", "at the span lengths along the line") & ", spans total " & Format(cum(nSp), "0.000") & " m" & vbCrLf
    s = s & NameCheck()
    s = s & "A1 " & Format(axX(0), "0.000") & " / " & Format(axY(0), "0.000") & ", A2 " & Format(axX(nSp), "0.000") & " / " & Format(axY(nSp), "0.000") & vbCrLf
    s = s & "Chord A1-A2 " & Format(cl, "0.000") & " m, largest axis offset from it " & Format(mx, "0.000") & " m" & vbCrLf
    s = s & "Spans on a curve: " & nCv & " of " & nSp & vbCrLf
    Process = s
End Function

' radius of span k: least-squares parabola v = a u (L - u) of the centerline about the span chord, sampled every
' metre (the clothoids are drawn as chords, three points alone would be too noisy). R < 0 left, R > 0 right as in
' the drawing labels, 0 = straight
Private Function SpanRadius(ByVal k As Long) As Double
    Dim n As Long, j As Long, s0 As Double, s1 As Double, L As Double, cx As Double, cy As Double, x As Double, y As Double
    Dim u As Double, v As Double, w As Double, sn As Double, sd As Double, a As Double
    s0 = axS(k - 1): s1 = axS(k)
    L = Sqr((axX(k) - axX(k - 1)) ^ 2 + (axY(k) - axY(k - 1)) ^ 2)
    If L <= 0 Then Exit Function
    cx = (axX(k) - axX(k - 1)) / L: cy = (axY(k) - axY(k - 1)) / L
    n = Int(Abs(s1 - s0)): If n < 8 Then n = 8
    For j = 1 To n - 1
        PointAt s0 + (s1 - s0) * j / n, x, y
        u = (x - axX(k - 1)) * cx + (y - axY(k - 1)) * cy
        v = -(x - axX(k - 1)) * cy + (y - axY(k - 1)) * cx
        w = u * (L - u)
        sn = sn + v * w: sd = sd + w * w
    Next j
    If sd <= 0 Then Exit Function
    a = sn / sd
    If Abs(a) < 0.5 / R_STRAIGHT Then Exit Function
    SpanRadius = Round(1 / (2 * a), 1)
End Function

' axis names near the line against the axis order: A1, P1 .. P(n-1), A2
Private Function NameCheck() As String
    Dim i As Long, k As Long, b As Long, dB As Double, d As Double, w As String, nOk As Long, nBadN As Long, bad As String
    If nNm = 0 Then Exit Function
    For i = 1 To nNm
        dB = 1E+30
        For k = 0 To nSp
            d = Abs(nameS(i) - axS(k))
            If d < dB Then dB = d: b = k
        Next k
        w = IIf(b = 0, "A1", IIf(b = nSp, "A2", "P" & b))
        If nameT(i) = w Then
            nOk = nOk + 1
        Else
            nBadN = nBadN + 1
            If nBadN <= 5 Then bad = bad & " " & nameT(i) & " at axis " & w
        End If
    Next i
    If nBadN = 0 And nOk = nSp + 1 Then
        NameCheck = "Axis names: all " & nOk & " in order" & vbCrLf
    Else
        NameCheck = "WARNING axis names: " & nOk & " of " & nSp + 1 & " in order, " & nBadN & " wrong" & IIf(bad = "", "", ":" & bad) & vbCrLf
    End If
End Function

' ---- table ------------------------------------------------------------------
Private Function NamedRange(ByVal nm As String) As Range
    Dim ws As Worksheet
    On Error Resume Next
    Set NamedRange = ThisWorkbook.Names(nm).RefersToRange
    On Error GoTo 0
    If NamedRange Is Nothing And nm = "G_AKS_XY" Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets("INPUT")
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Cells(223, 1).Value = "Aks X": ws.Cells(223, 2).Value = "Aks Y"
            ThisWorkbook.Names.Add Name:="G_AKS_XY", RefersTo:=ws.Range("A224:B264")
            Set NamedRange = ws.Range("A224:B264")
        End If
    End If
    If NamedRange Is Nothing Then Err.Raise vbObjectError + 20, , "INPUT name " & nm & " is missing in this workbook."
End Function

Private Sub WriteTable()
    Dim t As Range, k As Long
    Set t = NamedRange("G_AKS_XY")
    t.ClearContents
    t.Columns(1).Offset(0, 2).ClearContents
    If t.Row > 1 Then t.Cells(0, 3).Value = "A" & ChrW(231) & ChrW(305) & "kl" & ChrW(305) & "k R (m)"
    For k = 0 To nSp
        t.Cells(k + 1, 1).Value = axX(k)
        t.Cells(k + 1, 2).Value = axY(k)
        If k > 0 Then
            If axR(k) <> 0 Then t.Cells(k, 3).Value = axR(k)
        End If
    Next k
End Sub

Private Function Atn2(ByVal yy As Double, ByVal xx As Double) As Double
    If Abs(xx) < 0.000000000001 Then
        Atn2 = IIf(yy >= 0, PI / 2, -PI / 2)
    ElseIf xx > 0 Then
        Atn2 = Atn(yy / xx)
    Else
        Atn2 = Atn(yy / xx) + IIf(yy >= 0, PI, -PI)
    End If
End Function
