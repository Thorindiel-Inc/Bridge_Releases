Attribute VB_Name = "MIDAS_KENARAYAK"
Option Explicit
' ============================================================================
'  MIDAS_KENARAYAK - builds the abutment model of the abutment workbook (8'6_..._KA1 / _KA2) in
'  MIDAS Civil NX through the MAPI REST API. Templates: models/1AS_KA1_K.mct (piled) and
'  models/1AS_KA2_Y.mct (spread footing). Plan and owner decisions M1-M12:
'  docs/global-input/abutment-plan.md sections 4 and 4.1.
' ----------------------------------------------------------------------------
'  Inputs only by workbook names:
'    KA_*  local abutment block of INPUT (geometry, piles, soil, mesh, stiffness factors) and the
'          KA_ names on the 1_RAPOR load cells (seat loads KA_Y_*, pressures KA_P_*, kh KA_KH)
'    KA_KOMBINASYON = LC!A1:AN27, read by column position (D..Q = the 14 load cases in a fixed
'          order; its O1 header says PEQ-X but the column is PEQ-Y), R..AN = combinations
'    G_KIRIS_SAYISI / G_KIRIS_GABARI / G_HESAP_ACIKLIK (A1 = span type 1, A2 = last filled),
'    G_YATAY_GABARI (out-to-out width of the wing walls), G_DIGER_BETONLAR, G_BIRIM_AGIRLIK_BETONARME
'    KA_VEREVLIK (skew, degrees, + = support line counter-clockwise in plan); missing = square
'  Model frame (as the templates): x across the bridge (0 .. footing width), y along the bridge
'  axis (0 = back edge of the footing on the fill side, footing length = front edge on the span
'  side), z up, z = 0 = footing mid-plane. Skew (M8): parameter (u, v) -> X = u,
'  Y = v + (u - B/2) tan(theta): stem, backwall and seat run along the skew line, wing and ear
'  walls stay parallel to the bridge axis, the footing is a parallelogram, pile rows follow the
'  skew line. A1 and A2 use the same sign: the A2 frame is the A1 frame turned 180 degrees in plan,
'  which maps every line onto a parallel line. Dimensions along y (heel, toe, footing length, pile
'  row spacing and edge distance) and the stem thickness are measured NORMAL to the stem (owner M21);
'  dimensions across (footing width, wing walls, pile spacing across) along x. Ear walls only when
'  KA_KULAK_BOYU > 0 and KA_KULAK_KALINLIK > 0 (owner M23).
'  Plates: thick with drilling DOF (subtype 3, M10); thickness 1 footing, 2 wing walls, 3 backwall
'  (offset (perde - alin) / 2 toward the fill = flush with the stem back face), 4 stem, 5 ear walls
'  (offset KA_KULAK_KACIKLIK). Local z of every wall points toward the fill (stem, backwall) or
'  inward (wing and ear walls); footing +Z. Pressures are given in local z, negative = away
'  from the fill.
'  Piled (KA_TEMEL_TIPI = KAZIKLI): piles 14000 + n s + p (n piles, s = level from the top,
'  p = rows * column + row), weightless material (M6), lateral springs kh x D x L (M4) with the
'  full element length at the end nodes too (M5), tip kv x pi D^2 / 4, rigid link from the footing
'  node (master) to the pile head, section stiffness 0.5 Iyy / Izz in boundary group Sismik.
'  Spread (YUZEYSEL): ope/SSPS surface springs on the footing plates: COMP SUBGRADE kv with DIR 1
'  (DIR 0 / 6 = upside down, measured 2026-11-10) and LINEAR [kh, kh, 0]; nonlinear control
'  ACTL-M1 NUMINC 5 (compression-only springs).
'  Load cases (owner 2026-11-10, both foundation types): every 'Add' row of KA_KOMBINASYON is one
'  factored USER load case with the LC name (Servis-1 ... TKDY-Y, DEAD); inside it every component
'  load carries the load group of its component (DC, EV, EH0, EHA, LL, LS, EQ-X, EQ-Y, TU, WS, PEQ-X,
'  PEQ-Y, BR/ACC, CF on the seat loads and pressures; SPLIT_GROUPS = False merges them into one item
'  per node / plate). Self weight and its kh inertia: one body-force record per case in group DC
'  (FV [kh fPEQ-Y, kh fPEQ-X, -fDC]; a second record in one case makes the analysis fail, T6). Combinations (owner M22): only the 'Envelope' rows
'  (ENV_SER, ENV_STR, ENV_DEP over the factored cases, ENV_ALL over those), LCOM-GEN. Boundary change: cases
'  with TKDY in the name on EQ (cracked), the others on ST, THRSEV on EQ.
'  Seat node count (owner 2026-11-11): a live build writes the mesh count into KA_MESNET_NOKTA
'  (1_RAPOR AH190, the divisor of the report's per-node seat values) unless that cell holds a
'  formula; a dry run only warns when they differ.
'  Loads (M1, M2, M7): seat value x girders / seat nodes on the stem-top nodes between the wing
'  walls' inner faces; fill inertia only in PEQ-X; transverse soil pressure on the +x wing wall.
'  Entries: MIDAS_KENARAYAK_OLUSTUR (macro window; asks the transverse direction on a skewed abutment),
'  MIDAS_KENARAYAK_PAYLOAD outDir [, transverseSign] (dry run, no HTTP), MIDAS_KENARAYAK_BUILD reportFile,
'  apiKey [, saveModel [, transverseSign]] (scripted, no dialogs); transverseSign +1 = +X (template), -1 = -X.
'  MAPI key: read by MIDAS_CORE (name G_MAPI_KEY, then INPUT!T3, then MIDAS_MAPI_KEY, then asked; memory only).
' ============================================================================

Private Const SCRIPT_ID As String = "midas-abutment-build"
Private Const SCRIPT_VERSION As String = "2026-11-12a"
Private Const SCRIPT_CHANGELOG As String = "Shared helpers (key, HTTP, JSON, files) are in the module MIDAS_CORE, which UPDATE adds; MAPI key order: G_MAPI_KEY, INPUT!T3, MIDAS_MAPI_KEY, prompt; A build writes the stem-top seat node count of the mesh into KA_MESNET_NOKTA (1_RAPOR AH190); New: abutment model in Civil NX (plates, piles or spread-footing springs, loads, combinations, skew) from the KA_ inputs of the abutment workbook"
Private Const SCRIPT_NEEDS As String = "midas-core"   ' UPDATE adds this module when the workbook lacks it

Private Const MSG_TITLE As String = "MIDAS_KENARAYAK"
Private Const PILE0 As Long = 14000
Private Const GRAV_STYP As Double = 9.806
Private Const CHUNK As Long = 300
Private Const PROG_BAR_WIDTH As Long = 24
Private Const EPS As Double = 0.000001
Private Const PI_ As Double = 3.14159265358979
Private Const M_FOOT As Long = 1
Private Const M_STEM As Long = 2
Private Const M_BACK As Long = 3
Private Const M_WINGL As Long = 4
Private Const M_WINGR As Long = 5
Private Const M_EARL As Long = 6
Private Const M_EARR As Long = 7
Private Const N_CASE As Long = 14
Private Const SPLIT_GROUPS As Boolean = True    ' one load group per component inside each factored case
Private Const GRP_BR As String = "BR/ACC"         ' load group of the braking component ("BR-ACC" if Civil NX refuses "/")

' ---- inputs ----
Private abutNm As String, piled As Boolean, thetaDeg As Double, tanT As Double, cosT As Double
Private fB As Double, fL As Double, tFt As Double, tSt As Double, tBw As Double, tWg As Double, tEr As Double
Private hStem As Double, hBack As Double, hTot As Double, heelW As Double, toeW As Double, wOut As Double
Private lEar As Double, hTip As Double, hRoot As Double, offEar As Double
Private pDia As Double, nPx As Long, nPy As Long, sPx As Double, sPy As Double, cPx As Double, cPy As Double, lPile As Double, dzPile As Double
Private nLay As Long, layD() As Double, layK() As Double, kvTip As Double, kvFoot As Double, khFoot As Double
Private meshM As Double, fPlate As Double, fPile As Double
Private eConc As Double, denConc As Double, clsConc As String
Private nGird As Long, spanCol As Long, nSeatWant As Long
Private qDC As Double, qLL As Double, qTUX As Double, qTUY As Double, qWSX As Double, qWSY As Double, qCF As Double, qBR As Double, qPX As Double, qPY As Double
Private pEH0 As Double, pEHA As Double, pEV As Double, pLS As Double, pEQX As Double, pEQY As Double, pEQs As Double, pPEQX As Double, pPEQY As Double, khSeis As Double
Private projName As String, clientName As String, bridgeName As String
Private warnText As String, askUser As Boolean, trSign As Long   ' trSign: +1 transverse loads toward +x, -1 toward -x
Private caseNm(1 To N_CASE) As String, caseTy(1 To N_CASE) As String
Private nCmb As Long, cmbName() As String, cmbKind() As Long, cmbItems() As String, cmbTier() As Long
Private cmbAct() As Boolean, cmbEff() As Boolean   ' column B of KA_KOMBINASYON (Inactive = skipped); row written in the model
Private pssfMem(1 To 7) As Long                    ' plate stiffness factor records per member, for the read-back
Private Const PILE_COLOR As String = "0,72,72,0,192,192,0,72,72"   ' pile section colour W / HF / HE (template SECT-COLOR)
Private nFac As Long, facIdx() As Long     ' 'Add' rows of KA_KOMBINASYON = factored load cases

' ---- geometry ----
Private e0 As Double, xW1 As Double, xW2 As Double, fi1 As Double, fi2 As Double
Private vBk As Double, vFr As Double, yS As Double, tStV As Double, toeUsed As Double
Private zFt As Double, zSeat As Double, zTop As Double
Private uF() As Double, nUF As Long, vFt() As Double, nVF As Long, zW() As Double, nZW As Long, vE() As Double, nVE As Long
Private earOn As Boolean, fLv As Double   ' ear walls modelled (M23); footing length along the bridge axis (M21)
Private nGrupW As Long, nThkW As Long, nOffW As Long, nPssfW As Long, nBodfW As Long, nCnldN As Long, nCnldI As Long, nPresE As Long, nPresI As Long, nSeisW As Long   ' written, for the read-back
Private apF() As Double, wPlate As Double   ' applied load totals per factored case (FX, FY, FZ); plate weight

' ---- model ----
Private nodeKey As Object
Private nId() As Long, nX() As Double, nY() As Double, nZ() As Double, nCnt As Long
Private eId() As Long, eKind() As Long, eMem() As Long, eThk() As Long, eNds() As Long, eNn() As Long
Private eCu() As Double, eCv() As Double, eCz() As Double, eRjt() As Boolean, ePrj() As Boolean, ePssf() As Boolean, eCnt As Long
Private nPlate As Long, nP As Long, nSeg As Long, segL As Double
Private pileTop() As Long, pileMaster() As Long
Private seatN() As Long, nSeat As Long
Private springJson As String, nSpring As Long

Private payloads As Collection
Private progStep As Long, progTotal As Long

' ============================================================================
'  entries
' ============================================================================
Public Sub MIDAS_KENARAYAK_OLUSTUR()
    trSign = 1
    RunBuild "", False
End Sub

' transverseSign: +1 = transverse loads toward +x (template), -1 = toward -x
Public Sub MIDAS_KENARAYAK_PAYLOAD(ByVal outDir As String, Optional ByVal transverseSign As Long = 1)
    trSign = IIf(transverseSign < 0, -1, 1)
    RunBuild outDir, True
End Sub

' scripted build, no dialogs (tests)
Public Sub MIDAS_KENARAYAK_BUILD(ByVal reportFile As String, ByVal apiKey As String, Optional ByVal saveModel As Boolean = True, Optional ByVal transverseSign As Long = 1)
    Dim rep As String
    On Error GoTo Fail
    askUser = False
    trSign = IIf(transverseSign < 0, -1, 1)
    PrepareModel
    McSetKey apiKey
    rep = SendAll(saveModel)
    WriteText reportFile, rep
    Application.StatusBar = False
    Exit Sub
Fail:
    WriteText reportFile, "FAIL: " & Err.Description
    Application.StatusBar = False
End Sub

Private Sub PrepareModel()
    warnText = ""
    Application.StatusBar = "MIDAS_KENARAYAK: reading inputs..."
    ReadInputs
    If trSign = 0 Then trSign = 1
    ' skewed abutment (owner 2026-11-10): the window asks which way the transverse loads act
    If askUser And Abs(thetaDeg) > EPS Then
        Select Case UiMsg("The abutment is skewed (" & JN(thetaDeg) & " deg)." & vbCrLf & vbCrLf & _
            "Which way do the TRANSVERSE loads act (seat wind and seismic loads across the bridge, the wing wall that takes the transverse seismic soil pressure)?" & vbCrLf & vbCrLf & _
            "YES = +X (as the template)" & vbCrLf & "NO = -X", vbYesNoCancel + vbQuestion, MSG_TITLE)
            Case vbYes: trSign = 1
            Case vbNo: trSign = -1
            Case Else: Err.Raise vbObjectError + 60, , "Cancelled."
        End Select
    End If
    BuildGeometry
    BuildPayloads
End Sub

Private Sub RunBuild(ByVal outDir As String, ByVal dry As Boolean)
    Dim msg As String, rep As String
    On Error GoTo Fail
    askUser = Not dry
    PrepareModel
    If dry Then
        SeatWarnDry
        WritePayloads outDir
        Application.StatusBar = False
        Exit Sub
    End If
    msg = Summary() & vbCrLf & vbCrLf & _
          "YES = WIPE the open Civil NX model, build the abutment " & abutNm & " model and SAVE it." & vbCrLf & _
          "NO = only write the JSON payloads (MIDAS\payload_kenarayak)."
    Select Case UiMsg(msg, vbYesNoCancel + vbQuestion, MSG_TITLE)
        Case vbNo
            SeatWarnDry
            WritePayloads ThisWorkbook.Path & "\MIDAS\payload_kenarayak"
            UiMsg "Payloads written to " & ThisWorkbook.Path & "\MIDAS\payload_kenarayak", vbInformation, MSG_TITLE
            Application.StatusBar = False
            Exit Sub
        Case vbCancel
            Application.StatusBar = False
            Exit Sub
    End Select
    rep = SendAll()
    Application.StatusBar = False
    UiMsg rep, IIf(InStr(rep, "FAIL") > 0, vbExclamation, vbInformation), MSG_TITLE
    Exit Sub
Fail:
    Dim d As String: d = Err.Description
    On Error Resume Next
    Application.StatusBar = False
    If dry Then
        MakeDir outDir
        WriteText outDir & "\ERROR.txt", d
    Else
        UiMsg "Stopped: " & d, vbCritical, MSG_TITLE
    End If
End Sub

' ============================================================================
'  inputs
' ============================================================================
Private Sub ReadInputs()
    Dim r As Range, i As Long, v As Variant, t As String, gi As Long
    abutNm = UCase$(Trim$(NS("KA_SECIM")))
    If abutNm <> "A1" And abutNm <> "A2" Then Err.Raise vbObjectError + 30, , "KA_SECIM must be A1 or A2 (found '" & abutNm & "')."
    t = Fold(NS("KA_TEMEL_TIPI"))
    If t = "KAZIKLI" Then
        piled = True
    ElseIf t = "YUZEYSEL" Then
        piled = False
    Else
        Err.Raise vbObjectError + 31, , "KA_TEMEL_TIPI must be KAZIKLI or YUZEYSEL (found '" & t & "')."
    End If
    fB = NV("KA_TEMEL_EN"): fL = NV("KA_TEMEL_BOY"): tFt = NV("KA_TEMEL_KALINLIK")
    tSt = NV("KA_PERDE_KALINLIK"): tBw = NV("KA_ALIN_KALINLIK"): tWg = NV("KA_RICAT_KALINLIK"): tEr = NV("KA_KULAK_KALINLIK")
    hStem = NV("KA_PERDE_YUKSEKLIK"): hBack = NV("KA_ALIN_YUKSEKLIK"): hTot = NV("KA_YUKSEKLIK")
    If Abs(hStem + hBack - hTot) > 0.005 Then Warn "KA_PERDE_YUKSEKLIK + KA_ALIN_YUKSEKLIK = " & JN(hStem + hBack) & " differs from KA_YUKSEKLIK = " & JN(hTot) & "; the model uses perde + alin."
    heelW = NV("KA_ARKA_GENISLIK"): toeW = NV("KA_ON_GENISLIK")
    lEar = NV("KA_KULAK_BOYU"): hTip = NV("KA_KULAK_UC_YUKSEKLIK"): hRoot = NV("KA_KULAK_KOK_YUKSEKLIK"): offEar = NV("KA_KULAK_KACIKLIK")
    meshM = NV("KA_AG_BOYU"): fPlate = NV("KA_RIJITLIK_PLAK")
    If meshM < 0.1 Then Err.Raise vbObjectError + 32, , "KA_AG_BOYU = " & JN(meshM) & " m is too small (0.1 m at least)."
    If tFt <= 0 Or tSt <= 0 Or tBw <= 0 Or tWg <= 0 Or hStem <= 0 Or hBack <= 0 Then Err.Raise vbObjectError + 33, , "Thicknesses (footing, stem, backwall, wing walls) and heights must be > 0."
    ' ear walls (owner M23): KA_KULAK_BOYU = 0 or KA_KULAK_KALINLIK = 0 = no ear walls
    If lEar < 0 Or tEr < 0 Then Err.Raise vbObjectError + 33, , "KA_KULAK_BOYU and KA_KULAK_KALINLIK must be >= 0 (0 = no ear walls)."
    earOn = (lEar > EPS And tEr > EPS)
    If earOn And (hRoot <= 0 Or hTip < 0) Then Err.Raise vbObjectError + 33, , "Ear-wall heights must be > 0 (KA_KULAK_KOK_YUKSEKLIK, KA_KULAK_UC_YUKSEKLIK)."
    If HasName("KA_VEREVLIK") Then
        thetaDeg = NV("KA_VEREVLIK")
    Else
        thetaDeg = 0
        Warn "Name KA_VEREVLIK not found: the abutment is built square (skew 0)."
    End If
    If Abs(thetaDeg) >= 60 Then Err.Raise vbObjectError + 34, , "Skew " & JN(thetaDeg) & " deg cannot be built (60 deg at most)."
    tanT = Tan(thetaDeg * PI_ / 180#): cosT = Cos(thetaDeg * PI_ / 180#)
    ' span type: A1 = first, A2 = last filled column of G_HESAP_ACIKLIK (owner 2026-11-10)
    spanCol = 1
    If abutNm = "A2" Then
        Set r = NmRange(ThisWorkbook, "G_HESAP_ACIKLIK")
        For i = 1 To r.Cells.Count
            If NumOr0(r.Cells(i).Value) > 0 Then spanCol = i
        Next i
    End If
    nGird = CLng(Num(NmRange(ThisWorkbook, "G_KIRIS_SAYISI").Cells(spanCol).Value, "G_KIRIS_SAYISI span type " & spanCol))
    If nGird < 1 Then Err.Raise vbObjectError + 35, , "G_KIRIS_SAYISI span type " & spanCol & " must be >= 1."
    gi = 1
    If HasName("G_KIRIS_GABARI") Then
        v = NmRange(ThisWorkbook, "G_KIRIS_GABARI").Cells(spanCol).Value
        If Not IsError(v) Then
            If IsNumeric(v) And Not IsEmpty(v) Then gi = CLng(v)
        End If
    End If
    Set r = NmRange(ThisWorkbook, "G_YATAY_GABARI")
    If gi < 1 Or gi > r.Cells.Count Then gi = 1
    wOut = Num(r.Cells(gi).Value, "G_YATAY_GABARI " & gi)
    Set r = NmRange(ThisWorkbook, "G_DIGER_BETONLAR")
    clsConc = Trim$(CStr(r.Cells(1, 1).Value))
    eConc = Num(r.Cells(1, 2).Value, "G_DIGER_BETONLAR E (MPa)") * 1000#
    denConc = NV("G_BIRIM_AGIRLIK_BETONARME")
    If piled Then
        pDia = NV("KA_KAZIK_CAP"): nPx = CLng(NV("KA_KAZIK_ADET_ENINE")): nPy = CLng(NV("KA_KAZIK_ADET_BOYUNA"))
        sPx = NV("KA_KAZIK_ARALIK_ENINE"): sPy = NV("KA_KAZIK_ARALIK_BOYUNA")
        lPile = NV("KA_KAZIK_BOY"): dzPile = NV("KA_KAZIK_ELEMAN"): kvTip = NV("KA_KAZIK_UC_KV"): fPile = NV("KA_RIJITLIK_KAZIK")
        If nPx < 1 Or nPy < 1 Or pDia <= 0 Or lPile <= 0 Or dzPile <= 0 Then Err.Raise vbObjectError + 36, , "Pile count, diameter, length and element length must be > 0."
        Set r = NmRange(ThisWorkbook, "KA_ZEMIN_TABAKA")
        nLay = 0
        ReDim layD(1 To r.Rows.Count): ReDim layK(1 To r.Rows.Count)
        For i = 1 To r.Rows.Count
            If NumOr0(r.Cells(i, 1).Value) > 0 And NumOr0(r.Cells(i, 2).Value) > 0 Then
                nLay = nLay + 1: layD(nLay) = CDbl(r.Cells(i, 1).Value): layK(nLay) = CDbl(r.Cells(i, 2).Value)
                If nLay > 1 Then
                    If layD(nLay) <= layD(nLay - 1) Then Err.Raise vbObjectError + 37, , "KA_ZEMIN_TABAKA: layer bottoms must increase (row " & i & ")."
                End If
            End If
        Next i
        If nLay = 0 Then Err.Raise vbObjectError + 38, , "KA_ZEMIN_TABAKA holds no layer (bottom depth and kh > 0)."
    Else
        kvFoot = NV("KA_TEMEL_KV"): khFoot = NV("KA_TEMEL_KH")
        If kvFoot <= 0 Then Err.Raise vbObjectError + 39, , "KA_TEMEL_KV must be > 0."
    End If
    qDC = NV("KA_Y_DC"): qLL = NV("KA_Y_LL"): qTUX = NV("KA_Y_TUX"): qTUY = NV("KA_Y_TUY")
    qWSX = NV("KA_Y_WSX"): qWSY = NV("KA_Y_WSY"): qCF = NV("KA_Y_CF"): qBR = NV("KA_Y_BR"): qPX = NV("KA_Y_PEQX"): qPY = NV("KA_Y_PEQY")
    pEH0 = NV("KA_P_EH0"): pEHA = NV("KA_P_EHA"): pEV = NV("KA_P_EV"): pLS = NV("KA_P_LS")
    pEQX = NV("KA_P_EQX"): pEQY = NV("KA_P_EQY"): pEQs = NV("KA_P_EQ_EGIM"): pPEQX = NV("KA_P_PEQX"): pPEQY = NV("KA_P_PEQY"): khSeis = NV("KA_KH")
    nSeatWant = CLng(NumOr0(NmRange(ThisWorkbook, "KA_MESNET_NOKTA").Cells(1, 1).Value))
    If Abs(qCF) > EPS Then Warn "KA_Y_CF = " & JN(qCF) & " is not applied: case CF stays empty as in the template (OPEN for the owner)."
    projName = NSOr("G_PROJE_ADI"): clientName = NSOr("G_IDARE_ADI"): bridgeName = NSOr("G_KOPRU_ADI")
    SetCases
    ReadCombos
End Sub

Private Sub SetCases()
    Dim a As Variant, i As Long, p As Variant
    a = Array("DC|D", "EV|EV", "EH0|EH", "EHA|EH", "LL|L", "LS|LS", "EQ-X|E", "EQ-Y|E", "TU|T", "WS|W", "PEQ-X|E", "PEQ-Y|E", "BR/ACC|BRK", "CF|CF")
    For i = 1 To N_CASE
        p = Split(a(i - 1), "|")
        caseNm(i) = p(0): caseTy(i) = p(1)
    Next i
End Sub

' KA_KOMBINASYON (LC!A1:AN27): row 1 headers, then one combination per row:
' A name, C Add / Envelope, D..Q factors of the 14 cases (fixed order), R.. factors of combinations
Private Sub ReadCombos()
    Dim r As Range, i As Long, c As Long, nm As String, kind As String, it As String, v As Variant, hdr As String, k As Long, j As Long, chg As Boolean, refs As Object, p As Variant, q As Variant
    Dim st As String, act As Boolean, tier As Long, maxT As Long
    Set r = NmRange(ThisWorkbook, "KA_KOMBINASYON")
    If r.Columns.Count < 17 Then Err.Raise vbObjectError + 40, , "KA_KOMBINASYON must reach column Q (14 load cases)."
    For c = 4 To 17
        hdr = Trim$(CStr(r.Cells(1, c).Value))
        If StrComp(hdr, caseNm(c - 3), vbTextCompare) <> 0 And Not (c = 15 And hdr = "PEQ-X") Then
            Warn "KA_KOMBINASYON header " & ColName(r.Cells(1, c).Column) & "1 = '" & hdr & "', read as " & caseNm(c - 3) & " (column order is fixed)."
        End If
    Next c
    nCmb = 0
    ReDim cmbName(1 To r.Rows.Count): ReDim cmbKind(1 To r.Rows.Count): ReDim cmbItems(1 To r.Rows.Count): ReDim cmbTier(1 To r.Rows.Count)
    ReDim cmbAct(1 To r.Rows.Count): ReDim cmbEff(1 To r.Rows.Count)
    For i = 2 To r.Rows.Count
        v = r.Cells(i, 1).Value
        If IsError(v) Then GoTo NextRow
        nm = Trim$(CStr(v))
        If nm = "" Then GoTo NextRow
        kind = UCase$(Trim$(CStr(r.Cells(i, 3).Value)))
        ' column B: Active / Inactive (owner 2026-11-10); empty = Active
        v = r.Cells(i, 2).Value
        st = ""
        If Not IsError(v) Then st = UCase$(Trim$(CStr(v)))
        act = (st <> "INACTIVE")
        If st <> "" And st <> "ACTIVE" And st <> "INACTIVE" Then Warn "KA_KOMBINASYON row " & i & " (" & nm & "): column B = '" & st & "' is read as Active."
        it = ""
        If Left$(kind, 3) = "ENV" Then
            For c = 18 To r.Columns.Count
                v = r.Cells(i, c).Value
                If Not IsError(v) Then
                    If IsNumeric(v) And Not IsEmpty(v) Then
                        If CDbl(v) <> 0 Then it = it & ";CB|" & Trim$(CStr(r.Cells(1, c).Value)) & "|" & JN(CDbl(v))
                    End If
                End If
            Next c
            nCmb = nCmb + 1: cmbKind(nCmb) = 1
        Else
            For c = 4 To 17
                v = r.Cells(i, c).Value
                If Not IsError(v) Then
                    If IsNumeric(v) And Not IsEmpty(v) Then
                        If CDbl(v) <> 0 Then it = it & ";ST|" & caseNm(c - 3) & "|" & JN(CDbl(v))
                    End If
                End If
            Next c
            nCmb = nCmb + 1: cmbKind(nCmb) = 0
        End If
        If it = "" And act Then Err.Raise vbObjectError + 41, , "KA_KOMBINASYON row " & i & " (" & nm & ") has no factor."
        cmbName(nCmb) = nm: cmbItems(nCmb) = Mid$(it, 2): cmbAct(nCmb) = act
NextRow:
    Next i
    If nCmb = 0 Then Err.Raise vbObjectError + 42, , "KA_KOMBINASYON holds no combination."
    ' tiers: a combination that refers to others goes after them
    For k = 1 To 6
        chg = False
        For i = 1 To nCmb
            For Each p In Split(cmbItems(i), ";")
                If Len(p) > 0 Then
                    q = Split(p, "|")
                    If q(0) = "CB" Then
                        j = CmbIndex(CStr(q(1)))
                        If j = 0 Then Err.Raise vbObjectError + 43, , "Combination " & cmbName(i) & " refers to '" & q(1) & "', which is not in KA_KOMBINASYON."
                        If cmbTier(i) < cmbTier(j) + 1 Then cmbTier(i) = cmbTier(j) + 1: chg = True
                    End If
                End If
            Next p
        Next i
        If Not chg Then Exit For
    Next k
    ' every active 'Add' row becomes a factored load case (owner 2026-11-10); column B 'Inactive' = skipped
    nFac = 0
    ReDim facIdx(1 To nCmb)
    For i = 1 To nCmb
        cmbEff(i) = False
        If cmbKind(i) = 0 And cmbAct(i) Then nFac = nFac + 1: facIdx(nFac) = i: cmbEff(i) = True
    Next i
    If nFac = 0 Then Err.Raise vbObjectError + 44, , "KA_KOMBINASYON holds no active 'Add' row (no load case)."
    ' an envelope is written when it is active and keeps at least one written member (lower tiers first);
    ' inactive members are left out of it
    maxT = 0
    For i = 1 To nCmb
        If cmbTier(i) > maxT Then maxT = cmbTier(i)
    Next i
    For tier = 1 To maxT
        For i = 1 To nCmb
            If cmbKind(i) = 1 And cmbTier(i) = tier And cmbAct(i) Then
                For Each p In Split(cmbItems(i), ";")
                    If Len(p) > 0 Then
                        q = Split(p, "|")
                        j = CmbIndex(CStr(q(1)))
                        If j > 0 Then
                            If cmbEff(j) Then cmbEff(i) = True
                        End If
                    End If
                Next p
                If Not cmbEff(i) Then Warn "Envelope " & cmbName(i) & " is not written: none of its members is active."
            End If
        Next i
    Next tier
End Sub

Private Function CmbIndex(ByVal nm As String) As Long
    Dim i As Long
    For i = 1 To nCmb
        If StrComp(cmbName(i), nm, vbTextCompare) = 0 Then CmbIndex = i: Exit Function
    Next i
End Function

Private Function CaseIndex(ByVal nm As String) As Long
    Dim i As Long
    For i = 1 To N_CASE
        If StrComp(caseNm(i), nm, vbTextCompare) = 0 Then CaseIndex = i: Exit Function
    Next i
End Function

' ============================================================================
'  geometry and mesh
' ============================================================================
Private Sub BuildGeometry()
    Dim i As Long, j As Long, k As Long, a As Double, nB As Long
    Dim bLin() As Double, bPri() As Long
    e0 = (fB - wOut) / 2#
    If e0 < -EPS Then Err.Raise vbObjectError + 50, , "Wing walls " & JN(wOut) & " m out to out (G_YATAY_GABARI) are wider than the footing " & JN(fB) & " m (KA_TEMEL_EN)."
    If e0 < 0 Then e0 = 0
    xW1 = e0 + tWg / 2#: xW2 = fB - e0 - tWg / 2#
    fi1 = e0 + tWg: fi2 = fB - e0 - tWg
    If fi2 - fi1 < 0.5 Then Err.Raise vbObjectError + 51, , "Clear width between the wing walls is " & JN(fi2 - fi1) & " m."
    tStV = tSt / cosT
    ' owner M21: heel, toe and footing length are measured normal to the (skewed) stem; v runs along the bridge axis
    fLv = fL / cosT
    vBk = heelW / cosT: vFr = vBk + tStV: yS = vBk + tStV / 2#
    toeUsed = (fLv - vFr) * cosT
    If toeUsed < 0.05 Then Err.Raise vbObjectError + 52, , "Heel " & JN(heelW) & " + stem " & JN(tSt) & " leave no toe on a footing of " & JN(fL) & " m (normal to the stem)."
    If Abs(toeUsed - toeW) > 0.01 Then Warn "Toe from the footing length = " & JN(Round(toeUsed, 3)) & " m (KA_ON_GENISLIK = " & JN(toeW) & "); the model uses " & JN(Round(toeUsed, 3)) & "."
    zFt = tFt / 2#: zSeat = zFt + hStem: zTop = zSeat + hBack
    If piled Then
        cPx = (fB - (nPx - 1) * sPx) / 2#: cPy = (fL - (nPy - 1) * sPy) / 2#
        If cPx < pDia / 2# - EPS Or cPy < pDia / 2# - EPS Then Err.Raise vbObjectError + 53, , "Piles do not fit the footing (edge distance " & JN(Round(cPx, 3)) & " / " & JN(Round(cPy, 3)) & " m < D/2)."
        nP = nPx * nPy
        nSeg = Int(lPile / dzPile + 0.5): If nSeg < 1 Then nSeg = 1
        segL = lPile / nSeg
    End If
    ' footing lines across (u)
    nB = 0: ReDim bLin(1 To 64): ReDim bPri(1 To 64)
    PushL bLin, bPri, nB, 0, 2: PushL bLin, bPri, nB, fB, 2: PushL bLin, bPri, nB, e0, 2: PushL bLin, bPri, nB, xW1, 2: PushL bLin, bPri, nB, fi1, 2
    PushL bLin, bPri, nB, fi2, 2: PushL bLin, bPri, nB, xW2, 2: PushL bLin, bPri, nB, fB - e0, 2
    If piled Then
        For i = 0 To nPx - 1
            a = cPx + i * sPx
            PushPile bLin, bPri, nB, a, True
        Next i
    End If
    LinesFrom bLin, bPri, nB, 1.1 * meshM, uF, nUF, "Footing lines across"
    ' footing lines along (v); pile rows at their normal distance / cos(skew) (M21)
    nB = 0: ReDim bLin(1 To 64): ReDim bPri(1 To 64)
    PushL bLin, bPri, nB, 0, 2: PushL bLin, bPri, nB, fLv, 2: PushL bLin, bPri, nB, vBk, 2: PushL bLin, bPri, nB, yS, 2: PushL bLin, bPri, nB, vFr, 2
    If piled Then
        For j = 0 To nPy - 1
            a = (cPy + j * sPy) / cosT
            PushPile bLin, bPri, nB, a, False
        Next j
    End If
    LinesFrom bLin, bPri, nB, 1.1 * meshM, vFt, nVF, "Footing lines along"
    ' wall levels
    nZW = 0: ReDim zW(1 To 256)
    PushV zW, nZW, 0: PushV zW, nZW, zFt
    k = Int(hStem / meshM + 0.5): If k < 1 Then k = 1
    For i = 1 To k: PushV zW, nZW, zFt + hStem * i / k: Next i
    k = Int(hBack / meshM + 0.5): If k < 1 Then k = 1
    For i = 1 To k: PushV zW, nZW, zSeat + hBack * i / k: Next i
    ' ear walls: 0 back to -lEar
    nVE = 0: ReDim vE(1 To 64)
    k = Int(lEar / (1.5 * meshM) + 0.5): If k < 1 Then k = 1
    If earOn Then
        For i = 0 To k: PushV vE, nVE, -lEar * i / k: Next i
    End If
    ' ---- elements ----
    Set nodeKey = CreateObject("Scripting.Dictionary")
    nCnt = 0: eCnt = 0
    ReDim nId(1 To 4096): ReDim nX(1 To 4096): ReDim nY(1 To 4096): ReDim nZ(1 To 4096)
    ReDim eId(1 To 4096): ReDim eKind(1 To 4096): ReDim eMem(1 To 4096): ReDim eThk(1 To 4096): ReDim eNds(1 To 4096, 1 To 4): ReDim eNn(1 To 4096)
    ReDim eCu(1 To 4096): ReDim eCv(1 To 4096): ReDim eCz(1 To 4096): ReDim eRjt(1 To 4096): ReDim ePrj(1 To 4096): ReDim ePssf(1 To 4096)
    BuildFooting
    BuildWall M_STEM
    BuildWall M_BACK
    BuildWing M_WINGL
    BuildWing M_WINGR
    If nVE > 1 Then
        BuildEar M_EARL
        BuildEar M_EARR
    End If
    nPlate = eCnt
    If nCnt >= PILE0 Then Err.Raise vbObjectError + 54, , "Too many plate nodes (" & nCnt & "): numbers must stay below " & PILE0 & ". Use a larger KA_AG_BOYU."
    FindSeat
    If piled Then BuildPiles
End Sub

Private Sub PushV(arr() As Double, ByRef n As Long, ByVal v As Double)
    n = n + 1
    If n > UBound(arr) Then ReDim Preserve arr(1 To 2 * UBound(arr))
    arr(n) = v
End Sub

' pile axis (priority 1) and the two lines of the square inscribed in the pile circle (priority 0);
' across the bridge the square lines are skipped inside a wing-wall footprint (template); along the
' bridge they keep the normal distance d to the pile row (d / cos(skew) along v)
Private Sub PushPile(arr() As Double, pri() As Long, ByRef n As Long, ByVal a As Double, ByVal across As Boolean)
    Dim d As Double, k As Long, v As Double
    d = pDia / (2# * Sqr(2#))
    If Not across Then d = d / cosT
    For k = -1 To 1
        v = a + k * d
        If across And k <> 0 Then
            If (v > e0 + EPS And v < fi1 - EPS) Or (v > fi2 + EPS And v < fB - e0 - EPS) Then GoTo NextK
        End If
        PushL arr, pri, n, v, IIf(k = 0, 1, 0)
NextK:
    Next k
End Sub

Private Sub PushL(arr() As Double, pri() As Long, ByRef n As Long, ByVal v As Double, ByVal p As Long)
    n = n + 1
    If n > UBound(arr) Then
        ReDim Preserve arr(1 To 2 * UBound(arr)): ReDim Preserve pri(1 To 2 * UBound(pri))
    End If
    arr(n) = v: pri(n) = p
End Sub

' sorted unique lines inside [first, last of the outline], each gap in round(gap / stp) parts.
' Priority 2 = outline and walls, 1 = pile axis, 0 = pile-square line. A pile-square line closer than
' 0.2 x mesh to another line is snapped onto it (no sliver plates); two such close lines of priority
' >= 1 are both kept and reported.
Private Sub LinesFrom(bLin() As Double, bPri() As Long, ByVal nB As Long, ByVal stp As Double, out() As Double, ByRef nOut As Long, ByVal what As String)
    Dim i As Long, j As Long, t As Double, tp As Long, k As Long, lo As Double, hi As Double, u() As Double, up() As Long, nU As Long
    Dim tol As Double, nSnap As Long, nClose As Long, mv As Double, mvMax As Double
    lo = bLin(1): hi = bLin(2)
    tol = 0.2 * meshM
    For i = 2 To nB
        t = bLin(i): tp = bPri(i): j = i - 1
        Do While j >= 1
            If bLin(j) <= t Then Exit Do
            bLin(j + 1) = bLin(j): bPri(j + 1) = bPri(j): j = j - 1
        Loop
        bLin(j + 1) = t: bPri(j + 1) = tp
    Next i
    ReDim u(1 To nB): ReDim up(1 To nB): nU = 0
    For i = 1 To nB
        If bLin(i) >= Min2(lo, hi) - EPS And bLin(i) <= Max2(lo, hi) + EPS Then
            If nU = 0 Then
                nU = 1: u(1) = bLin(i): up(1) = bPri(i)
            ElseIf bLin(i) - u(nU) <= 0.0001 Then
                If bPri(i) > up(nU) Then up(nU) = bPri(i)
            ElseIf bLin(i) - u(nU) < tol And (bPri(i) = 0 Or up(nU) = 0) Then
                mv = bLin(i) - u(nU): If mv > mvMax Then mvMax = mv
                nSnap = nSnap + 1
                If bPri(i) > up(nU) Then u(nU) = bLin(i): up(nU) = bPri(i)
            Else
                If bLin(i) - u(nU) < tol Then nClose = nClose + 1
                nU = nU + 1: u(nU) = bLin(i): up(nU) = bPri(i)
            End If
        End If
    Next i
    If nSnap > 0 Then Warn what & ": " & nSnap & " pile-square mesh line(s) closer than " & JN(tol) & " m to another line snapped onto it (largest move " & JN(Round(mvMax, 3)) & " m)."
    If nClose > 0 Then Warn what & ": " & nClose & " pair(s) of wall / pile lines closer than " & JN(tol) & " m; the narrow plates between them are kept."
    nOut = 0: ReDim out(1 To 512)
    PushV out, nOut, u(1)
    For i = 2 To nU
        k = Int((u(i) - u(i - 1)) / stp + 0.5): If k < 1 Then k = 1
        For j = 1 To k
            PushV out, nOut, u(i - 1) + (u(i) - u(i - 1)) * j / k
        Next j
    Next i
End Sub

Private Function Min2(ByVal a As Double, ByVal b As Double) As Double
    Min2 = IIf(a < b, a, b)
End Function

Private Function Max2(ByVal a As Double, ByVal b As Double) As Double
    Max2 = IIf(a > b, a, b)
End Function

Private Function InRange(ByVal x As Double, ByVal a As Double, ByVal b As Double) As Boolean
    InRange = (x >= a - EPS And x <= b + EPS)
End Function

Private Sub BuildFooting()
    Dim i As Long, j As Long, n(1 To 4) As Long, cu As Double, cv As Double, rj As Boolean, pr As Boolean
    For i = 1 To nUF - 1
        For j = 1 To nVF - 1
            n(1) = NodeAt(uF(i), vFt(j), 0): n(2) = NodeAt(uF(i + 1), vFt(j), 0)
            n(3) = NodeAt(uF(i + 1), vFt(j + 1), 0): n(4) = NodeAt(uF(i), vFt(j + 1), 0)
            cu = (uF(i) + uF(i + 1)) / 2#: cv = (vFt(j) + vFt(j + 1)) / 2#
            rj = ((InRange(cu, e0, fi1) Or InRange(cu, fi2, fB - e0)) And cv <= vFr + EPS) Or (InRange(cv, vBk, vFr) And InRange(cu, e0, fB - e0))
            AddPlate M_FOOT, 1, n, 4, 0, 0, 1, cu, cv, 0, rj, Not rj
        Next j
    Next i
End Sub

Private Sub BuildWall(ByVal mem As Long)
    Dim i As Long, k As Long, n(1 To 4) As Long, cu As Double, cz As Double, rj As Boolean, zA As Double, zB As Double
    zA = IIf(mem = M_STEM, 0, zSeat): zB = IIf(mem = M_STEM, zSeat, zTop)
    For i = 1 To nUF - 1
        If uF(i) >= xW1 - EPS And uF(i + 1) <= xW2 + EPS Then
            For k = 1 To nZW - 1
                If zW(k) >= zA - EPS And zW(k + 1) <= zB + EPS Then
                    n(1) = NodeAt(uF(i), yS, zW(k)): n(2) = NodeAt(uF(i + 1), yS, zW(k))
                    n(3) = NodeAt(uF(i + 1), yS, zW(k + 1)): n(4) = NodeAt(uF(i), yS, zW(k + 1))
                    cu = (uF(i) + uF(i + 1)) / 2#: cz = (zW(k) + zW(k + 1)) / 2#
                    rj = (cz < zFt) Or (cu < fi1) Or (cu > fi2)
                    ' local z toward the fill (-y)
                    AddPlate mem, IIf(mem = M_STEM, 4, 3), n, 4, 0, -1, 0, cu, yS, cz, rj, (mem = M_BACK) Or (cz > zFt)
                End If
            Next k
        End If
    Next i
End Sub

Private Sub BuildWing(ByVal mem As Long)
    Dim j As Long, k As Long, n(1 To 4) As Long, x As Double, cv As Double, cz As Double, rj As Boolean
    x = IIf(mem = M_WINGL, xW1, xW2)
    For j = 1 To nVF - 1
        If vFt(j + 1) <= yS + EPS Then
            For k = 1 To nZW - 1
                n(1) = NodeAt(x, vFt(j), zW(k)): n(2) = NodeAt(x, vFt(j + 1), zW(k))
                n(3) = NodeAt(x, vFt(j + 1), zW(k + 1)): n(4) = NodeAt(x, vFt(j), zW(k + 1))
                cv = (vFt(j) + vFt(j + 1)) / 2#: cz = (zW(k) + zW(k + 1)) / 2#
                rj = (cz < zFt) Or (cv > vBk)
                AddPlate mem, 2, n, 4, IIf(mem = M_WINGL, 1, -1), 0, 0, x, cv, cz, rj, Not rj
            Next k
        End If
    Next j
End Sub

' ear wall: bottom edge from zTop - hRoot at v = 0 to zTop - hTip at v = -lEar, snapped to the
' wall levels; a column whose bottom steps is closed by triangles
Private Sub BuildEar(ByVal mem As Long)
    Dim i As Long, k As Long, x As Double, b() As Long, lo As Long, hi As Long, n(1 To 4) As Long, deep As Double, shal As Double, zb As Double, nrm As Double
    x = IIf(mem = M_EARL, xW1, xW2): nrm = IIf(mem = M_EARL, 1, -1)
    ReDim b(1 To nVE)
    For i = 1 To nVE
        zb = zTop - hRoot + (hRoot - hTip) * (-vE(i) / lEar)
        b(i) = NearestLevel(zb)
    Next i
    For i = 1 To nVE - 1
        lo = IIf(b(i) < b(i + 1), b(i), b(i + 1)): hi = IIf(b(i) > b(i + 1), b(i), b(i + 1))
        For k = hi To nZW - 1
            n(1) = NodeAt(x, vE(i + 1), zW(k)): n(2) = NodeAt(x, vE(i), zW(k))
            n(3) = NodeAt(x, vE(i), zW(k + 1)): n(4) = NodeAt(x, vE(i + 1), zW(k + 1))
            AddPlate mem, 5, n, 4, nrm, 0, 0, x, (vE(i) + vE(i + 1)) / 2#, (zW(k) + zW(k + 1)) / 2#, False, False
        Next k
        If hi > lo Then
            ' fan from the shallow side's lowest node
            If b(i) < b(i + 1) Then deep = vE(i): shal = vE(i + 1) Else deep = vE(i + 1): shal = vE(i)
            For k = lo To hi - 1
                n(1) = NodeAt(x, shal, zW(hi)): n(2) = NodeAt(x, deep, zW(k)): n(3) = NodeAt(x, deep, zW(k + 1)): n(4) = 0
                AddPlate mem, 5, n, 3, nrm, 0, 0, x, (2 * deep + shal) / 3#, (zW(hi) + zW(k) + zW(k + 1)) / 3#, False, False
            Next k
        End If
    Next i
End Sub

Private Function NearestLevel(ByVal z As Double) As Long
    Dim k As Long, best As Long, d As Double
    best = 1: d = 1E+30
    For k = 1 To nZW
        If Abs(zW(k) - z) < d Then d = Abs(zW(k) - z): best = k
    Next k
    If best >= nZW Then best = nZW - 1
    NearestLevel = best
End Function

Private Function XOf(ByVal u As Double, ByVal v As Double) As Double
    XOf = u
End Function

Private Function YOf(ByVal u As Double, ByVal v As Double) As Double
    YOf = v + (u - fB / 2#) * tanT
End Function

Private Function NodeAt(ByVal u As Double, ByVal v As Double, ByVal z As Double) As Long
    Dim k As String, x As Double, y As Double
    x = XOf(u, v): y = YOf(u, v)
    k = KeyOf(x, y, z)
    If nodeKey.Exists(k) Then NodeAt = nodeKey(k): Exit Function
    nCnt = nCnt + 1
    If nCnt > UBound(nX) Then
        ReDim Preserve nId(1 To 2 * UBound(nX)): ReDim Preserve nY(1 To 2 * UBound(nX)): ReDim Preserve nZ(1 To 2 * UBound(nX)): ReDim Preserve nX(1 To 2 * UBound(nX))
    End If
    nId(nCnt) = nCnt: nX(nCnt) = x: nY(nCnt) = y: nZ(nCnt) = z
    nodeKey(k) = nCnt
    NodeAt = nCnt
End Function

Private Function NodeFind(ByVal u As Double, ByVal v As Double, ByVal z As Double) As Long
    Dim k As String
    k = KeyOf(XOf(u, v), YOf(u, v), z)
    If nodeKey.Exists(k) Then NodeFind = nodeKey(k)
End Function

Private Function KeyOf(ByVal x As Double, ByVal y As Double, ByVal z As Double) As String
    KeyOf = Trim$(Str$(Round(x, 4))) & "|" & Trim$(Str$(Round(y, 4))) & "|" & Trim$(Str$(Round(z, 4)))
End Function

Private Function NodeIdx(ByVal id As Long) As Long
    Dim i As Long
    If id < PILE0 Then NodeIdx = id: Exit Function
    For i = nCnt To 1 Step -1
        If nId(i) = id Then NodeIdx = i: Exit Function
    Next i
End Function

' plate with nn = 3 or 4 nodes; the node order is reversed when its normal points away from (dx, dy, dz)
Private Sub AddPlate(ByVal mem As Long, ByVal thk As Long, n() As Long, ByVal nn As Long, ByVal dx As Double, ByVal dy As Double, ByVal dz As Double, _
                     ByVal cu As Double, ByVal cv As Double, ByVal cz As Double, ByVal rj As Boolean, ByVal ps As Boolean)
    Dim ax As Double, ay As Double, az As Double, bx As Double, by As Double, bz As Double, cx As Double, cy As Double, cz2 As Double, t As Long, last As Long
    last = nn
    ax = nX(n(2)) - nX(n(1)): ay = nY(n(2)) - nY(n(1)): az = nZ(n(2)) - nZ(n(1))
    bx = nX(n(last)) - nX(n(1)): by = nY(n(last)) - nY(n(1)): bz = nZ(n(last)) - nZ(n(1))
    cx = ay * bz - az * by: cy = az * bx - ax * bz: cz2 = ax * by - ay * bx
    If cx * dx + cy * dy + cz2 * dz < 0 Then
        t = n(2): n(2) = n(last): n(last) = t
    End If
    eCnt = eCnt + 1
    GrowE
    eId(eCnt) = eCnt: eKind(eCnt) = 1: eMem(eCnt) = mem: eThk(eCnt) = thk: eNn(eCnt) = nn
    For t = 1 To 4: eNds(eCnt, t) = IIf(t <= nn, n(t), 0): Next t
    eCu(eCnt) = cu: eCv(eCnt) = cv: eCz(eCnt) = cz: eRjt(eCnt) = rj: ePssf(eCnt) = ps
End Sub

Private Sub GrowE()
    Dim m As Long, a() As Long, i As Long, j As Long
    If eCnt <= UBound(eId) Then Exit Sub
    m = 2 * UBound(eId)
    ReDim Preserve eId(1 To m): ReDim Preserve eKind(1 To m): ReDim Preserve eMem(1 To m): ReDim Preserve eThk(1 To m): ReDim Preserve eNn(1 To m)
    ReDim Preserve eCu(1 To m): ReDim Preserve eCv(1 To m): ReDim Preserve eCz(1 To m): ReDim Preserve eRjt(1 To m): ReDim Preserve ePrj(1 To m): ReDim Preserve ePssf(1 To m)
    ReDim a(1 To m, 1 To 4)
    For i = 1 To UBound(eNds, 1)
        For j = 1 To 4: a(i, j) = eNds(i, j): Next j
    Next i
    eNds = a
End Sub

' dry run: the workbook is not written, only the difference is reported
Private Sub SeatWarnDry()
    If nSeatWant <> nSeat Then Warn "Seat nodes: the mesh has " & nSeat & " stem-top nodes between the wing walls, KA_MESNET_NOKTA (1_RAPOR AH190) says " & nSeatWant & "; the loads are split over " & nSeat & ". A build writes " & nSeat & " into KA_MESNET_NOKTA."
End Sub

' live build (owner 2026-11-11): the report's per-node seat values follow the mesh; a formula in the cell is left alone
Private Function SeatWrite() As String
    Dim c As Range, was As String
    On Error GoTo Fail
    Set c = NmRange(ThisWorkbook, "KA_MESNET_NOKTA").Cells(1, 1)
    was = CStr(c.Value)
    If c.HasFormula Then
        If nSeatWant <> nSeat Then Warn "Seat nodes: the mesh has " & nSeat & ", KA_MESNET_NOKTA holds a formula giving " & was & "; not overwritten."
        SeatWrite = "  seat nodes " & nSeat & " (KA_MESNET_NOKTA is a formula, not written)" & vbCrLf
    ElseIf nSeatWant = nSeat Then
        SeatWrite = "  seat nodes " & nSeat & " = KA_MESNET_NOKTA (1_RAPOR AH190)" & vbCrLf
    Else
        c.Value = nSeat
        SeatWrite = "  seat nodes " & nSeat & " written to KA_MESNET_NOKTA (1_RAPOR AH190, was " & was & "); save the workbook" & vbCrLf
    End If
    Exit Function
Fail:
    Warn "Seat nodes: could not write " & nSeat & " into KA_MESNET_NOKTA (" & Err.Description & ")."
    SeatWrite = ""
End Function

' seat nodes: stem top between the wing walls' inner faces
Private Sub FindSeat()
    Dim i As Long, nd As Long
    ReDim seatN(1 To nUF): nSeat = 0
    For i = 1 To nUF
        If uF(i) >= fi1 - EPS And uF(i) <= fi2 + EPS Then
            nd = NodeFind(uF(i), yS, zSeat)
            If nd > 0 Then nSeat = nSeat + 1: seatN(nSeat) = nd
        End If
    Next i
    If nSeat = 0 Then Err.Raise vbObjectError + 55, , "No stem-top node between the wing walls."
End Sub

' piles: node / element 14000 + nP s + p, p = nPy * column + row
Private Sub BuildPiles()
    Dim i As Long, j As Long, s As Long, p As Long, u As Double, v As Double, x As Double, y As Double, m As Long, id As Long, k As Double, kz As Double, d As Double
    Dim sj As String
    ReDim pileTop(0 To nP - 1): ReDim pileMaster(0 To nP - 1)
    springJson = "": nSpring = 0
    For i = 0 To nPx - 1
        For j = 0 To nPy - 1
            p = nPy * i + j
            u = cPx + i * sPx: v = (cPy + j * sPy) / cosT
            m = NodeFind(u, v, 0)
            If m = 0 Then Err.Raise vbObjectError + 56, , "Pile " & p + 1 & " is not on a footing node."
            pileMaster(p) = m
            x = XOf(u, v): y = YOf(u, v)
            For s = 0 To nSeg
                id = PILE0 + nP * s + p
                nCnt = nCnt + 1
                If nCnt > UBound(nX) Then
                    ReDim Preserve nId(1 To 2 * UBound(nX)): ReDim Preserve nY(1 To 2 * UBound(nX)): ReDim Preserve nZ(1 To 2 * UBound(nX)): ReDim Preserve nX(1 To 2 * UBound(nX))
                End If
                nId(nCnt) = id: nX(nCnt) = x: nY(nCnt) = y: nZ(nCnt) = -zFt - s * segL
                If s = 0 Then pileTop(p) = id
                d = s * segL
                k = LayerK(d) * pDia * segL
                kz = 0
                If s = nSeg Then kz = kvTip * PI_ * pDia * pDia / 4#
                sj = QS(CStr(id)) & ":{""ITEMS"":[{""ID"":1,""TYPE"":""LINEAR"",""GROUP_NAME"":""Statik"",""FormType"":0,""SDR"":[" & JN(k) & "," & JN(k) & "," & JN(kz) & ",0,0,0]," & _
                     """F_S"":[false,false,false,false,false,false],""DAMPING"":false,""Cr"":[0,0,0,0,0,0],""EFFAREA"":0,""DK"":[0,0,0]}]}"
                springJson = springJson & IIf(nSpring = 0, "", ",") & sj
                nSpring = nSpring + 1
                If s < nSeg Then
                    eCnt = eCnt + 1
                    GrowE
                    eId(eCnt) = id: eKind(eCnt) = 2: eMem(eCnt) = 0: eThk(eCnt) = 1: eNn(eCnt) = 2
                    eNds(eCnt, 1) = id: eNds(eCnt, 2) = id + nP: eNds(eCnt, 3) = 0: eNds(eCnt, 4) = 0
                    eCu(eCnt) = u: eCv(eCnt) = v: eCz(eCnt) = -zFt - (s + 0.5) * segL
                End If
            Next s
        Next j
    Next i
    ' footing quads around each pile node (group KAZIK_RJT)
    For i = 1 To nPlate
        If eMem(i) = M_FOOT Then
            For p = 0 To nP - 1
                For j = 1 To 4
                    If eNds(i, j) = pileMaster(p) Then ePrj(i) = True
                Next j
            Next p
        End If
    Next i
End Sub

' lateral subgrade modulus at depth d below the pile head: first layer whose bottom is deeper
Private Function LayerK(ByVal d As Double) As Double
    Dim i As Long
    For i = 1 To nLay
        If d < layD(i) - EPS Then LayerK = layK(i): Exit Function
    Next i
    LayerK = layK(nLay)
End Function

' ============================================================================
'  loads
' ============================================================================
' nodal pressures of plate e in case c (negative = away from the fill); False = not loaded
Private Function ElemP(ByVal c As String, ByVal e As Long, p() As Double) As Boolean
    Dim mem As Long, earth As Boolean, i As Long, z As Double, nd As Long
    mem = eMem(e)
    If eKind(e) <> 1 Then Exit Function
    earth = (mem = M_STEM Or mem = M_BACK Or mem = M_EARL Or mem = M_EARR)
    If mem = M_WINGL Or mem = M_WINGR Then earth = (eCv(e) <= vBk + EPS And eCz(e) >= zFt - EPS)
    For i = 1 To 4: p(i) = 0: Next i
    Select Case c
        Case "EV"
            If mem <> M_FOOT Then Exit Function
            If Not (InRange(eCu(e), fi1, fi2) And eCv(e) <= vBk + EPS) Then Exit Function
            For i = 1 To eNn(e): p(i) = -pEV: Next i
        Case "EH0", "EHA"
            If Not earth Then Exit Function
            For i = 1 To eNn(e)
                z = nZ(NodeIdx(eNds(e, i)))
                p(i) = -IIf(c = "EH0", pEH0, pEHA) * (zTop - z)
            Next i
        Case "LS"
            If Not earth Then Exit Function
            For i = 1 To eNn(e): p(i) = -pLS: Next i
        Case "EQ-X"
            If Not (mem = M_STEM Or mem = M_BACK) Then Exit Function
            For i = 1 To eNn(e)
                z = nZ(NodeIdx(eNds(e, i)))
                p(i) = -Max2(0, pEQX - pEQs * (zTop - z))
            Next i
        Case "EQ-Y"
            If Not ((mem = TrWing() And earth) Or mem = TrEar()) Then Exit Function
            For i = 1 To eNn(e)
                z = nZ(NodeIdx(eNds(e, i)))
                p(i) = -Max2(0, pEQY - pEQs * (zTop - z))
            Next i
        Case "PEQ-X"
            If Not (mem = M_STEM Or mem = M_BACK) Then Exit Function
            If Not (InRange(eCu(e), fi1, fi2) And eCz(e) >= zFt - EPS) Then Exit Function
            For i = 1 To eNn(e): p(i) = -pPEQX: Next i
        Case "PEQ-Y"
            If Not (mem = TrWing() And earth) Then Exit Function
            For i = 1 To eNn(e): p(i) = -pPEQY: Next i
        Case Else
            Exit Function
    End Select
    ElemP = True
End Function

' wing and ear wall on the side the transverse loads push toward (template: +x)
Private Function TrWing() As Long
    TrWing = IIf(trSign < 0, M_WINGL, M_WINGR)
End Function

Private Function TrEar() As Long
    TrEar = IIf(trSign < 0, M_EARL, M_EARR)
End Function

' force on each seat node of case c (local frame: y along the bridge, x across)
Private Sub SeatF(ByVal c As String, ByRef fx As Double, ByRef fy As Double, ByRef fz As Double)
    Dim g As Double
    fx = 0: fy = 0: fz = 0
    g = nGird / nSeat
    Select Case c
        Case "DC": fz = -qDC * g
        Case "LL": fz = -qLL * g
        Case "TU": fy = qTUX * g: fx = trSign * qTUY * g
        Case "WS": fx = trSign * qWSY * g: fy = qWSX * g
        Case "PEQ-X": fy = qPX * g
        Case "PEQ-Y": fx = trSign * qPY * g
        Case "BR/ACC": fy = qBR * g
    End Select
End Sub

Private Sub BodyF(ByVal c As String, ByRef fx As Double, ByRef fy As Double, ByRef fz As Double)
    fx = 0: fy = 0: fz = 0
    Select Case c
        Case "DC": fz = -1
        Case "PEQ-X": fy = khSeis
        Case "PEQ-Y": fx = trSign * khSeis
    End Select
End Sub

Private Function PresItem(ByVal id As Long, ByVal lc As String, ByVal grp As String, p() As Double, ByVal nn As Long) As String
    Dim uni As Boolean, i As Long, f As String
    uni = True
    For i = 2 To nn
        If Abs(p(i) - p(1)) > 0.0000001 Then uni = False
    Next i
    If uni Then
        f = "[" & JN(p(1)) & ",0,0,0,0]"
    Else
        f = "[0," & JN(p(1)) & "," & JN(p(2)) & "," & JN(p(3)) & "," & JN(IIf(nn = 4, p(4), 0)) & "]"
    End If
    PresItem = "{""ID"":" & id & ",""LCNAME"":" & QS(lc) & ",""GROUP_NAME"":" & QS(grp) & ",""CMD"":""" & IIf(uni, "PRES", "HYDRO") & """,""ELEM_TYPE"":""PLATE"",""FACE_EDGE_TYPE"":""FACE"",""DIRECTION"":""LZ"",""FORCES"":" & f & "}"
End Function

Private Function CnldItem(ByVal id As Long, ByVal lc As String, ByVal grp As String, ByVal fx As Double, ByVal fy As Double, ByVal fz As Double) As String
    CnldItem = "{""ID"":" & id & ",""LCNAME"":" & QS(lc) & ",""GROUP_NAME"":" & QS(grp) & ",""FX"":" & JN(fx) & ",""FY"":" & JN(fy) & ",""FZ"":" & JN(fz) & ",""MX"":0,""MY"":0,""MZ"":0}"
End Function

' load cases of the model: one factored USER case per 'Add' row, named as the LC sheet (both types)
Private Function ModelCases() As Collection
    Dim c As New Collection, i As Long
    For i = 1 To nFac: c.Add cmbName(facIdx(i)) & "|USER": Next i
    Set ModelCases = c
End Function

' factor of component k in model case m
Private Function FactorOf(ByVal m As Long, ByVal k As Long) As Double
    Dim p As Variant, q As Variant
    For Each p In Split(cmbItems(facIdx(m)), ";")
        q = Split(p, "|")
        If q(0) = "ST" And StrComp(CStr(q(1)), caseNm(k), vbTextCompare) = 0 Then FactorOf = FactorOf + Val(q(2))
    Next p
End Function

' load group of component k
Private Function GrpOf(ByVal k As Long) As String
    GrpOf = IIf(caseNm(k) = "BR/ACC", GRP_BR, caseNm(k))
End Function

' factored seismic cases (H-TKDY-X/Y, TKDY-X/Y) run on the cracked boundary combination EQ
Private Function IsSeismicCase(ByVal nm As String) As Boolean
    IsSeismicCase = (InStr(UCase$(nm), "TKDY") > 0)
End Function

' ============================================================================
'  payloads
' ============================================================================
Private Sub BuildPayloads()
    Dim s As String, i As Long, j As Long, k As Long, cnt As Long, mc As Collection, m As Long, nm As String
    Dim p(1 To 4) As Double, acc(1 To 4) As Double, hit As Boolean, f As Double, it As String, nIt As Long
    Dim fx As Double, fy As Double, fz As Double, ax As Double, ay As Double, az As Double
    Dim ar As Double, ux As Double, uy As Double, uz As Double
    Set payloads = New Collection
    nGrupW = 0: nThkW = 0: nOffW = 0: nPssfW = 0: nBodfW = 0: nCnldN = 0: nCnldI = 0: nPresE = 0: nPresI = 0: nSeisW = 0
    For i = 1 To 7: pssfMem(i) = 0: Next i
    ReDim apF(1 To nFac, 1 To 3)
    wPlate = 0
    For i = 1 To nPlate
        PlateGeo i, ar, ux, uy, uz
        wPlate = wPlate + denConc * ThkOf(eThk(i)) * ar
    Next i
    For i = 1 To nFac
        If IsSeismicCase(cmbName(facIdx(i))) Then nSeisW = nSeisW + 1
    Next i
    Add "db/UNIT", "{""Assign"":{""1"":{""FORCE"":""KN"",""DIST"":""M"",""HEAT"":""J"",""TEMPER"":""C""}}}"
    Add "db/STYP", "{""Assign"":{""1"":{""STYP"":0,""MASS"":1,""GRAV"":" & JN(GRAV_STYP) & ",""TEMP"":0,""bALIGNBEAM"":false,""bALIGNSLAB"":false,""bMASSOFFSET"":false,""bSELFWEIGHT"":true,""SMASS"":1}}}"
    Add "db/PJCF", "{""Assign"":{""1"":{""PROJECT"":" & QS(Pj(projName)) & ",""USER"":" & QS(Pj(Application.UserName)) & ",""CLIENT"":" & QS(Pj(clientName)) & _
        ",""TITLE"":" & QS(Pj(bridgeName & " " & abutNm)) & ",""COMMENT"":" & QS(Pj(abutNm & " " & IIf(piled, "KAZIKLI", "YUZEYSEL") & "; " & SCRIPT_ID & " " & SCRIPT_VERSION & " ABUTMENT")) & "}}}"
    Add "db/MATL", "{""Assign"":{" & MatJ(2, clsConc, denConc) & "," & MatJ(3, clsConc & " (W=0)", 0) & "}}"
    If piled Then
        Add "db/SECT", "{""Assign"":{""1"":{""SECTTYPE"":""DBUSER"",""SECT_NAME"":""KAZIK"",""SECT_BEFORE"":{""OFFSET_PT"":""CC"",""OFFSET_CENTER"":0,""USER_OFFSET_REF"":0,""HORZ_OFFSET_OPT"":0,""USERDEF_OFFSET_YI"":0," & _
            """VERT_OFFSET_OPT"":0,""USERDEF_OFFSET_ZI"":0,""USE_SHEAR_DEFORM"":true,""USE_WARPING_EFFECT"":false,""SHAPE"":""SR"",""DATATYPE"":2,""SECT_I"":{""vSIZE"":[" & JN(pDia) & ",0,0,0,0,0,0,0,0,0]}}}}}"
        Add "db/CO_S", "{""Assign"":{" & ColJ("1", PILE_COLOR) & "}}"
    End If
    Add "db/THIK", "{""Assign"":{" & ThkJ(1, tFt, 0) & "," & ThkJ(2, tWg, 0) & "," & ThkJ(3, tBw, (tSt - tBw) / 2#) & "," & ThkJ(4, tSt, 0) & IIf(earOn, "," & ThkJ(5, tEr, offEar), "") & "}}"
    nThkW = IIf(earOn, 5, 4)
    nOffW = IIf(Abs((tSt - tBw) / 2#) > EPS, 1, 0) + IIf(earOn And Abs(offEar) > EPS, 1, 0)
    ' view colours by member as the templates' *THIK-COLOR (owner 2026-11-10); db/CO_T is GET / PUT only
    Add "db/CO_T", "{""Assign"":{" & ColJ("1", ThkColor(1)) & "," & ColJ("2", ThkColor(2)) & "," & ColJ("3", ThkColor(3)) & "," & ColJ("4", ThkColor(4)) & IIf(earOn, "," & ColJ("5", ThkColor(5)), "") & "}}"
    ' nodes
    s = "": cnt = 0
    For i = 1 To nCnt
        s = s & IIf(s = "", "", ",") & QS(CStr(nId(i))) & ":{""X"":" & JN(nX(i)) & ",""Y"":" & JN(nY(i)) & ",""Z"":" & JN(nZ(i)) & "}"
        cnt = cnt + 1
        If cnt = 4 * CHUNK Then Add "db/NODE", "{""Assign"":{" & s & "}}": s = "": cnt = 0
    Next i
    If s <> "" Then Add "db/NODE", "{""Assign"":{" & s & "}}"
    ' elements
    s = "": cnt = 0
    For i = 1 To eCnt
        If eKind(i) = 1 Then
            s = s & IIf(s = "", "", ",") & QS(CStr(eId(i))) & ":{""TYPE"":""PLATE"",""MATL"":2,""SECT"":" & eThk(i) & ",""NODE"":[" & eNds(i, 1) & "," & eNds(i, 2) & "," & eNds(i, 3) & "," & eNds(i, 4) & ",0,0,0,0],""ANGLE"":0,""STYPE"":3}"
        Else
            s = s & IIf(s = "", "", ",") & QS(CStr(eId(i))) & ":{""TYPE"":""BEAM"",""MATL"":3,""SECT"":1,""NODE"":[" & eNds(i, 1) & "," & eNds(i, 2) & ",0,0,0,0,0,0],""ANGLE"":0,""STYPE"":0}"
        End If
        cnt = cnt + 1
        If cnt = 4 * CHUNK Then Add "db/ELEM", "{""Assign"":{" & s & "}}": s = "": cnt = 0
    Next i
    If s <> "" Then Add "db/ELEM", "{""Assign"":{" & s & "}}"
    Add "db/GRUP", "{""Assign"":{" & GrupJ() & "}}"
    Add "db/BNGR", "{""Assign"":{""1"":{""NAME"":""Sismik"",""AUTOTYPE"":0},""2"":{""NAME"":""Statik"",""AUTOTYPE"":0}}}"
    If piled Then
        Add "db/NSPR", "{""Assign"":{" & springJson & "}}"
        s = ""
        For i = 0 To nP - 1
            s = s & IIf(s = "", "", ",") & QS(CStr(pileMaster(i))) & ":{""ITEMS"":[{""ID"":1,""GROUP_NAME"":""Statik"",""DOF"":111111,""S_NODE"":[" & pileTop(i) & "]}]}"
        Next i
        Add "db/RIGD", "{""Assign"":{" & s & "}}"
        Add "db/SECF", "{""Assign"":{""1"":{""ITEMS"":[{""ID"":1,""GROUP_NAME"":""Sismik"",""AREA_SF"":1,""ASY_SF"":1,""ASZ_SF"":1,""IXX_SF"":1,""IYY_SF"":" & JN(fPile) & _
            ",""IZZ_SF"":" & JN(fPile) & ",""WGT_SF"":1,""W_SF"":1,""IPART"":1,""bDiffIJ"":false}]}}}"
    End If
    ' plate stiffness factors in group Sismik (owner 2026-11-10, TKDY 4.3.7.6): every WALL plate (stem, backwall,
    ' wing and ear walls; a junction plate by its member) all 8 factors = KA_RIJITLIK_PLAK; footing plates none
    s = "": cnt = 0
    For i = 1 To nPlate
        If eMem(i) <> M_FOOT Then
            nPssfW = nPssfW + 1: pssfMem(eMem(i)) = pssfMem(eMem(i)) + 1
            s = s & IIf(s = "", "", ",") & QS(CStr(eId(i))) & ":{""ITEMS"":[{""ID"":1,""GROUP_NAME"":""Sismik"",""AXIAL_X"":" & JN(fPlate) & ",""AXIAL_Y"":" & JN(fPlate) & ",""SHEAR"":" & JN(fPlate) & _
                ",""OUT_BENDING_X"":" & JN(fPlate) & ",""OUT_BENDING_Y"":" & JN(fPlate) & ",""OUT_TORSION"":" & JN(fPlate) & ",""OUT_SHEAR_X"":" & JN(fPlate) & ",""OUT_SHEAR_Y"":" & JN(fPlate) & "}]}"
            cnt = cnt + 1
            If cnt = 4 * CHUNK Then Add "db/PSSF", "{""Assign"":{" & s & "}}": s = "": cnt = 0
        End If
    Next i
    If s <> "" Then Add "db/PSSF", "{""Assign"":{" & s & "}}"
    If Not piled Then
        s = ""
        For i = 1 To nPlate
            If eMem(i) = M_FOOT Then s = s & IIf(s = "", "", ",") & eId(i)
        Next i
        AddPost "ope/SSPS", "{""Argument"":{""CONVERT_TO"":""POINT_SPRING"",""GROUP_NAME"":""Statik"",""NODE_ELEMS"":{""KEYS"":[" & s & "]},""ELEMENT"":{""TYPE"":""PLANAR""},""BOUNDARY"":{""TYPE"":""COMP"",""DIR"":1,""SUBGRADE"":" & JN(kvFoot) & "}}}"
        If khFoot > 0 Then AddPost "ope/SSPS", "{""Argument"":{""CONVERT_TO"":""POINT_SPRING"",""GROUP_NAME"":""Statik"",""NODE_ELEMS"":{""KEYS"":[" & s & "]},""ELEMENT"":{""TYPE"":""PLANAR""},""BOUNDARY"":{""TYPE"":""LINEAR"",""STIFF"":[" & JN(khFoot) & "," & JN(khFoot) & ",0],""bDAMP"":false,""DAMP"":[0,0,0]}}}"
    End If
    ' load groups: one per component (owner 2026-11-10)
    s = ""
    For k = 1 To N_CASE
        s = s & IIf(k = 1, "", ",") & QS(CStr(k)) & ":{""NAME"":" & QS(GrpOf(k)) & "}"
    Next k
    Add "db/LDGR", "{""Assign"":{" & s & "}}"
    ' load cases
    Set mc = ModelCases()
    s = ""
    For m = 1 To mc.Count
        s = s & IIf(m = 1, "", ",") & QS(CStr(m)) & ":{""NO"":" & m & ",""NAME"":" & QS(CStr(Split(mc(m), "|")(0))) & ",""TYPE"":" & QS(CStr(Split(mc(m), "|")(1))) & ",""DESC"":""""}"
    Next m
    Add "db/STLD", "{""Assign"":{" & s & "}}"
    ' self weight and its seismic inertia: ONE record per factored case, the combined vector in load
    ' group DC (owner 2026-11-10, option B: two records in one case make doc/ANAL fail silently, T6)
    s = "": j = 0
    For m = 1 To mc.Count
        ax = 0: ay = 0: az = 0
        For k = 1 To N_CASE
            f = FactorOf(m, k)
            If f <> 0 Then
                BodyF caseNm(k), fx, fy, fz
                ax = ax + f * fx: ay = ay + f * fy: az = az + f * fz
            End If
        Next k
        apF(m, 1) = apF(m, 1) + ax * wPlate: apF(m, 2) = apF(m, 2) + ay * wPlate: apF(m, 3) = apF(m, 3) + az * wPlate
        If Abs(ax) + Abs(ay) + Abs(az) > EPS Then
            j = j + 1
            s = s & IIf(s = "", "", ",") & QS(CStr(j)) & ":{""LCNAME"":" & QS(CStr(Split(mc(m), "|")(0))) & ",""GROUP_NAME"":""DC"",""FV"":[" & JN(ax) & "," & JN(ay) & "," & JN(az) & "]}"
        End If
    Next m
    nBodfW = j
    Add "db/BODF", "{""Assign"":{" & s & "}}"
    ' seat loads
    s = ""
    For i = 1 To nSeat
        it = "": nIt = 0
        For m = 1 To mc.Count
            ax = 0: ay = 0: az = 0
            For k = 1 To N_CASE
                f = FactorOf(m, k)
                If f <> 0 Then
                    SeatF caseNm(k), fx, fy, fz
                    apF(m, 1) = apF(m, 1) + f * fx: apF(m, 2) = apF(m, 2) + f * fy: apF(m, 3) = apF(m, 3) + f * fz
                    If SPLIT_GROUPS Then
                        If Abs(fx) + Abs(fy) + Abs(fz) > EPS Then
                            nIt = nIt + 1
                            it = it & IIf(it = "", "", ",") & CnldItem(nIt, CStr(Split(mc(m), "|")(0)), GrpOf(k), f * fx, f * fy, f * fz)
                        End If
                    Else
                        ax = ax + f * fx: ay = ay + f * fy: az = az + f * fz
                    End If
                End If
            Next k
            If Not SPLIT_GROUPS And Abs(ax) + Abs(ay) + Abs(az) > EPS Then
                nIt = nIt + 1
                it = it & IIf(it = "", "", ",") & CnldItem(nIt, CStr(Split(mc(m), "|")(0)), "", ax, ay, az)
            End If
        Next m
        If it <> "" Then s = s & IIf(s = "", "", ",") & QS(CStr(seatN(i))) & ":{""ITEMS"":[" & it & "]}": nCnldN = nCnldN + 1: nCnldI = nCnldI + nIt
    Next i
    If s <> "" Then Add "db/CNLD", "{""Assign"":{" & s & "}}"
    ' pressures
    s = "": cnt = 0
    For i = 1 To nPlate
        it = "": nIt = 0
        PlateGeo i, ar, ux, uy, uz
        For m = 1 To mc.Count
            For j = 1 To 4: acc(j) = 0: Next j
            hit = False
            For k = 1 To N_CASE
                f = FactorOf(m, k)
                If f <> 0 Then
                    If ElemP(caseNm(k), i, p) Then
                        If SPLIT_GROUPS Then
                            For j = 1 To 4: p(j) = f * p(j): Next j
                            AddPres m, p, eNn(i), ar, ux, uy, uz
                            nIt = nIt + 1
                            it = it & IIf(it = "", "", ",") & PresItem(nIt, CStr(Split(mc(m), "|")(0)), GrpOf(k), p, eNn(i))
                        Else
                            hit = True
                            For j = 1 To 4: acc(j) = acc(j) + f * p(j): Next j
                        End If
                    End If
                End If
            Next k
            If hit Then
                AddPres m, acc, eNn(i), ar, ux, uy, uz
                nIt = nIt + 1
                it = it & IIf(it = "", "", ",") & PresItem(nIt, CStr(Split(mc(m), "|")(0)), "", acc, eNn(i))
            End If
        Next m
        If it <> "" Then
            s = s & IIf(s = "", "", ",") & QS(CStr(eId(i))) & ":{""ITEMS"":[" & it & "]}"
            nPresE = nPresE + 1: nPresI = nPresI + nIt
            cnt = cnt + 1
            If cnt = CHUNK Then Add "db/PRES", "{""Assign"":{" & s & "}}": s = "": cnt = 0
        End If
    Next i
    If s <> "" Then Add "db/PRES", "{""Assign"":{" & s & "}}"
    AddCombos mc
    Add "db/BCCT", "{""Assign"":{""1"":{""bSPT"":true,""bSPR"":true,""bGSPR"":true,""bCGLINK"":true,""bSSSF"":true,""bPSSF"":true,""bRLS"":true,""bCDOF"":false," & _
        """vBOUNDARY"":[{""BGCNAME"":""ST"",""vBG"":[""Statik""]},{""BGCNAME"":""EQ"",""vBG"":[""Sismik"",""Statik""]}],""vLOADANAL"":[" & BcctRows(mc) & "]}}}"
    If piled Then
        Add "db/EIGV", "{""Assign"":{""1"":{""TYPE"":""LANCZOS"",""iFREQ"":100,""bMINMAX"":false,""bSTRUM"":false}}}"
    Else
        AddSoft "db/ACTL-M1", "{""Assign"":{""1"":{""ARCD"":true,""ANRC"":true,""CSECF"":false,""TRS"":true,""BMSTRESS"":false,""CRBAR"":false,""CLATS"":false,""CLFORM"":true,""CABINIT"":false,""BSCHG"":""CHANGE""," & _
            """TCELEM"":{""NUMINC"":5,""INTOUT"":""LAST"",""CONVERGENCE"":{""DISPL"":{""OPT_USE"":true,""VALUE"":0.0001},""LOAD"":{""OPT_USE"":false},""WORK"":{""OPT_USE"":false}}}}}}"
    End If
End Sub

Private Function MatJ(ByVal no As Long, ByVal nm As String, ByVal den As Double) As String
    MatJ = QS(CStr(no)) & ":{""TYPE"":""CONC"",""NAME"":" & QS(nm) & ",""HE_SPEC"":0,""HE_COND"":0,""bMASS_DENS"":false,""DAMP_RAT"":0.05,""PARAM"":[{""P_TYPE"":2,""ELAST"":" & JN(eConc) & _
        ",""POISN"":0.2,""THERMAL"":0.00001,""DEN"":" & JN(den) & "}]}"
End Function

Private Function ThkJ(ByVal no As Long, ByVal t As Double, ByVal off As Double) As String
    ThkJ = QS(CStr(no)) & ":{""NAME"":" & QS(CStr(no)) & ",""TYPE"":""VALUE"",""bINOUT"":false,""T_IN"":" & JN(t) & ",""T_OUT"":0,""OFFSET"":" & IIf(Abs(off) > EPS, 2, 0) & ",""O_VALUE"":" & JN(off) & "}"
End Function

' area and unit normal (local z) of plate e: node order n1 -> n2 -> n3 (-> n4) counter-clockwise about +z
Private Sub PlateGeo(ByVal e As Long, ByRef ar As Double, ByRef ux As Double, ByRef uy As Double, ByRef uz As Double)
    Dim i1 As Long, i2 As Long, i3 As Long, i4 As Long, ax As Double, ay As Double, az As Double, bx As Double, by As Double, bz As Double, cx As Double, cy As Double, cz As Double, ln As Double
    i1 = NodeIdx(eNds(e, 1)): i2 = NodeIdx(eNds(e, 2)): i3 = NodeIdx(eNds(e, 3))
    If eNn(e) = 4 Then
        i4 = NodeIdx(eNds(e, 4))
        ax = nX(i3) - nX(i1): ay = nY(i3) - nY(i1): az = nZ(i3) - nZ(i1)
        bx = nX(i4) - nX(i2): by = nY(i4) - nY(i2): bz = nZ(i4) - nZ(i2)
    Else
        ax = nX(i2) - nX(i1): ay = nY(i2) - nY(i1): az = nZ(i2) - nZ(i1)
        bx = nX(i3) - nX(i1): by = nY(i3) - nY(i1): bz = nZ(i3) - nZ(i1)
    End If
    cx = ay * bz - az * by: cy = az * bx - ax * bz: cz = ax * by - ay * bx
    ln = Sqr(cx * cx + cy * cy + cz * cz)
    ar = 0.5 * ln: ux = 0: uy = 0: uz = 0
    If ln > 0 Then ux = cx / ln: uy = cy / ln: uz = cz / ln
End Sub

Private Function ThkOf(ByVal id As Long) As Double
    Select Case id
        Case 1: ThkOf = tFt
        Case 2: ThkOf = tWg
        Case 3: ThkOf = tBw
        Case 4: ThkOf = tSt
        Case 5: ThkOf = tEr
    End Select
End Function

' resultant of one pressure item (local z, nodal values p(1..nn), mean x area) into the case total
Private Sub AddPres(ByVal m As Long, p() As Double, ByVal nn As Long, ByVal ar As Double, ByVal ux As Double, ByVal uy As Double, ByVal uz As Double)
    Dim j As Long, pm As Double
    For j = 1 To nn: pm = pm + p(j): Next j
    pm = pm / nn
    apF(m, 1) = apF(m, 1) + pm * ar * ux: apF(m, 2) = apF(m, 2) + pm * ar * uy: apF(m, 3) = apF(m, 3) + pm * ar * uz
End Sub

' colours of thickness id (1 footing, 2 wing walls, 3 backwall, 4 stem, 5 ear walls): W rgb, HF rgb, HE rgb
Private Function ThkColor(ByVal id As Long) As String
    Select Case id
        Case 1: ThkColor = "106,67,0,255,192,87,106,67,0"
        Case 2: ThkColor = "0,40,125,160,192,255,0,40,125"
        Case 3: ThkColor = "87,87,0,192,192,0,87,87,0"
        Case 4: ThkColor = "70,70,0,192,255,0,70,70,0"
        Case 5: ThkColor = "159,0,40,255,87,128,159,0,40"
    End Select
End Function

' one db/CO_T or db/CO_S record (the schema spells the opacity switch bBLEMD)
Private Function ColJ(ByVal key As String, ByVal c As String) As String
    Dim v As Variant
    v = Split(c, ",")
    ColJ = QS(key) & ":{""W_R"":" & v(0) & ",""W_G"":" & v(1) & ",""W_B"":" & v(2) & ",""HF_R"":" & v(3) & ",""HF_G"":" & v(4) & ",""HF_B"":" & v(5) & _
        ",""HE_R"":" & v(6) & ",""HE_G"":" & v(7) & ",""HE_B"":" & v(8) & ",""bBLEMD"":false,""FACT"":0.5}"
End Function

' record "key" of a GET answer (spaces removed) holds the 9 colour numbers, in any field order
Private Function ColHas(ByVal resp As String, ByVal key As String, ByVal c As String) As Boolean
    Dim p As Long, q As Long, rec As String, v As Variant, f As Variant, i As Long
    p = InStr(resp, """" & key & """:{")
    If p = 0 Then Exit Function
    q = InStr(p, resp, "}")
    If q = 0 Then Exit Function
    rec = Mid$(resp, p, q - p + 1)
    v = Split(c, ",")
    f = Array("W_R", "W_G", "W_B", "HF_R", "HF_G", "HF_B", "HE_R", "HE_G", "HE_B")
    For i = 0 To 8
        If InStr(rec, """" & f(i) & """:" & v(i) & ",") = 0 And InStr(rec, """" & f(i) & """:" & v(i) & "}") = 0 Then Exit Function
    Next i
    ColHas = True
End Function

Private Function GrupJ() As String
    Dim nmA As Variant, g As Long, i As Long, j As Long, eL As String, nL As String, d As Object, s As String, inG As Boolean, k As Long, kv As Variant
    nmA = Array("KAZIK", "TEMEL", "PERDE", "RICAT", "ALIN", "RJT", "KAZIK_RJT", "KULAK")
    For g = 0 To UBound(nmA)
        eL = "": Set d = CreateObject("Scripting.Dictionary")
        For i = 1 To eCnt
            Select Case nmA(g)
                Case "KAZIK": inG = (eKind(i) = 2)
                Case "TEMEL": inG = (eMem(i) = M_FOOT)
                Case "PERDE": inG = (eMem(i) = M_STEM)
                Case "RICAT": inG = (eMem(i) = M_WINGL Or eMem(i) = M_WINGR)
                Case "ALIN": inG = (eMem(i) = M_BACK)
                Case "RJT": inG = (eKind(i) = 1 And eRjt(i))
                Case "KAZIK_RJT": inG = (eKind(i) = 2) Or (eKind(i) = 1 And ePrj(i))
                Case "KULAK": inG = (eMem(i) = M_EARL Or eMem(i) = M_EARR)
            End Select
            If inG Then
                eL = eL & IIf(eL = "", "", ",") & eId(i)
                For j = 1 To IIf(eKind(i) = 1, eNn(i), 2)
                    d(eNds(i, j)) = 1
                Next j
            End If
        Next i
        If nmA(g) = "KAZIK" And piled Then
            For i = 0 To nP - 1: d(pileMaster(i)) = 1: Next i
        End If
        If eL <> "" Then
            nL = ""
            For Each kv In SortedLongKeys(d)
                nL = nL & IIf(nL = "", "", ",") & kv
            Next kv
            k = k + 1
            s = s & IIf(s = "", "", ",") & QS(CStr(k)) & ":{""NAME"":" & QS(CStr(nmA(g))) & ",""P_TYPE"":0,""N_LIST"":[" & nL & "],""E_LIST"":[" & eL & "]}"
        End If
    Next g
    nGrupW = k
    GrupJ = s
End Function

' LCOM-GEN (owner M22): no combination per factored case, only the 'Envelope' rows tier by tier (an
' envelope used by another goes first); an item that names a factored case is a load case (ST)
Private Sub AddCombos(mc As Collection)
    Dim t As String, i As Long, tier As Long, maxT As Long, p As Variant, q As Variant, its As String, nmb As Long, j As Long, an As String
    For i = 1 To nCmb
        If cmbTier(i) > maxT Then maxT = cmbTier(i)
    Next i
    For tier = 1 To maxT
        t = ""
        For i = 1 To nCmb
            If cmbKind(i) = 1 And cmbTier(i) = tier And cmbEff(i) Then
                its = ""
                For Each p In Split(cmbItems(i), ";")
                    If Len(p) > 0 Then
                        q = Split(p, "|")
                        j = CmbIndex(CStr(q(1)))
                        If j > 0 Then
                            If cmbEff(j) Then
                                an = IIf(cmbKind(j) = 0, "ST", "CB")
                                its = its & IIf(its = "", "", ",") & "{""ANAL"":""" & an & """,""LCNAME"":" & QS(CStr(q(1))) & ",""FACTOR"":" & q(2) & "}"
                            End If
                        End If
                    End If
                Next p
                nmb = nmb + 1
                t = t & IIf(t = "", "", ",") & QS(CStr(nmb)) & ":{""NO"":" & nmb & ",""NAME"":" & QS(cmbName(i)) & ",""ACTIVE"":""ACTIVE"",""bCB"":true,""iTYPE"":1,""DESC"":"""",""vCOMB"":[" & its & "]}"
            End If
        Next i
        If t <> "" Then Add "db/LCOM-GEN", "{""Assign"":{" & t & "}}"
    Next tier
End Sub

Private Function CombCount() As Long
    Dim i As Long
    For i = 1 To nCmb
        If cmbKind(i) = 1 And cmbEff(i) Then CombCount = CombCount + 1
    Next i
End Function

Private Function BcctRows(mc As Collection) As String
    Dim s As String, m As Long, nm As String
    For m = 1 To mc.Count
        nm = CStr(Split(mc(m), "|")(0))
        s = s & IIf(s = "", "", ",") & "{""TYPE"":""ST"",""BGCNAME"":""" & IIf(IsSeismicCase(nm), "EQ", "ST") & """,""LCNAME"":" & QS(nm) & "}"
    Next m
    BcctRows = s & ",{""TYPE"":""MV"",""BGCNAME"":""ST""},{""TYPE"":""SM"",""BGCNAME"":""ST""},{""TYPE"":""ULAT"",""BGCNAME"":""EQ""},{""TYPE"":""THRSEV"",""BGCNAME"":""EQ""}," & _
        "{""TYPE"":""THNS"",""BGCNAME"":""UNCHANGED""},{""TYPE"":""PO"",""BGCNAME"":""UNCHANGED""}"
End Function

Private Sub Add(ByVal ep As String, ByVal body As String)
    payloads.Add Array("PUT", ep, body, False)
End Sub

Private Sub AddPost(ByVal ep As String, ByVal body As String)
    payloads.Add Array("POST", ep, body, False)
End Sub

' a failure is reported as a warning, the build goes on
Private Sub AddSoft(ByVal ep As String, ByVal body As String)
    payloads.Add Array("PUT", ep, body, True)
End Sub

Private Sub WritePayloads(ByVal dirPath As String)
    Dim i As Long, p As Variant
    MakeDir dirPath
    For i = 1 To payloads.Count
        p = payloads(i)
        WriteText dirPath & "\" & Format(i, "00") & "_" & Replace(Replace(p(1), "db/", ""), "/", "_") & ".json", CStr(p(2))
    Next i
    WriteText dirPath & "\summary.txt", Summary()
End Sub

' ============================================================================
'  send and check
' ============================================================================
Private Function SendAll(Optional ByVal saveModel As Boolean = True) As String
    Dim rep As String, i As Long, p As Variant, resp As String, t0 As Single, cleanRep As String, n As Long, softRep As String
    t0 = Timer
    If Not GetKey() Then SendAll = "FAIL: no MAPI key.": Exit Function
    McCloseHttp
    If Not Api("GET", "db/NODE", "", resp) Then SendAll = "FAIL: Civil NX does not answer (db/NODE): " & Left$(resp, 300): Exit Function
    If Not ModelOk(rep) Then SendAll = rep: Exit Function
    rep = ""
    If ModelIsEmpty() Then
        ProgReset 1 + payloads.Count + 6
        cleanRep = "Clean pass skipped: the open model is empty."
    Else
        ProgReset 2 * (UBound(CleanTables()) + 1) + 1 + payloads.Count + 6
        cleanRep = CleanModel()
    End If
    If Left$(cleanRep, 7) = "STOPPED" Then SendAll = cleanRep & vbCrLf & "Nothing was written.": Exit Function
    For i = 1 To payloads.Count
        p = payloads(i)
        Prog "table " & i & "/" & payloads.Count & "  " & p(1) & "  " & Format$(Len(CStr(p(2))) / 1024#, "0") & " kB"
        If Not Api(CStr(p(0)), CStr(p(1)), CStr(p(2)), resp) Then
            If p(3) Then
                softRep = softRep & "  " & p(1) & " not accepted (" & Left$(resp, 200) & ")" & vbCrLf
            Else
                SendAll = cleanRep & vbCrLf & "FAIL at step " & i & " (" & p(1) & "):" & vbCrLf & Left$(resp, 600) & vbCrLf & vbCrLf & "Model is incomplete; nothing saved."
                Exit Function
            End If
        End If
    Next i
    rep = ReadBack() & SeatWrite()
    If softRep <> "" Then Warn "Optional tables not accepted:" & vbCrLf & softRep
    Prog "save"
    If InStr(rep, "FAIL") > 0 Then
        rep = rep & "NOT SAVED: the read-back failed." & vbCrLf
    ElseIf Not saveModel Then
        rep = rep & "Not saved (saveModel = False)." & vbCrLf
    ElseIf Api("POST", "doc/SAVE", "{""Argument"":{}}", resp) Then
        rep = rep & "Saved (open model file)." & vbCrLf
    Else
        rep = rep & "FAIL save: " & Left$(resp, 300) & vbCrLf
    End If
    If warnText <> "" Then rep = rep & "WARNINGS:" & vbCrLf & warnText
    SendAll = cleanRep & vbCrLf & "All " & payloads.Count & " tables sent (" & Format(Timer - t0, "0") & " s)." & vbCrLf & rep & vbCrLf & SkippedText() & AppliedTable() & vbCrLf & OpenItems() & vbCrLf & SCRIPT_ID & " " & SCRIPT_VERSION & ", " & bridgeName & " " & abutNm
End Function

' read back every table written and compare counts and key values; any FAIL = the model is not saved
Private Function ReadBack() As String
    Dim r As String, resp As String, i As Long, miss As String, dup As String, bad As String, fv As Variant
    r = "Read-back:" & vbCrLf
    Prog "read-back nodes"
    If RbGet("db/NODE", resp, r) Then r = r & RbN("nodes", CountOf(resp, """X"":"), nCnt)
    Prog "read-back elements"
    If RbGet("db/ELEM", resp, r) Then r = r & RbN("elements", CountOf(resp, """TYPE"":"), eCnt) & RbN("plates", CountOf(resp, """TYPE"":""PLATE"""), nPlate)
    Prog "read-back thickness"
    If RbGet("db/THIK", resp, r) Then
        r = r & RbN("thickness records", CountOf(resp, """T_IN"":"), nThkW) & RbN("thicknesses with an offset", CountOf(resp, """OFFSET"":2"), nOffW)
        r = r & "  thickness offsets: backwall " & IIf(Abs((tSt - tBw) / 2#) < EPS Or InStr(resp, """O_VALUE"":" & JN((tSt - tBw) / 2#)) > 0, "OK", "FAIL") & _
            IIf(earOn, ", ear walls " & IIf(Abs(offEar) < EPS Or InStr(resp, """O_VALUE"":" & JN(offEar)) > 0, "OK", "FAIL"), ", no ear walls") & vbCrLf
    End If
    Prog "read-back groups"
    If RbGet("db/GRUP", resp, r) Then r = r & RbN("structure groups", CountOf(resp, """E_LIST"":"), nGrupW)
    If RbGet("db/BNGR", resp, r) Then r = r & "  boundary groups Sismik / Statik " & IIf(InStr(resp, """Sismik""") > 0 And InStr(resp, """Statik""") > 0, "OK", "FAIL") & vbCrLf
    Prog "read-back supports"
    r = r & SpringCheck()
    If piled Then
        If RbGet("db/RIGD", resp, r) Then r = r & RbN("rigid links", CountOf(resp, """S_NODE"":"), nP)
        If RbGet("db/SECF", resp, r) Then r = r & RbN("pile stiffness factor records", CountOf(resp, """IYY_SF"":"), 1)
        If RbGet("db/EIGV", resp, r) Then r = r & RbN("eigen control records", CountOf(resp, """iFREQ"":"), 1)
    ElseIf Api("GET", "db/ACTL-M1", "", resp) Then
        r = r & "  nonlinear control NUMINC 5: " & IIf(InStr(Replace(resp, " ", ""), """NUMINC"":5") > 0, "OK", "not found (warning: Civil NX default increments)") & vbCrLf
    End If
    If RbGet("db/PSSF", resp, r) Then
        r = r & RbN("plate stiffness factor records (stem " & pssfMem(M_STEM) & ", backwall " & pssfMem(M_BACK) & ", wing walls " & pssfMem(M_WINGL) + pssfMem(M_WINGR) & _
            ", ear walls " & pssfMem(M_EARL) + pssfMem(M_EARR) & ", footing none)", CountOf(resp, """OUT_BENDING_X"":"), nPssfW)
        bad = ""
        For Each fv In Array("AXIAL_X", "AXIAL_Y", "SHEAR", "OUT_BENDING_X", "OUT_BENDING_Y", "OUT_TORSION", "OUT_SHEAR_X", "OUT_SHEAR_Y")
            If CountOf(resp, """" & fv & """:" & JN(fPlate)) <> nPssfW Then bad = bad & " " & fv
        Next fv
        r = r & "  plate stiffness factors " & JN(fPlate) & " on all 8 components: " & IIf(bad = "", "OK", "differ for" & bad & "  FAIL") & vbCrLf
    End If
    Prog "read-back colours"
    If RbGet("db/CO_T", resp, r) Then
        bad = ""
        For i = 1 To nThkW
            If Not ColHas(resp, CStr(i), ThkColor(i)) Then bad = bad & " " & i
        Next i
        r = r & "  thickness colours (" & nThkW & "): " & IIf(bad = "", "OK", "differ for thickness" & bad & "  FAIL") & vbCrLf
    End If
    If piled Then
        If RbGet("db/CO_S", resp, r) Then r = r & "  pile section colour: " & IIf(ColHas(resp, "1", PILE_COLOR), "OK", "FAIL") & vbCrLf
    End If
    Prog "read-back load groups and cases"
    If RbGet("db/LDGR", resp, r) Then
        miss = ""
        For i = 1 To N_CASE
            If InStr(resp, """NAME"":" & QS(GrpOf(i))) = 0 Then miss = miss & " " & GrpOf(i)
        Next i
        r = r & RbN("load groups", CountOf(resp, """NAME"":"), N_CASE) & IIf(miss = "", "", "  load groups missing:" & miss & "  FAIL" & vbCrLf)
    End If
    If RbGet("db/STLD", resp, r) Then
        miss = ""
        For i = 1 To nFac
            If InStr(resp, """NAME"":" & QS(cmbName(facIdx(i)))) = 0 Then miss = miss & " " & cmbName(facIdx(i))
        Next i
        r = r & RbN("load cases", CountOf(resp, """NAME"":"), nFac) & IIf(miss = "", "", "  load cases missing:" & miss & "  FAIL" & vbCrLf)
    End If
    Prog "read-back loads"
    If RbGet("db/BODF", resp, r) Then
        dup = ""
        For i = 1 To nFac
            If CountOf(resp, """LCNAME"":" & QS(cmbName(facIdx(i)))) > 1 Then dup = dup & " " & cmbName(facIdx(i))
        Next i
        r = r & RbN("self-weight records", CountOf(resp, """FV"":"), nBodfW) & IIf(dup = "", "", "  more than one self-weight record in:" & dup & "  FAIL" & vbCrLf)
    End If
    If RbGet("db/CNLD", resp, r) Then
        r = r & RbN("seat-load nodes", CountOf(resp, """ITEMS"":"), nCnldN) & RbN("seat-load items", CountOf(resp, """FX"":"), nCnldI)
        If SPLIT_GROUPS Then r = r & RbN("seat-load items with a load group", CountOf(resp, """GROUP_NAME"":") - CountOf(resp, """GROUP_NAME"":"""""), nCnldI)
    End If
    If RbGet("db/PRES", resp, r) Then
        r = r & RbN("loaded plates", CountOf(resp, """ITEMS"":"), nPresE) & RbN("pressure items", CountOf(resp, """CMD"":"), nPresI)
        If SPLIT_GROUPS Then r = r & RbN("pressure items with a load group", CountOf(resp, """GROUP_NAME"":") - CountOf(resp, """GROUP_NAME"":"""""), nPresI)
    End If
    Prog "read-back combinations"
    If RbGet("db/LCOM-GEN", resp, r) Then r = r & RbN("combinations (envelopes)", CountOf(resp, """iTYPE"":"), CombCount())
    If RbGet("db/BCCT", resp, r) Then
        r = r & RbN("cases on the cracked boundary set EQ", CountOf(resp, """TYPE"":""ST"",""BGCNAME"":""EQ"""), nSeisW) & _
            RbN("cases on the boundary set ST", CountOf(resp, """TYPE"":""ST"",""BGCNAME"":""ST"""), nFac - nSeisW) & _
            "  boundary change includes the plate stiffness factors (bPSSF): " & IIf(InStr(resp, """bPSSF"":true") > 0, "OK", "FAIL") & vbCrLf
    End If
    ReadBack = r
End Function

' GET a table with spaces removed and "\/" read as "/"; a FAIL line when it cannot be read
Private Function RbGet(ByVal ep As String, ByRef resp As String, ByRef r As String) As Boolean
    If Api("GET", ep, "", resp) Then
        resp = Replace(Replace(resp, " ", ""), "\/", "/")
        RbGet = True
    Else
        r = r & "  " & ep & " not readable  FAIL" & vbCrLf
    End If
End Function

Private Function RbN(ByVal what As String, ByVal n As Long, ByVal want As Long) As String
    RbN = "  " & what & " " & n & " / " & want & IIf(n = want, "", "  FAIL") & vbCrLf
End Function

' springs: piles = one record per pile node; spread = per footing node a COMP item pointing down
' (DIR 0 / 6 of ope/SSPS gives [0,0,1] = upside down, silently) and the total kv x footing area
Private Function SpringCheck() As String
    Dim resp As String, n As Long, nDown As Long, nUp As Long, tot As Double, want As Double, nFootN As Long, i As Long, d As Object, j As Long
    If Not Api("GET", "db/NSPR", "", resp) Then SpringCheck = "  springs: FAIL (db/NSPR not readable)" & vbCrLf: Exit Function
    If piled Then
        n = CountOf(resp, """ITEMS"":")
        SpringCheck = "  pile springs " & n & " / " & nSpring & IIf(n = nSpring, "", "  FAIL") & vbCrLf
        Exit Function
    End If
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To nPlate
        If eMem(i) = M_FOOT Then
            For j = 1 To eNn(i): d(eNds(i, j)) = 1: Next j
        End If
    Next i
    nFootN = d.Count
    nDown = CountOf(Replace(resp, " ", ""), """DV"":[0,0,-1]")
    nUp = CountOf(Replace(resp, " ", ""), """DV"":[0,0,1]")
    tot = SumAfter(resp, """STIFF"":")
    want = kvFoot * FootArea()
    SpringCheck = "  footing springs: " & nDown & " / " & nFootN & " compression springs pointing down" & IIf(nDown = nFootN And nUp = 0, "", "  FAIL") & vbCrLf & _
        "  sum of kv x area " & Format(tot, "0") & " / " & Format(want, "0") & " kN/m" & IIf(Abs(tot - want) <= 0.005 * want, "", "  FAIL") & vbCrLf
End Function

Private Function FootArea() As Double
    FootArea = fB * fLv
End Function

Private Function SumAfter(ByVal s As String, ByVal tag As String) As Double
    Dim p As Long
    p = InStr(1, s, tag)
    Do While p > 0
        SumAfter = SumAfter + Val(Mid$(s, p + Len(tag), 30))
        p = InStr(p + Len(tag), s, tag)
    Loop
End Function

Private Function ModelIsEmpty() As Boolean
    Dim t As Variant, resp As String
    For Each t In Array("NODE", "ELEM", "MATL", "STLD")
        If Not Api("GET", "db/" & t, "", resp) Then Exit Function
        If TableHasRecords(resp, CStr(t)) Then Exit Function
    Next t
    ModelIsEmpty = True
End Function

' do not wipe a bridge or superstructure model by mistake
Private Function ModelOk(ByRef msg As String) As Boolean
    Dim p As String, txt As String
    ModelOk = True
    If Not Api("GET", "db/PJCF", "", p) Then Exit Function
    If InStr(p, "midas-ustyapi-build") = 0 And InStr(p, "midas-bridge-build") = 0 Then Exit Function
    txt = "The model open in Civil NX was built by " & IIf(InStr(p, "midas-ustyapi-build") > 0, "MIDAS_USTYAPI", "MIDAS_KOPRU") & ", not by this macro."
    If askUser Then
        If UiMsg(txt & vbCrLf & vbCrLf & "WIPE it anyway?", vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) = vbYes Then Exit Function
    End If
    msg = "STOPPED: " & txt & " Nothing was changed. Open an empty file or an abutment model."
    ModelOk = False
End Function

Private Function CleanTables() As Variant
    CleanTables = Array("LCOM-CONC", "LCOM-GEN", "BCCT", "EIGV", "PRES", "CNLD", "BODF", "STLD", "LDGR", "PSSF", "SECF", "NSPR", "RIGD", "CONS", "BNGR", "GRUP", "ELEM", "NODE", "THIK", "SECT", "MATL")
End Function

Private Function CleanModel() As String
    Dim t As Variant, resp As String, remain As String, nTry As Long, leftTbl As Collection, keys As String
    Set leftTbl = New Collection
    For Each t In CleanTables(): leftTbl.Add CStr(t): Next t
    For nTry = 1 To 3
        If nTry > 1 Then
            For Each t In leftTbl
                Prog "clean pass " & nTry & "  delete by key " & t
                If Api("GET", "db/" & t, "", resp) Then
                    keys = TopKeys(resp)
                    If keys <> "" Then Call Api("DELETE", "db/" & t, "{""Argument"":[" & keys & "]}", resp)
                End If
            Next t
        Else
            For Each t In leftTbl
                Prog "clean pass 1  delete " & t
                Call Api("DELETE", "db/" & t, "{}", resp)
            Next t
        End If
        remain = ""
        Set leftTbl = New Collection
        For Each t In CleanTables()
            Prog "clean pass " & nTry & "  check " & t
            If Not Api("GET", "db/" & t, "", resp) Then
                remain = remain & "  " & t & " (cannot read: " & Left$(resp, 80) & ")" & vbCrLf
                leftTbl.Add CStr(t)
            ElseIf TableHasRecords(resp, CStr(t)) Then
                remain = remain & "  " & t & vbCrLf
                leftTbl.Add CStr(t)
            End If
        Next t
        If remain = "" Then Exit For
    Next nTry
    If remain = "" Then
        CleanModel = "Clean pass: OK (" & (UBound(CleanTables()) + 1) & " tables deleted and read back empty)."
    Else
        CleanModel = "STOPPED: the clean pass could not empty the model open in Civil NX. Still holding records:" & vbCrLf & remain & _
            "Delete them by hand in Civil NX (or open an empty model) and run again."
    End If
End Function

' ============================================================================
'  texts
' ============================================================================
Private Function Summary() As String
    Dim s As String, nPl As Long, i As Long
    For i = 1 To eCnt
        If eKind(i) = 1 Then nPl = nPl + 1
    Next i
    s = "Abutment " & abutNm & ", " & IIf(piled, "piled (KAZIKLI)", "spread footing (YUZEYSEL)") & ", skew " & JN(thetaDeg) & " deg, transverse loads toward " & IIf(trSign < 0, "-X", "+X") & vbCrLf & _
        "Footing " & JN(fB) & " x " & JN(fL) & " x " & JN(tFt) & " m; stem " & JN(tSt) & " x " & JN(hStem) & " m, backwall " & JN(tBw) & " x " & JN(hBack) & " m; wings " & JN(wOut) & " m out to out, " & JN(tWg) & " m" & vbCrLf & _
        "Heel " & JN(heelW) & " m, toe " & JN(Round(toeUsed, 3)) & " m (normal to the stem); " & IIf(earOn, "ear walls " & JN(lEar) & " m (" & JN(hRoot) & " -> " & JN(hTip) & " m, offset " & JN(offEar) & ")", "no ear walls (KA_KULAK_BOYU or KA_KULAK_KALINLIK = 0)") & vbCrLf & _
        "Mesh " & JN(meshM) & " m: " & nCnt & " nodes, " & nPl & " plates" & IIf(piled, ", " & eCnt - nPl & " pile elements", "") & vbCrLf & _
        "Seat: " & nSeat & " nodes, girders " & nGird & " (span type " & spanCol & ")" & vbCrLf
    If piled Then
        s = s & "Piles " & nPx & " x " & nPy & " D " & JN(pDia) & " m, spacing " & JN(sPx) & " / " & JN(sPy) & ", edge " & JN(Round(cPx, 3)) & " / " & JN(Round(cPy, 3)) & ", length " & JN(lPile) & " m in " & nSeg & " elements; springs kh x D x L (" & nLay & " layers), tip kv " & JN(kvTip) & vbCrLf
    Else
        s = s & "Surface springs kv " & JN(kvFoot) & " (compression only), kh " & JN(khFoot) & " kN/m3" & vbCrLf
    End If
    s = s & "Load cases: " & nFac & " factored cases (KA_KOMBINASYON 'Add' rows), " & IIf(SPLIT_GROUPS, "one load group per component", "components merged") & "; combinations: " & CombCount() & " envelopes (LCOM-GEN)" & vbCrLf
    If warnText <> "" Then s = s & "WARNINGS:" & vbCrLf & warnText
    Summary = s & SkippedText() & AppliedTable() & OpenItems()
End Function

' KA_KOMBINASYON rows skipped as Inactive (column B) and envelopes left without an active member
Private Function SkippedText() As String
    Dim t As String, u As String, i As Long
    For i = 1 To nCmb
        If Not cmbAct(i) Then
            t = t & IIf(t = "", "", ", ") & cmbName(i)
        ElseIf cmbKind(i) = 1 And Not cmbEff(i) Then
            u = u & IIf(u = "", "", ", ") & cmbName(i)
        End If
    Next i
    If t <> "" Then SkippedText = "Skipped as Inactive (KA_KOMBINASYON column B): " & t & vbCrLf
    If u <> "" Then SkippedText = SkippedText & "Envelopes not written (no active member): " & u & vbCrLf
End Function

' applied load totals per factored case (self weight with its kh inertia, seat loads, pressures); the
' macro does not run the analysis: Civil NX's reaction totals of each case should equal minus these
Private Function AppliedTable() As String
    Dim t As String, m As Long
    t = "Applied loads per load case (kN; plate weight " & JN(Round(wPlate, 1)) & " kN). The macro does not run the analysis: the reaction totals of Civil NX should equal minus these." & vbCrLf
    For m = 1 To nFac
        t = t & "  " & cmbName(facIdx(m)) & ": FX " & JN(Round(apF(m, 1), 1)) & ", FY " & JN(Round(apF(m, 2), 1)) & ", FZ " & JN(Round(apF(m, 3), 1)) & vbCrLf
    Next m
    AppliedTable = t
End Function

Private Function OpenItems() As String
    OpenItems = "Owner rules (2026-11-10, plan 4.1 M1-M23): earth pressure on the stem and backwall to the footing mid-plane, wing walls from the footing top; dynamic increment as an inverted triangle; EV on the heel only; " & _
        "longitudinal seat loads toward the span; the transverse direction moves the seat WS / TU-Y / PEQ-Y loads, the PEQ-Y inertia and the wing wall with the transverse soil pressure; no vertical seismic; seismic stiffness factors 0.5 on all 8 components of every wall plate, none on the footing, piles Iyy / Izz 0.5 (TKDY 4.3.7.6 / 7.2.4.4); " & _
        "skewed heel, toe, footing length and pile row spacing normal to the stem; combinations: envelopes only (LCOM-GEN). Still OPEN: case CF stays empty (KA_Y_CF is not applied)." & vbCrLf
End Function

Private Sub Warn(ByVal t As String)
    warnText = warnText & "- " & t & vbCrLf
End Sub

' ============================================================================
'  helpers
' ============================================================================
Private Function NV(ByVal nm As String) As Double
    NV = Num(NmRange(ThisWorkbook, nm).Cells(1, 1).Value, nm)
End Function

Private Function NS(ByVal nm As String) As String
    Dim v As Variant
    v = NmRange(ThisWorkbook, nm).Cells(1, 1).Value
    If IsError(v) Then Err.Raise vbObjectError + 21, , nm & ": error value."
    NS = CStr(v)
End Function

Private Function NSOr(ByVal nm As String) As String
    Dim v As Variant
    On Error Resume Next
    v = ThisWorkbook.Names(nm).RefersToRange.Cells(1, 1).Value
    If Err.Number = 0 And Not IsError(v) Then NSOr = CStr(v)
    On Error GoTo 0
End Function

Private Function HasName(ByVal nm As String) As Boolean
    Dim r As Range
    On Error Resume Next
    Set r = ThisWorkbook.Names(nm).RefersToRange
    HasName = Not (r Is Nothing)
    On Error GoTo 0
End Function

Private Function Fold(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, ChrW(220), "U"): s = Replace(s, ChrW(252), "U")
    s = Replace(s, ChrW(304), "I"): s = Replace(s, ChrW(305), "I")
    s = Replace(s, ChrW(350), "S"): s = Replace(s, ChrW(351), "S")
    s = Replace(s, ChrW(286), "G"): s = Replace(s, ChrW(287), "G")
    s = Replace(s, ChrW(199), "C"): s = Replace(s, ChrW(231), "C")
    s = Replace(s, ChrW(214), "O"): s = Replace(s, ChrW(246), "O")
    Fold = s
End Function

Private Function ColName(ByVal c As Long) As String
    ColName = Split(Cells(1, c).Address(True, False), "$")(0)
End Function

Private Function SortedLongKeys(d As Object) As Variant
    Dim a() As Long, i As Long, j As Long, t As Long, k As Variant
    If d.Count = 0 Then SortedLongKeys = Array(): Exit Function
    ReDim a(1 To d.Count)
    For Each k In d.Keys
        i = i + 1: a(i) = CLng(k)
    Next k
    For i = 2 To UBound(a)
        t = a(i): j = i - 1
        Do While j >= 1
            If a(j) <= t Then Exit Do
            a(j + 1) = a(j): j = j - 1
        Loop
        a(j + 1) = t
    Next i
    SortedLongKeys = a
End Function

Private Sub ProgReset(ByVal total As Long)
    progStep = 0
    progTotal = IIf(total < 1, 1, total)
End Sub

Private Sub Prog(ByVal stepName As String)
    Dim pct As Long, filled As Long
    progStep = progStep + 1
    If progTotal < 1 Then progTotal = 1
    pct = CLng((progStep - 1) * 100 / progTotal)
    If pct < 0 Then pct = 0
    If pct > 100 Then pct = 100
    filled = CLng(pct * PROG_BAR_WIDTH / 100)
    UiProgress pct, stepName
    On Error Resume Next
    Application.StatusBar = "MIDAS_KENARAYAK  [" & String(filled, ChrW(9608)) & String(PROG_BAR_WIDTH - filled, ChrW(9617)) & _
        "] " & pct & "%   (" & progStep & "/" & progTotal & ")  " & stepName
    On Error GoTo 0
    DoEvents
End Sub

Private Function Api(ByVal method As String, ByVal ep As String, ByVal body As String, ByRef resp As String) As Boolean
    Api = McApi(method, ep, body, resp)
    If McStatus() = 404 And Left$(ep, 4) = "doc/" Then McForgetKey
End Function

Private Function GetKey() As Boolean
    GetKey = McGetKey(ThisWorkbook, askUser, MSG_TITLE)
End Function

Private Function Pj(ByVal s As String) As String
    Dim i As Long, c As Long, o As String
    s = Fold(s)
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        If c >= 32 And c <= 126 Then o = o & ChrW(c) Else o = o & "?"
    Next i
    Pj = Left$(o, 80)
End Function
