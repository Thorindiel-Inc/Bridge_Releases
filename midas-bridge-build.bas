Attribute VB_Name = "MIDAS_KOPRU"
Option Explicit
' ============================================================================
'  MIDAS_KOPRU - builds the general whole-bridge model in MIDAS Civil NX through the MAPI
'  REST API (reference models/ALTYAPI.mct). Runs from the 1-3A workbook.
'  Step A (owner 2026-11-05): geometry - deck stick per span, deck-end slabs, bearings
'  on every girder, pier caps, piers, fixed bases and abutment bearing nodes.
'  Step B (owner 2026-11-06): loads (values from 1-3A, not from ALTYAPI), LM71 lanes, spectra,
'  TKDY stiffness factors, combinations of 1_GIRIS Tablo 14 / 15. No rail layer yet.
'  RTI = rail layer, track springs and stages, see the RTI BRANCH section.
'  Three models (owner 2026-11-06): STATIC = all loads, full stiffness, no mass / spectra.
'  SEISMIC = only what the response spectrum needs: DC, DW_1, DW_2, LL_KUTLE (80 kN/m static
'  train case), loads to masses, Ritz, spectra, EQ cases, TKDY / H-TKDY combinations; stiffness
'  factors with no boundary group (a group + db/BCCT did not reach the modes), so EVERY case
'  in that model runs at the reduced stiffness.
' ----------------------------------------------------------------------------
'  Button macro: MIDAS_KOPRU_OLUSTUR
'    0. asks STATIC, SEISMIC or RTI
'    1. reads sarkiklik (G_SARKIKLIK) and ky, kv (labels on the first 4<letter>_KDME sheet)
'    2. asks element length, bearing link length, fixed horizontal stiffness kh
'    3. reads 1-3A: span list, G_AYAK_TABLOSU (one row per axis), span table right of
'       G_SABIT_UC (fixed end SOL / SAG, Yatay Gabari, Kiris, Mesnet: left bearing of each
'       span, row 'Son' = last right), G_BASLIK_KESIT, G_BASLIK_BOYU, girder data.
'       One deck section per Yatay Gabari + Kiris combination. Joint spacing: B25, or the
'       cell of the same Hesap Acikligi in B25:G25 when filled.
'    4. summary; Yes = clean pass on the model open in Civil NX, MCT import of the
'       composite deck section (no JSON form; IMPORTMXT replaces the model, so it runs
'       first), all tables, read-back, doc/SAVE
'  Dry run (no dialogs, no HTTP): MIDAS_KOPRU_PAYLOAD outDir [, seismicModel]
'  Levels: deck node Z 0 (section offset CT), bearing top -(H + tc), bearing bottom = bearing seat
'  (cap top for a rectangle or when the bearings sit on the web of the inverted T, flange top when they
'  sit on its ledges); cap nodes at the cap centroid (offset CC), pier top = cap bottom, pier base =
'  pier top - table height (foundation top to cap bottom; owner 2026-11-09).
'  X: A1 bearing axis = 0; pier bearings at e from the pier axis, e = (axis spacing - Hesap Acikligi) / 2
'  (end spans: axis spacing - Hesap Acikligi); the two spans at a pier must agree (owner 2026-11-09).
'  Plan: G_AKS_XY (axis X / Y from the DXF, DXF_OKU; A1 first, span count + 1 rows, chords checked
'  against the span list) filled = each span a straight chord between its axes, caps on the
'  bisector of the two chords, piers and bearing links turned with them; model X = chord A1 to A2.
'  Empty table = straight. Level deck.
'  Numbers stay below 10500 so the rail layer (RTI) can be added later.
'  MAPI key: INPUT!T3, else MIDAS_MAPI_KEY, else asked (memory only).
' ============================================================================

Private Const SCRIPT_ID As String = "midas-bridge-build"
Private Const SCRIPT_VERSION As String = "2026-11-09f"
Private Const SCRIPT_CHANGELOG As String = "Macro window log shows the size of each table sent; Skew per span from the span table (column R picks G_VEREVLIK B22:F22): bearings, cap and pier turned; Seismic: TKDY 4.2.1.3 and 4.3.2.4 warnings; RTI: track springs along the track on curved spans (links rail to track), windows load exactly 300 m / 30 m (partial element loads), loaded springs under the whole second train, single track (G_GABARI_HAT_2 empty or 0): braking or traction, never both; rail as a rail-shaped H per track, groups for the result pull (rail and track nodes per track, deck ends at every axis), braking / traction loads on segments without a rail node now staged; Cap at its centroid, pier from the cap bottom, rectangle cap when B1 = B2 = 0, axis to bearing distance calculated, cap fit warnings; Static and seismic models use AASHTO Ec, RTI Eurocode Ecm"

Private Const API_BASE As String = "https://moa-engineers.midasit.com:443/civil/"
Private Const MSG_TITLE As String = "MIDAS_KOPRU"
Private Const KEY_SHEET As String = "INPUT"
Private Const KEY_CELL As String = "T3"
Private Const DEF_MESH As Double = 1.5          ' deck element length (m)
Private Const DEF_LINK_LEN As Double = 0.35     ' bearing link length (m)
Private Const DEF_K_FIX As Double = 100000#     ' fixed horizontal bearing stiffness kh (kN/m)
Private Const N_PIER_EL As Long = 5             ' at least this many equal elements per pier
Private Const MAX_PIER_EL As Double = 3#        ' and no element longer than this (m) (owner 2026-11-05)
Private Const GRAV_STYP As Double = 9.806
Private Const SEC_CAP As Long = 3
Private Const SEC_DECK0 As Long = 20          ' deck type d: section 1 for d = 1, else SEC_DECK0 + d
Private Const SEC_SLAB0 As Long = 40          ' slab of deck type d: section 2 for d = 1, else SEC_SLAB0 + d
Private Const SEC_PIER0 As Long = 60
Private Const SECT_FILL_TINT As Double = 0.45  ' 3D fill = section colour moved this far toward white
Private Const SECT_SHADE_STEP As Double = 0.18 ' each further deck or pier type this much darker
Private Const MAX_TYPES As Long = 20
Private Const PI As Double = 3.14159265358979
Private Const NUM_LIMIT As Long = 10500         ' rail layer numbers start here
Private Const SOLVE_H As Double = 0.05          ' deck section solver: coarse cell size (m); fine = half
' Step B constants (code values; the workbooks hold them as report text only)
Private Const RAIL_OVER_BALLAST As Double = 0.16 ' sleeper + rail above the ballast: 0.81 m rail level of RTI.mct with 0.60 + 0.05
Private Const WL_HEIGHT As Double = 4#          ' train wind above rail level (2_YUKLER C252)
Private Const CF_HEIGHT As Double = 1.8         ' centrifugal force above rail level (2_YUKLER C160)
Private Const BRAKE_Q As Double = 20#           ' kN/m x alpha (EN 1991-2 6.5.3, 2_YUKLER C150)
Private Const BRAKE_L As Double = 300#          ' m, longest braking length: windows of whole spans, see BrakeWindows
Private Const TRAC_Q As Double = 33#            ' kN/m x alpha (2_YUKLER C144)
Private Const TRAC_L As Double = 30#            ' m, traction length per span case
Private Const RTI_EMB As Double = 100#         ' embankment each side (UIC 774-3 minimum, owner 2026-11-06)
Private Const LM71_Q As Double = 80#            ' kN/m per track, LM71 uniform load (nonlinear live load stages)
Private Const LM71_AX As Double = 1000#         ' kN per track, four 250 kN axles
Private Const LM71_AXL As Double = 6.4          ' m, length the axles replace the uniform load on
Private Const TRK_U0 As Double = 0.002          ' m, ballast track law: displacement at the plateau
Private Const TRK_Q_UNL As Double = 12#         ' kN/m per track, unloaded, moderate maintenance (owner 2026-11-05)
Private Const TRK_Q_LOAD As Double = 60#        ' kN/m per track, loaded
Private Const RAIL_E As Double = 210000000#     ' kN/m2
Private Const RAIL_SIZE As String = "0.172,0.144,0.033,0.051,0.3,0.015,0,0,0,0" ' m: H, B1, tw, tf1, B2, tf2 of 2 x UIC 60 per track
Private Const SEC_RAIL As Long = 90
Private Const MAT_RAIL As Long = 9
Private Const EQ_FUNC As String = "DD2a"        ' KOS 1, KK level, elastic (owner 2026-11-06)
Private Const EQ_STIFF As Double = 0.5          ' TKDY effective bending stiffness: deck, cap, pier
Private Const EQ_STIFF_SLAB As Double = 0.25    ' TKDY: connection slab
Private Const RITZ_X As Long = 35, RITZ_Y As Long = 35, RITZ_Z As Long = 30   ' for now (owner 2026-11-06)
Private Const LINK_ON_SKEW As Boolean = False   ' bearing link axes along the girders on a skewed support (owner 2026-11-09); True = along the support line
Private Const MASS_CASE As String = "LL_KUTLE"  ' uniform train load x Tablo 14 TKDY LL factor: mass and TKDY combinations (owner 2026-11-06)
Private Const BMLD_CHUNK As Long = 60
Private Const Q_BAL As Long = 1, Q_KERB As Long = 2, Q_RAIL As Long = 3, Q_PANEL As Long = 4, Q_PROT As Long = 5, Q_TRK As Long = 6, Q_ENL As Long = 7

' ---- inputs ----
Private nG As Long, sG As Double, Hg As Double, Lc As Double, deckW As Double
Private tc As Double, eG As Double, eS As Double, eO As Double, denG As Double, denS As Double, clsG As String, clsS As String, clsO As String
Private sark As Double, kyB As Double, kEB As Double, jointL As Double, eAx As Double
Private capS(1 To 5) As Double, capL As Double, capRect As Boolean, capYc As Double, brgW As Double, brgL As Double
Private eLs() As Double, eRs() As Double, ePier() As Double, seatOnLedge() As Boolean
Private nSp As Long, spL() As Double, fixLeft() As Boolean
Private pH() As Double, pSec() As Long     ' pier 1..nSp-1: height, section number
Private nPS As Long, psTyp() As String, psEn() As Double, psBoy() As Double, psT() As Double
' deck types: one per Yatay Gabari + Kiris combination
Private nDk As Long, spType() As Long, tHes(1 To MAX_TYPES) As Long, tGab(1 To MAX_TYPES) As Long
Private tNG(1 To MAX_TYPES) As Long, tSG(1 To MAX_TYPES) As Double, tHg(1 To MAX_TYPES) As Double, tLc(1 To MAX_TYPES) As Double
Private tW(1 To MAX_TYPES) As Double, tJoint(1 To MAX_TYPES) As Double, tKb(1 To MAX_TYPES, 1 To 12) As Double, tName(1 To MAX_TYPES) As String
Private tJ(1 To MAX_TYPES) As Double, tAsy(1 To MAX_TYPES) As Double, tAsz(1 To MAX_TYPES) As Double
Private meshDx As Double, linkLen As Double, kFix As Double
Private kb(1 To 12) As Double, kbY(1 To 12) As Double
Private projName As String, clientName As String, bridgeName As String
Private warnText As String, askUser As Boolean, seisMode As Boolean, rtiMode As Boolean
Private rigMap As Object, nTrk As Long, rN(1 To 2) As Long, rNode() As Long, rGr() As Long, rSt() As Double, rTr() As Double, rAng() As Double, rOn() As Boolean
Private rFirst() As Long, rLast() As Long, xNo As Long, xX() As Double, xY() As Double, xZ() As Double, reCnt As Long, rE1() As Long, rE2() As Long, rET() As Long, rEI() As Long
Private rtiNL As Boolean, nLL As Long, lpA() As Long, lpB() As Long, lpX() As Long
Private nWin As Long, wNm() As String, wA() As Double, wB() As Double, wC() As Double, wD() As Double, wE() As Double, wF() As Double, wBr() As Boolean, wTr() As Boolean, wS() As Long, cut1() As Double, cut2() As Double
Private bgUsed As Object, lgUsed As Object, railItems As Object, fnIds As Object

' ---- model ----
Private nodeKey As Object
Private nX() As Double, nY() As Double, nZ() As Double, nCnt As Long
Private eMat() As Long, eSec() As Long, eN1() As Long, eN2() As Long, eCnt As Long
Private grpDeck As String, grpCap As String, grpPier As String
Private rigdJson As String, elnkJson As String, consJson As String, nLink As Long
Private capSlaves As Object, polyArea As Double, deckArea As Double
' plan: axis points, cap direction (unit), span chord direction (unit), chord length, heading (deg)
Private pxA() As Double, pyA() As Double, rxA() As Double, ryA() As Double, cxS() As Double, cyS() As Double, lenS() As Double, phiS() As Double, phiA() As Double
Private pierAng As Object, maxOff As Double, nAxXY As Long
Private gCz As Double, gIy As Double, gIz As Double, gPer As Double
Private dkLog As String
' Step B
Private tQ(1 To MAX_TYPES, 1 To 7) As Double, tKL(1 To MAX_TYPES) As Double, tKR(1 To MAX_TYPES) As Double, tH1(1 To MAX_TYPES) As Double, tH2(1 To MAX_TYPES) As Double, tPhi(1 To MAX_TYPES) As Double
Private wG As Double, wV As Double, alphaLM As Double, qsk As Double, tuPos As Double, tuNeg As Double, vDes As Double, qvk As Double, railH As Double
Private facRows As Object, dcMax As Double, dcMin As Double, dcMax4 As Double, dwMax As Double, dwMin As Double, eqMain As Double, eqOther As Double
Private specNames As Collection, specJson As Collection, phiByElem As Boolean
Private skS() As Double, skA() As Double   ' skew per span / per axis (deg), + = support line turned counter-clockwise in plan
Private spR() As Double, dkA() As Long, dkB() As Long, mnA() As Long, mnB() As Long, dxS() As Double
Private brgNode() As Long, brgSpan() As Long, nBrg As Long
Private nFW As Long, fwA() As Long, fwB() As Long
Private loadItems As Object, loadIds As Object, hasCF As Boolean, nCfSpans As Long, cfMax As Double, combNo As Long

Private payloads As Collection
Private mKey As String
Private http As Object
Private progStep As Long, progTotal As Long
Private Const PROG_BAR_WIDTH As Long = 24

Public Sub MIDAS_KOPRU_OLUSTUR()
    RunBuild "", False, False, False
End Sub

Public Sub MIDAS_KOPRU_PAYLOAD(ByVal outDir As String, Optional ByVal seismicModel As Boolean = False, Optional ByVal rtiModel As Boolean = False, Optional ByVal rtiNonlinear As Boolean = False)
    rtiNL = rtiNonlinear
    RunBuild outDir, True, seismicModel, rtiModel
End Sub

' scripted build, no dialogs (tests)
Public Sub MIDAS_KOPRU_BUILD(ByVal reportFile As String, ByVal apiKey As String, Optional ByVal saveModel As Boolean = True, Optional ByVal seismicModel As Boolean = False, Optional ByVal rtiModel As Boolean = False, Optional ByVal rtiNonlinear As Boolean = False)
    Dim rep As String
    On Error GoTo Fail
    warnText = ""
    seisMode = seismicModel: rtiMode = rtiModel: rtiNL = rtiNonlinear
    meshDx = DEF_MESH: linkLen = DEF_LINK_LEN: kFix = DEF_K_FIX
    PrepareModel
    mKey = apiKey
    askUser = False
    rep = SendAll(saveModel)
    WriteText reportFile, rep
    Application.StatusBar = False
    Exit Sub
Fail:
    WriteText reportFile, "FAIL: " & Err.Description
    Application.StatusBar = False
End Sub

Private Sub PrepareModel()
    Application.StatusBar = "MIDAS_KOPRU: reading inputs..."
    ReadInputs ThisWorkbook
    BuildModel
    If seisMode Then SeisChecks ThisWorkbook
    BuildLoads
    BuildPayloads
End Sub

Private Sub RunBuild(ByVal outDir As String, ByVal dry As Boolean, ByVal seismicModel As Boolean, ByVal rtiModel As Boolean)
    Dim msg As String, rep As String, v As Variant
    On Error GoTo Fail
    warnText = ""
    meshDx = DEF_MESH: linkLen = DEF_LINK_LEN: kFix = DEF_K_FIX
    askUser = Not dry
    seisMode = seismicModel: rtiMode = rtiModel
    If Not dry Then
        v = UiIdx("model", 1)
        seisMode = (v = 2): rtiMode = (v = 3)
        ' live load gate of the RTI model: 1 moving load (default), 2 moving load + nonlinear stages
        rtiNL = rtiMode And (UiIdx("ll", 1) = 2)
    End If
    If Not dry Then
        If Not AskModelValues() Then Exit Sub
    End If
    PrepareModel
    If dry Then
        WritePayloads outDir
        Application.StatusBar = False
        Exit Sub
    End If
    msg = FitBox(Summary(), "kopru_summary", 170) & vbCrLf & vbCrLf & _
          "YES = WIPE the open Civil NX model, rebuild it as the " & ModelName() & " model and SAVE it." & vbCrLf & _
          "NO = only write the JSON payloads (MIDAS\payload_kopru)."
    Select Case UiMsg(msg, vbYesNoCancel + vbQuestion, MSG_TITLE)
        Case vbNo
            WritePayloads ThisWorkbook.Path & "\MIDAS\payload_kopru"
            UiMsg "Payloads written to " & ThisWorkbook.Path & "\MIDAS\payload_kopru", vbInformation, MSG_TITLE
            Application.StatusBar = False
            Exit Sub
        Case vbCancel
            Application.StatusBar = False
            Exit Sub
    End Select
    rep = SendAll()
    Application.StatusBar = False
    UiMsg FitBox(rep, "kopru_report", 0), IIf(InStr(rep, "FAIL") > 0, vbExclamation, vbInformation), MSG_TITLE
    Exit Sub
Fail:
    Dim d As String: d = Err.Description
    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If dry Then
        MakeDir outDir
        WriteText outDir & "\ERROR.txt", d
    Else
        UiMsg "Stopped: " & d, vbCritical, MSG_TITLE
    End If
End Sub

Private Function AskModelValues() As Boolean
    meshDx = UiNum("mesh", DEF_MESH)
    linkLen = UiNum("link", DEF_LINK_LEN)
    kFix = UiNum("kfix", DEF_K_FIX)
    AskModelValues = True
End Function

Private Sub ReadInputs(wb As Workbook)
    Dim r As Range, gcol As Long, i As Long, v As Variant, n As Long, t As String
    projName = CStr(NmRange(wb, "G_PROJE_ADI").Cells(1).Value)
    clientName = CStr(NmRange(wb, "G_IDARE_ADI").Cells(1).Value)
    bridgeName = CStr(NmRange(wb, "G_KOPRU_ADI").Cells(1).Value)
    tc = Num(NmRange(wb, "G_TABLIYE_KALINLIGI").Cells(1).Offset(0, 1).Value, "Tabliye Kalinligi (INPUT column C)")
    ' concrete E (owner 2026-11-06): static and seismic = AASHTO Ec (INPUT column C), RTI = Eurocode Ecm (column D)
    Dim eCol As Long, eTxt As String
    eCol = IIf(rtiMode, 3, 2): eTxt = IIf(rtiMode, "Ecm (INPUT column D)", "AASHTO Ec (INPUT column C)")
    Set r = NmRange(wb, "G_KIRIS_BETONU"): clsG = CStr(r.Cells(1, 1).Value)
    eG = Num(r.Cells(1, eCol).Value, "Kiris betonu " & eTxt)
    Set r = NmRange(wb, "G_TABLIYE_BETONU"): clsS = CStr(r.Cells(1, 1).Value)
    eS = Num(r.Cells(1, eCol).Value, "Tabliye betonu " & eTxt)
    Set r = NmRange(wb, "G_DIGER_BETONLAR"): clsO = CStr(r.Cells(1, 1).Value)
    eO = Num(r.Cells(1, eCol).Value, "Diger betonlar " & eTxt)
    denG = Num(NmRange(wb, "G_BIRIM_AGIRLIK_ONGERME").Cells(1).Value, "G_BIRIM_AGIRLIK_ONGERME")
    denS = Num(NmRange(wb, "G_BIRIM_AGIRLIK_BETONARME").Cells(1).Value, "G_BIRIM_AGIRLIK_BETONARME")
    For i = 1 To 5
        capS(i) = Num(NmRange(wb, "G_BASLIK_KESIT").Cells(i).Value, "G_BASLIK_KESIT row " & i)
    Next i
    capL = Num(NmRange(wb, "G_BASLIK_BOYU").Cells(1).Value, "G_BASLIK_BOYU")
    ' B1 = B2 = 0: rectangle H x tw (owner 2026-11-09); else MIDAS invert-T (UDT): web tw on top, bottom flange
    ' tf with outstands B1 / B2 (read from the MIDAS shape table: H B1 B2 tw tf; outstand meaning not verified live)
    capRect = (capS(2) = 0 And capS(3) = 0)
    If capS(1) <= 0 Or capS(4) <= 0 Then Err.Raise vbObjectError + 37, , "G_BASLIK_KESIT: H and tw must be > 0."
    If capRect Then
        capYc = capS(1) / 2
    Else
        If capS(5) <= 0 Or capS(5) >= capS(1) Then Err.Raise vbObjectError + 37, , "G_BASLIK_KESIT: tf must be between 0 and H."
        capYc = (capS(4) * (capS(1) - capS(5)) ^ 2 / 2 + (capS(2) + capS(4) + capS(3)) * capS(5) * (capS(1) - capS(5) / 2)) / _
                (capS(4) * (capS(1) - capS(5)) + (capS(2) + capS(4) + capS(3)) * capS(5))
    End If
    brgW = NumOr0(NmRange(wb, "G_MESNET_ENI").Cells(1).Value) / 100#
    brgL = NumOr0(NmRange(wb, "G_MESNET_BOYU").Cells(1).Value) / 100#

    sark = Num(NmRange(wb, "G_SARKIKLIK").Cells(1).Value, "G_SARKIKLIK") / 100#
    Dim rp As Worksheet: Set rp = FirstKdme(wb)
    kyB = Num(LabelValue(rp, 3, "Elastomer Yanal Kesme Rijitli", 19), "Elastomer Yanal Kesme Rijitligi")
    kEB = Num(LabelValue(rp, 3, "Elastomer Eksenel Rijitli", 19), "Elastomer Eksenel Rijitligi")

    ' spans (axis to axis), both columns of the span list, first empty cell ends the list
    ReDim spL(1 To 60)
    n = 0
    For Each v In Array("G_ACIKLIK_LISTE_1", "G_ACIKLIK_LISTE_2")
        Set r = NmRange(wb, CStr(v))
        For i = 1 To r.Cells.Count
            If IsEmpty(r.Cells(i).Value) Or Not IsNumeric(r.Cells(i).Value) Then GoTo SpansDone
            If CDbl(r.Cells(i).Value) <= 0 Then GoTo SpansDone
            n = n + 1: spL(n) = CDbl(r.Cells(i).Value)
        Next i
    Next v
SpansDone:
    nSp = n
    If nSp < 1 Then Err.Raise vbObjectError + 30, , "Span list is empty (G_ACIKLIK_LISTE_1)."
    ReDim Preserve spL(1 To nSp)

    ' span table: fixed end, Yatay Gabari, Kiris, Mesnet
    ReDim fixLeft(1 To nSp): ReDim spType(1 To nSp)
    Set r = NmRange(wb, "G_SABIT_UC")
    nDk = 0
    For i = 1 To nSp
        spType(i) = DeckType(wb, r, i)
        CheckBearing r.Cells(i).Offset(0, 3).Value, "Span " & i
        t = UCase$(Left$(Trim$(CStr(r.Cells(i).Value)), 2))
        If t = "SO" Then
            fixLeft(i) = True
        ElseIf Left$(t, 1) = "S" Then
            fixLeft(i) = False
        Else
            Err.Raise vbObjectError + 31, , "G_SABIT_UC span " & i & ": write SOL or SAG."
        End If
    Next i
    ' last right bearing: the row 'Son' under the 60 span rows
    If StrComp(Trim$(CStr(r.Cells(r.Rows.Count + 1).Offset(0, -1).Value)), "Son", vbTextCompare) <> 0 Then Err.Raise vbObjectError + 35, , "Span table: row 'Son' not found under the last span row."
    CheckBearing r.Cells(r.Rows.Count + 1).Offset(0, 3).Value, "Last right"
    ReadSkew wb, r

    ' pier table: one row per axis, label in column 1 (P1, P2, ...); height, type, width, length in 2..5
    Dim tb As Range, rr As Long, found As Boolean, q As Long, en As Double, boy As Double, th As Double
    Set tb = NmRange(wb, "G_AYAK_TABLOSU")
    nPS = 0
    If nSp > 1 Then
        ReDim pH(1 To nSp - 1): ReDim pSec(1 To nSp - 1)
        ReDim psTyp(1 To nSp): ReDim psEn(1 To nSp): ReDim psBoy(1 To nSp): ReDim psT(1 To nSp)
    End If
    For i = 1 To nSp - 1
        found = False
        For rr = 1 To tb.Rows.Count
            If StrComp(Trim$(CStr(tb.Cells(rr, 1).Value)), "P" & i, vbTextCompare) = 0 Then found = True: Exit For
        Next rr
        If Not found Then Err.Raise vbObjectError + 32, , "G_AYAK_TABLOSU: no row for P" & i & "."
        pH(i) = Num(tb.Cells(rr, 2).Value, "P" & i & " height")
        If pH(i) <= 0 Then Err.Raise vbObjectError + 32, , "P" & i & " height must be > 0."
        t = PierShape(CStr(tb.Cells(rr, 3).Value))
        en = Num(tb.Cells(rr, 4).Value, "P" & i & " En / Cap")
        boy = NumOr0(tb.Cells(rr, 5).Value)
        th = NumOr0(tb.Cells(rr, 6).Value)
        If t <> "SR" And boy <= 0 Then Err.Raise vbObjectError + 32, , "P" & i & " Boy must be > 0."
        If t = "B" And (th <= 0 Or 2 * th >= en Or 2 * th >= boy) Then Err.Raise vbObjectError + 32, , "P" & i & " Kalinlik is not valid for a box."
        For q = 1 To nPS
            If psTyp(q) = t And Abs(psEn(q) - en) < 0.0001 And Abs(psBoy(q) - boy) < 0.0001 And Abs(psT(q) - th) < 0.0001 Then Exit For
        Next q
        If q > nPS Then
            nPS = q: psTyp(q) = t: psEn(q) = en: psBoy(q) = boy: psT(q) = th
        End If
        pSec(i) = SEC_PIER0 + q - 1
    Next i
    ReadLoadData wb
    For i = 1 To nDk
        If (tJoint(i) - tLc(i)) / 2 <= sark + 0.0005 Then warnText = warnText & "- " & tName(i) & ": joint spacing gives no deck-end slab." & vbCrLf
    Next i
End Sub

' first span sheet of the 4A chapter: bearing values (only Mesnet - 1 today)
Private Function FirstKdme(wb As Workbook) As Worksheet
    Dim i As Long
    On Error Resume Next
    For i = 1 To 6
        Set FirstKdme = wb.Worksheets("4" & Chr$(64 + i) & "_KDME")
        If Not FirstKdme Is Nothing Then Exit For
    Next i
    On Error GoTo 0
    If FirstKdme Is Nothing Then Err.Raise vbObjectError + 56, , "No span sheet 4<letter>_KDME in this workbook."
End Function

Private Function RailMat() As String
    RailMat = QS(CStr(MAT_RAIL)) & ":{""TYPE"":""USER"",""NAME"":""RAY"",""HE_SPEC"":0,""HE_COND"":0,""PLMT"":0,""P_NAME"":"""",""bMASS_DENS"":false,""DAMP_RAT"":0.02,""PARAM"":[{""P_TYPE"":2,""ELAST"":" & JN(RAIL_E) & ",""POISN"":0.3,""THERMAL"":0.000012,""DEN"":0,""MASS"":0}]}"
End Function

Private Function RailSect() As String
    ' the two UIC 60 rails of one track as one rail-shaped H (owner 2026-11-09): head 2 x 72, web 2 x 16.5,
    ' foot 2 x 150, height 172 mm, as the rail of models/RTI.mct halved; A 0.015342 m2, Iyy 6.064e-5 m4
    ' (2 x UIC 60: 0.01534, 6.076e-5), so the Beam Stress table gives the rail stress directly
    RailSect = DbSect(SEC_RAIL, "RAY", SectHead2("CC", 0, 0), "H", RAIL_SIZE)
End Function

Private Function GrupAll() As String
    Dim s As String, i As Long, t As Long, k As Long, n As Long, nl As String, el As String, ax As String
    s = """4"":{""NAME"":""SG_TUM"",""P_TYPE"":0,""N_LIST"":[" & Seq(1, nCnt) & "," & Seq(NUM_LIMIT, NUM_LIMIT + xNo - 1) & "],""E_LIST"":[" & Seq(1, eCnt) & "," & Seq(NUM_LIMIT, NUM_LIMIT + reCnt - 1) & "]}"
    n = 4
    ' groups the RTI result pull reads: rail elements and nodes per track, track nodes (rigid to the deck
    ' on the bridge, fixed on the embankment) per track, deck end node left / right of every axis
    For t = 1 To nTrk
        nl = "": el = ""
        For i = 1 To rN(t): nl = nl & IIf(nl = "", "", ",") & rNode(t, i): Next i
        For i = 1 To reCnt
            If rET(i) = t Then el = el & IIf(el = "", "", ",") & (NUM_LIMIT + i - 1)
        Next i
        n = n + 1: s = s & "," & QS(CStr(n)) & ":{""NAME"":""SG_RAY_" & t & """,""P_TYPE"":0,""N_LIST"":[" & nl & "],""E_LIST"":[" & el & "]}"
        nl = ""
        For i = 1 To rN(t): nl = nl & IIf(nl = "", "", ",") & rGr(t, i): Next i
        n = n + 1: s = s & "," & QS(CStr(n)) & ":{""NAME"":""SG_IZ_" & t & """,""P_TYPE"":0,""N_LIST"":[" & nl & "]}"
    Next t
    For k = 0 To nSp
        ax = IIf(k = 0, "A1", IIf(k = nSp, "A2", "P" & k))
        If k >= 1 Then n = n + 1: s = s & "," & QS(CStr(n)) & ":{""NAME"":""DE_" & ax & "_L"",""P_TYPE"":0,""N_LIST"":[" & eN2(dkB(k)) & "]}"
        If k < nSp Then n = n + 1: s = s & "," & QS(CStr(n)) & ":{""NAME"":""DE_" & ax & "_R"",""P_TYPE"":0,""N_LIST"":[" & eN1(dkA(k + 1)) & "]}"
    Next k
    GrupAll = s
End Function

Private Function Seq(ByVal a As Long, ByVal b As Long) As String
    Dim i As Long, s() As String
    ReDim s(0 To b - a)
    For i = a To b: s(i - a) = CStr(i): Next i
    Seq = Join(s, ",")
End Function

' deck type of span i: one type per Yatay Gabari + Kiris combination
Private Function DeckType(wb As Workbook, r As Range, ByVal i As Long) As Long
    Dim gab As Long, k As Long, d As Long, t As String, p As Long, c As Range
    gab = CLng(Num(r.Cells(i).Offset(0, 1).Value, "Yatay Gabari, span " & i))
    t = Trim$(CStr(r.Cells(i).Offset(0, 2).Value))
    p = InStrRev(t, "-")
    If p = 0 Then Err.Raise vbObjectError + 36, , "Kiris, span " & i & ": '" & t & "' has no Hesap Acikligi number."
    k = CLng(Val(Mid$(t, p + 1)))
    If k < 1 Or k > 6 Then Err.Raise vbObjectError + 36, , "Kiris, span " & i & ": '" & t & "'."
    For d = 1 To nDk
        If tHes(d) = k And tGab(d) = gab Then DeckType = d: Exit Function
    Next d
    If nDk = MAX_TYPES Then Err.Raise vbObjectError + 37, , "More than " & MAX_TYPES & " deck types."
    nDk = nDk + 1: d = nDk
    tHes(d) = k: tGab(d) = gab
    tLc(d) = Num(NmRange(wb, "G_HESAP_ACIKLIK").Cells(k).Value, "G_HESAP_ACIKLIK " & k)
    tNG(d) = CLng(Num(NmRange(wb, "G_KIRIS_SAYISI").Cells(k).Value, "G_KIRIS_SAYISI " & k))
    tSG(d) = Num(NmRange(wb, "G_KIRIS_ARALIK").Cells(k).Value, "G_KIRIS_ARALIK " & k)
    Hg = Num(NmRange(wb, "G_KIRIS_YUKSEKLIGI").Cells(k).Value, "G_KIRIS_YUKSEKLIGI " & k)
    If tNG(d) < 2 Then Err.Raise vbObjectError + 2, , "G_KIRIS_SAYISI " & k & " must be >= 2."
    tW(d) = Num(NmRange(wb, "G_GABARI_KALDIRIM_SOL").Cells(gab).Value, "G_GABARI_KALDIRIM_SOL") + _
            Num(NmRange(wb, "G_GABARI_HAT_1").Cells(gab).Value, "G_GABARI_HAT_1") + _
            NumOr0(NmRange(wb, "G_GABARI_HAT_2").Cells(gab).Value) + _
            Num(NmRange(wb, "G_GABARI_KALDIRIM_SAG").Cells(gab).Value, "G_GABARI_KALDIRIM_SAG")
    ReadGirder wb, CStr(NmRange(wb, "G_KIRIS_TIPI").Cells(k).Value)
    tHg(d) = Hg
    For p = 1 To 12: tKb(d, p) = kb(p): Next p
    Set c = NmRange(wb, "G_GENLESME_DERZ_ARALIGI").Cells(1, 1)
    tJoint(d) = NumOr0(c.Offset(0, k - 1).Value)
    If tJoint(d) <= 0 Then tJoint(d) = Num(c.Value, "G_GENLESME_DERZ_ARALIGI")
    tName(d) = "Deck type " & d & " (Hesap Acikligi " & k & ", Gabari " & gab & ")"
    tKL(d) = NmRange(wb, "G_GABARI_KALDIRIM_SOL").Cells(gab).Value: tKR(d) = NmRange(wb, "G_GABARI_KALDIRIM_SAG").Cells(gab).Value
    tH1(d) = NmRange(wb, "G_GABARI_HAT_1").Cells(gab).Value: tH2(d) = NumOr0(NmRange(wb, "G_GABARI_HAT_2").Cells(gab).Value)
    For p = 1 To 7
        tQ(d, p) = Num(NmRange(wb, CStr(Array("G_YUK_BALAST", "G_YUK_BORDUR", "G_YUK_KORKULUK", "G_YUK_CEPHE_PANELI", "G_YUK_KORUMA_BETONU", "G_YUK_RAY", "G_YUK_ENLEME")(p - 1))).Cells(k).Value, "Yuk Bilgileri (G_YUK_*), Hesap Acikligi " & k)
    Next p
    tPhi(d) = Num(NmRange(wb, "G_PHI3").Cells(k).Value, "G_PHI3, Hesap Acikligi " & k)
    DeckType = d
End Function

Private Sub UseType(ByVal d As Long)
    Dim i As Long
    nG = tNG(d): sG = tSG(d): Hg = tHg(d): Lc = tLc(d): deckW = tW(d): jointL = tJoint(d)
    For i = 1 To 12: kb(i) = tKb(d, i): Next i
    GirderPolygon
End Sub

Private Function DeckSec(ByVal d As Long) As Long
    DeckSec = IIf(d = 1, 1, SEC_DECK0 + d)
End Function

Private Function SlabSec(ByVal d As Long) As Long
    SlabSec = IIf(d = 1, 2, SEC_SLAB0 + d)
End Function

' 4-7 reports only the first bearing today (owner 2026-11-05: it will report the others later)
Private Sub CheckBearing(ByVal v As Variant, ByVal where As String)
    If Val(Mid$(Trim$(CStr(v)), InStrRev(Trim$(CStr(v)), "-") + 1)) <> 1 Then Err.Raise vbObjectError + 38, , where & ": '" & v & "'. 4-7 reports kv and ky for Mesnet - 1 only for now."
End Sub

' plan of the axes: G_AKS_XY (DXF_OKU) if filled, else straight
' TKDY checks of the seismic model, warnings only (owner 2026-11-09): vertical earthquake (4.2.1.3) and
' deck as one bar line (4.3.2.4). Span = axis to axis chord, width = deck width of the span's deck type.
Private Sub SeisChecks(wb As Workbook)
    Dim i As Long, kos As Double, lMax As Double, iMax As Long, r As Double, rMin As Double, iMin As Long
    kos = NumOr0(NmRange(wb, "G_KOPRU_ONEM_SINIFI").Cells(1).Value)
    rMin = 1E+30
    For i = 1 To nSp
        If lenS(i) > lMax Then lMax = lenS(i): iMax = i
        If tW(spType(i)) > 0 Then
            r = lenS(i) / tW(spType(i))
            If r < rMin Then rMin = r: iMin = i
        End If
    Next i
    If lMax > 80 And (kos = 1 Or kos = 2) Then warnText = warnText & "- Span " & iMax & " is " & Format(lMax, "0.0") & " m (> 80 m) with KOS " & kos & ": TKDY 4.2.1.3 asks for the vertical earthquake when DTS is 1 or 2. The model has no Z spectrum case." & vbCrLf
    If iMin > 0 And rMin <= 2 Then warnText = warnText & "- Span " & iMin & ": span / deck width = " & Format(rMin, "0.00") & " (<= 2). TKDY 4.3.2.4 allows the deck as one bar line only above 2 (otherwise 4.3.2.5)." & vbCrLf
End Sub

' skew (owner 2026-11-09): span table column R (4 right of G_SABIT_UC) = 1..5 picks G_VEREVLIK B22:F22, 0 or empty
' = square. Degrees from the normal to the girders, + = support line turned counter-clockwise in plan (A1 -> A2
' left to right). A span has its skew at both ends; the two spans at a pier must agree (one cap, one direction).
Private Sub ReadSkew(wb As Workbook, r As Range)
    Dim i As Long, v As Variant, k As Long, sk As Range
    Set sk = NmRange(wb, "G_VEREVLIK").Cells(1)
    ReDim skS(1 To nSp): ReDim skA(0 To nSp)
    For i = 1 To nSp
        v = r.Cells(i).Offset(0, 4).Value
        k = 0
        If Trim$(CStr(v)) <> "" Then
            If Not IsNumeric(v) Then Err.Raise vbObjectError + 63, , "Span table Verevlik, span " & i & ": '" & v & "' (write 1 to 5 = G_VEREVLIK B22:F22, or 0)."
            k = CLng(v)
            If k <> CDbl(v) Or k < 0 Or k > 5 Then Err.Raise vbObjectError + 63, , "Span table Verevlik, span " & i & ": '" & v & "' (write 1 to 5 = G_VEREVLIK B22:F22, or 0)."
        End If
        If k > 0 Then skS(i) = Num(sk.Offset(0, k - 1).Value, "G_VEREVLIK skew " & k & " (INPUT row 22), span " & i)
        If Abs(skS(i)) >= 85 Then Err.Raise vbObjectError + 63, , "Span " & i & ": skew " & skS(i) & " deg cannot be built."
    Next i
    skA(0) = skS(1): skA(nSp) = skS(nSp)
    For i = 1 To nSp - 1
        If Abs(skS(i) - skS(i + 1)) > 0.0001 Then Err.Raise vbObjectError + 64, , "Spans " & i & " and " & i + 1 & " meet at one pier with skews " & skS(i) & " and " & skS(i + 1) & " deg: one cap has one direction."
        skA(i) = skS(i)
    Next i
End Sub

Private Function SkewText() As String
    Dim i As Long, n As Long, mx As Double
    For i = 0 To nSp
        If skA(i) <> 0 Then n = n + 1
        If Abs(skA(i)) > mx Then mx = Abs(skA(i))
    Next i
    If n = 0 Then SkewText = "Skew: none (square)" Else SkewText = "Skew: " & n & " of " & nSp + 1 & " axes, largest " & Format(mx, "0.0") & " deg; bearing links along the " & IIf(LINK_ON_SKEW, "support line", "girders")
End Function

Private Sub BuildPlan()
    Dim i As Long
    ReDim pxA(0 To nSp): ReDim pyA(0 To nSp): ReDim rxA(0 To nSp): ReDim ryA(0 To nSp)
    ReDim cxS(1 To nSp): ReDim cyS(1 To nSp): ReDim lenS(1 To nSp): ReDim phiS(1 To nSp): ReDim phiA(0 To nSp)
    ReDim spR(1 To nSp)
    nAxXY = ReadAxisXY()
    If nAxXY = 0 Then
        warnText = warnText & "- G_AKS_XY is empty: the bridge is built STRAIGHT. Check the name range or run DXF OKU." & vbCrLf
        For i = 1 To nSp: pxA(i) = pxA(i - 1) + spL(i): Next i
    Else
        ' model X = chord A1 to A2
        Dim a As Double, c As Double, sn As Double, xx As Double
        a = Atn2(pyA(nSp) - pyA(0), pxA(nSp) - pxA(0))
        c = Cos(-a): sn = Sin(-a)
        For i = 0 To nSp
            xx = pxA(i) * c - pyA(i) * sn
            pyA(i) = pxA(i) * sn + pyA(i) * c
            pxA(i) = xx
            If Abs(pyA(i)) > maxOff Then maxOff = Abs(pyA(i))
        Next i
    End If
    For i = 1 To nSp
        lenS(i) = Sqr((pxA(i) - pxA(i - 1)) ^ 2 + (pyA(i) - pyA(i - 1)) ^ 2)
        cxS(i) = (pxA(i) - pxA(i - 1)) / lenS(i): cyS(i) = (pyA(i) - pyA(i - 1)) / lenS(i)
        phiS(i) = Atn2(cyS(i), cxS(i)) * 180 / PI
    Next i
    ' cap direction at an axis = bisector of the normals of its two chords
    Dim nx As Double, ny As Double, q As Double, ck As Double, sq As Double
    For i = 0 To nSp
        If i = 0 Then
            nx = -cyS(1): ny = cxS(1)
        ElseIf i = nSp Then
            nx = -cyS(nSp): ny = cxS(nSp)
        Else
            nx = -cyS(i) - cyS(i + 1): ny = cxS(i) + cxS(i + 1)
        End If
        q = Sqr(nx ^ 2 + ny ^ 2)
        rxA(i) = nx / q: ryA(i) = ny / q
        If skA(i) <> 0 Then
            ck = Cos(skA(i) * PI / 180): sq = Sin(skA(i) * PI / 180)
            q = rxA(i): rxA(i) = q * ck - ryA(i) * sq: ryA(i) = q * sq + ryA(i) * ck
        End If
        phiA(i) = Atn2(ryA(i), rxA(i)) * 180 / PI - 90
    Next i
End Sub

' G_AKS_XY (DXF_OKU): axis X / Y in drawing coordinates, A1 first. Returns the number of points
' (0 = name missing or table empty); fills pxA / pyA relative to A1. Each chord must match its span.
Private Function ReadAxisXY() As Long
    Dim t As Range, r As Long, n As Long, i As Long, x0 As Double, y0 As Double, ch As Double
    On Error Resume Next
    Set t = ThisWorkbook.Names("G_AKS_XY").RefersToRange
    On Error GoTo 0
    If t Is Nothing Then Exit Function
    For r = 1 To t.Rows.Count
        If IsEmpty(t.Cells(r, 1).Value) Then Exit For
        n = n + 1
    Next r
    If n = 0 Then Exit Function
    If n <> nSp + 1 Then Err.Raise vbObjectError + 43, , "G_AKS_XY holds " & n & " axis points, the span list needs " & nSp + 1 & ". Run DXF OKU again."
    x0 = Num(t.Cells(1, 1).Value, "G_AKS_XY X row 1"): y0 = Num(t.Cells(1, 2).Value, "G_AKS_XY Y row 1")
    For i = 0 To nSp
        pxA(i) = Num(t.Cells(i + 1, 1).Value, "G_AKS_XY X row " & i + 1) - x0
        pyA(i) = Num(t.Cells(i + 1, 2).Value, "G_AKS_XY Y row " & i + 1) - y0
        If i > 0 Then
            spR(i) = NumOr0(t.Cells(i, 3).Value)   ' span radius next to the table (DXF_OKU): < 0 left, > 0 right, empty = straight
            ch = Sqr((pxA(i) - pxA(i - 1)) ^ 2 + (pyA(i) - pyA(i - 1)) ^ 2)
            If Abs(ch - spL(i)) > 0.05 Then Err.Raise vbObjectError + 44, , "G_AKS_XY: span " & i & " chord " & Format(ch, "0.000") & " m, span list " & Format(spL(i), "0.000") & " m. Run DXF OKU again."
        End If
    Next i
    ReadAxisXY = n
    For i = 1 To nSp
        If spR(i) <> 0 Then Exit Function
    Next i
    warnText = warnText & "- No span radius next to G_AKS_XY: no centrifugal case. Run DXF OKU if the bridge is on a curve." & vbCrLf
End Function

Private Function Atn2(ByVal yy As Double, ByVal xx As Double) As Double
    If Abs(xx) < 0.000000000001 Then
        Atn2 = IIf(yy >= 0, PI / 2, -PI / 2)
    ElseIf xx > 0 Then
        Atn2 = Atn(yy / xx)
    Else
        Atn2 = Atn(yy / xx) + IIf(yy >= 0, PI, -PI)
    End If
End Function

' node of span i at u along the chord from its left axis, v to the left of it
Private Function SpN(ByVal i As Long, ByVal u As Double, ByVal v As Double, ByVal z As Double) As Long
    SpN = AddNode(pxA(i - 1) + u * cxS(i) - v * cyS(i), pyA(i - 1) + u * cyS(i) + v * cxS(i), z)
End Function

' node on the cap line of axis a at distance v from the axis
Private Function CapN(ByVal a As Long, ByVal v As Double, ByVal z As Double) As Long
    CapN = AddNode(pxA(a) + v * rxA(a), pyA(a) + v * ryA(a), z)
End Function

' INPUT column type -> MIDAS DB user shape
Private Function PierShape(ByVal s As String) As String
    Dim t As String: t = UCase$(Trim$(s))
    If t = "OVAL" Then
        PierShape = "STRK"
    ElseIf t = "KUTU" Then
        PierShape = "B"
    ElseIf InStr(t, "RTGEN") > 0 Then
        PierShape = "SB"
    ElseIf InStr(t, "RESEL") > 0 Then
        PierShape = "SR"
    Else
        Err.Raise vbObjectError + 33, , "Pier type '" & s & "' is not supported yet (OVAL, KUTU, DIKDORTGEN, DAIRESEL)."
    End If
End Function

' bearing offsets from the axes (owner 2026-11-09): e = (axis spacing - Hesap Acikligi) / 2 for inner spans,
' axis spacing - Hesap Acikligi for end spans (A1 / A2 bearing axis on the abutment axis); the two spans at a
' pier must give the same e. Seat: rectangle or e on the web = cap top; inverted T with e beyond the web =
' ledge (flange top). Cap length and width are checked against the bearing size (warnings).
Private Sub BearingOffsets()
    Dim i As Long, Lci As Double, ys As Double, d As Variant, msg As String, half As Double, ep As Double
    ReDim eLs(1 To nSp): ReDim eRs(1 To nSp): ReDim ePier(0 To nSp): ReDim seatOnLedge(0 To nSp)
    For i = 1 To nSp
        Lci = tLc(spType(i))
        If nSp = 1 Then
            eLs(i) = 0: eRs(i) = 0
        ElseIf i = 1 Then
            eLs(i) = 0: eRs(i) = lenS(i) - Lci
        ElseIf i = nSp Then
            eLs(i) = lenS(i) - Lci: eRs(i) = 0
        Else
            eLs(i) = (lenS(i) - Lci) / 2: eRs(i) = eLs(i)
        End If
        If eLs(i) < -0.0005 Or eRs(i) < -0.0005 Then msg = msg & "span " & i & ": axis spacing " & Format(lenS(i), "0.000") & " m < Hesap Acikligi " & Lci & " m; "
    Next i
    For i = 1 To nSp - 1
        If Abs(eRs(i) - eLs(i + 1)) > 0.02 Then msg = msg & "axis " & i & ": " & Format(eRs(i), "0.000") & " m from span " & i & ", " & Format(eLs(i + 1), "0.000") & " m from span " & i + 1 & "; "
        ePier(i) = IIf(eRs(i) > eLs(i + 1), eRs(i), eLs(i + 1))
        seatOnLedge(i) = (Not capRect) And ePier(i) * Cos(skA(i) * PI / 180) > capS(4) / 2 + 0.0005
    Next i
    If msg <> "" Then Err.Raise vbObjectError + 38, , "Axis to bearing distance does not fit the span list and Hesap Acikligi: " & msg
    ' warnings: bearings on the cap
    For i = 1 To nSp - 1
        ep = ePier(i) * Cos(skA(i) * PI / 180)
        If capRect Or Not seatOnLedge(i) Then
            If ep + brgL / 2 > capS(4) / 2 + 0.0005 Then warnText = warnText & "- Axis " & i & ": bearing (e " & Format(ePier(i), "0.000") & " + " & brgL / 2 & ") reaches past the cap width " & capS(4) & " m." & vbCrLf
        Else
            half = capS(4) / 2 + IIf(capS(2) < capS(3), capS(2), capS(3))
            If ep + brgL / 2 > half + 0.0005 Or ep - brgL / 2 < capS(4) / 2 - 0.0005 Then warnText = warnText & "- Axis " & i & ": bearing (e " & Format(ePier(i), "0.000") & " m, length " & brgL & " m) does not fit on the ledge between " & capS(4) / 2 & " and " & Format(half, "0.000") & " m." & vbCrLf
        End If
        For Each d In Array(spType(i), spType(i + 1))
            ys = (tNG(d) - 1) / 2# * tSG(d) / Cos(skA(i) * PI / 180) + brgW / 2
            If ys > capL / 2 + 0.0005 Then warnText = warnText & "- Axis " & i & ": outer bearing reaches " & Format(ys, "0.000") & " m from the axis, cap half length " & Format(capL / 2, "0.000") & " m." & vbCrLf
        Next d
    Next i
End Sub

Private Sub BuildModel()
    Dim i As Long, j As Long, m As Long, k As Long, a As Long, b As Long, xL As Double, xR As Double, ext As Double
    Dim zLs() As Double, y() As Double, dx As Double, zCap() As Double
    Set nodeKey = CreateObject("Scripting.Dictionary")
    Set capSlaves = CreateObject("Scripting.Dictionary")
    Set rigMap = CreateObject("Scripting.Dictionary")
    nCnt = 0: eCnt = 0: nLink = 0: nBrg = 0
    ReDim dkA(1 To nSp): ReDim dkB(1 To nSp): ReDim mnA(1 To nSp): ReDim mnB(1 To nSp): ReDim dxS(1 To nSp)
    ReDim brgNode(1 To 2 * nSp): ReDim brgSpan(1 To 2 * nSp)
    grpDeck = "": grpCap = "": grpPier = "": rigdJson = "": elnkJson = "": consJson = ""
    ReDim nX(1 To 2000): ReDim nY(1 To 2000): ReDim nZ(1 To 2000)
    ReDim eMat(1 To 2000): ReDim eSec(1 To 2000): ReDim eN1(1 To 2000): ReDim eN2(1 To 2000)
    ReDim zLs(1 To nSp): ReDim zCap(0 To nSp)
    Set pierAng = CreateObject("Scripting.Dictionary")
    maxOff = 0
    BuildPlan
    For i = 1 To nSp: zLs(i) = -(tHg(spType(i)) + tc) - linkLen: Next i
    ' cap level at a pier = the lower bearing bottom of its two spans
    For i = 1 To nSp - 1
        zCap(i) = IIf(zLs(i) < zLs(i + 1), zLs(i), zLs(i + 1))
    Next i
    BearingOffsets
    ' zCap = bearing seat; cap top above it by H - tf when the bearings sit on the ledges of the inverted T
    Dim zTop() As Double: ReDim zTop(0 To nSp)
    For i = 1 To nSp - 1
        zTop(i) = zCap(i) + IIf(seatOnLedge(i), capS(1) - capS(5), 0)
    Next i

    For i = 1 To nSp
        UseType spType(i)
        xL = eLs(i)
        xR = lenS(i) - eRs(i)
        If Abs(xR - xL - Lc) > 0.01 Then warnText = warnText & "- Span " & i & ": bearing to bearing " & Format(xR - xL, "0.000") & " m, Hesap Acikligi " & Lc & " m." & vbCrLf
        ext = (jointL - Lc) / 2
        m = CLng((xR - xL) / meshDx): If m < 1 Then m = 1
        dx = (xR - xL) / m
        dxS(i) = dx: dkA(i) = eCnt + 1
        If i > 1 And ext > sark + 0.0005 Then Member grpDeck, 2, SlabSec(spType(i)), SpN(i, xL - ext, 0, 0), SpN(i, xL - sark, 0, 0)
        Member grpDeck, 1, DeckSec(spType(i)), SpN(i, xL - sark, 0, 0), SpN(i, xL, 0, 0)
        mnA(i) = eCnt + 1
        For k = 1 To m
            Member grpDeck, 1, DeckSec(spType(i)), SpN(i, xL + (k - 1) * dx, 0, 0), SpN(i, IIf(k = m, xR, xL + k * dx), 0, 0)
        Next k
        mnB(i) = eCnt
        Member grpDeck, 1, DeckSec(spType(i)), SpN(i, xR, 0, 0), SpN(i, xR + sark, 0, 0)
        If i < nSp And ext > sark + 0.0005 Then Member grpDeck, 2, SlabSec(spType(i)), SpN(i, xR + sark, 0, 0), SpN(i, xR + ext, 0, 0)
        dkB(i) = eCnt
    Next i

    For i = 1 To nSp
        UseType spType(i)
        ReDim y(0 To nG - 1)
        For j = 0 To nG - 1: y(j) = (j - (nG - 1) / 2#) * sG: Next j
        xL = eLs(i)
        xR = lenS(i) - eRs(i)
        BearingLine i, xL, i - 1, (i = 1), fixLeft(i), y, zLs(i) + linkLen, zLs(i), zTop(i - 1) - capYc
        BearingLine i, xR, i, (i = nSp), Not fixLeft(i), y, zLs(i) + linkLen, zLs(i), zTop(i) - capYc
    Next i

    ' caps and piers
    Dim ys As Object, yk As Variant, top As Long, q As Variant, sp As Variant, nEl As Long
    For i = 1 To nSp - 1
        Set ys = CreateObject("Scripting.Dictionary")
        AddX ys, -capL / 2: AddX ys, capL / 2: AddX ys, 0
        For Each sp In Array(i, i + 1)
            For j = 0 To tNG(spType(sp)) - 1: AddX ys, (j - (tNG(spType(sp)) - 1) / 2#) * tSG(spType(sp)) / Cos(skA(i) * PI / 180): Next j
        Next sp
        yk = SortedKeys(ys)
        For k = UBound(yk) To 2 Step -1
            Member grpCap, 3, SEC_CAP, CapN(i, yk(k), zTop(i) - capYc), CapN(i, yk(k - 1), zTop(i) - capYc)
        Next k
        ' pier top = cap bottom (owner 2026-11-09), rigid to the cap centroid node
        top = CapN(i, 0, zTop(i) - capS(1))
        Rigid CapN(i, 0, zTop(i) - capYc), CStr(top)
        a = top
        nEl = -Int(-pH(i) / MAX_PIER_EL + 0.000000001)
        If nEl < N_PIER_EL Then nEl = N_PIER_EL
        For k = 1 To nEl
            b = CapN(i, 0, zTop(i) - capS(1) - pH(i) * k / nEl)
            Member grpPier, 3, pSec(i), a, b
            pierAng(CStr(eCnt)) = -phiA(i)
            a = b
        Next k
        Fix1 a
    Next i
    For Each q In capSlaves.Keys
        Rigid CLng(q), CStr(capSlaves(q))
    Next q
    If nCnt >= NUM_LIMIT Or eCnt >= NUM_LIMIT Then Err.Raise vbObjectError + 34, , "Node or element numbers reach " & NUM_LIMIT & " (reserved for the rail layer)."
    DeckProps
End Sub

' one bearing line: deck node -> rigid -> bearing top nodes -> elastic links -> bottom nodes
' -> rigid to the cap (pier) or fixed (abutment)
Private Sub BearingLine(ByVal sp As Long, ByVal u As Double, ByVal axisNo As Long, ByVal abut As Boolean, ByVal fixedEnd As Boolean, y() As Double, ByVal zU As Double, ByVal zL As Double, ByVal zC As Double)
    Dim j As Long, up As Long, lo As Long, sl As String, cn As Long, kLong As Double, tk As Double, ck As Double
    tk = Tan(skA(axisNo) * PI / 180): ck = Cos(skA(axisNo) * PI / 180)
    kLong = IIf(fixedEnd, kFix, kyB)
    For j = 0 To nG - 1
        up = SpN(sp, u - y(j) * tk, y(j), zU)
        lo = SpN(sp, u - y(j) * tk, y(j), zL)
        sl = sl & IIf(sl = "", "", ",") & up
        nLink = nLink + 1
        elnkJson = elnkJson & IIf(elnkJson = "", "", ",") & QS(CStr(nLink)) & ":{""NODE"":[" & up & "," & lo & "],""ANGLE"":" & JN(-phiS(sp) - IIf(LINK_ON_SKEW, skA(axisNo), 0)) & ",""LINK"":""GEN""," & _
            """R_S"":[false,false,false,false,false,false],""SDR"":[" & JN(kEB) & "," & JN(kFix) & "," & JN(kLong) & ",0,0,0],""bSHEAR"":false,""DR"":[0.5,0.5],""BNGR_NAME"":""BG_MESNET""}"
        If abut Then
            Fix1 lo
        Else
            cn = CapN(axisNo, y(j) / ck, zC)
            If capSlaves.Exists(CStr(cn)) Then
                capSlaves(CStr(cn)) = capSlaves(CStr(cn)) & "," & lo
            Else
                capSlaves(CStr(cn)) = CStr(lo)
            End If
        End If
    Next j
    Rigid SpN(sp, u, 0, 0), sl
    nBrg = nBrg + 1: brgNode(nBrg) = SpN(sp, u, 0, 0): brgSpan(nBrg) = sp
End Sub

Private Sub Rigid(ByVal master As Long, ByVal slaves As String)
    If rigMap.Exists(master) Then
        rigMap(master) = rigMap(master) & "," & slaves
    Else
        rigMap(master) = slaves
    End If
End Sub

Private Function RigdAll() As String
    Dim k As Variant
    For Each k In rigMap.Keys
        RigdAll = RigdAll & IIf(RigdAll = "", "", ",") & QS(CStr(k)) & ":{""ITEMS"":[{""ID"":1,""GROUP_NAME"":""BG_MESNET"",""DOF"":111111,""S_NODE"":[" & rigMap(k) & "]}]}"
    Next k
End Function

Private Sub Fix1(ByVal n As Long)
    consJson = consJson & IIf(consJson = "", "", ",") & QS(CStr(n)) & ":{""ITEMS"":[{""ID"":1,""GROUP_NAME"":""BG_MESNET"",""CONSTRAINT"":""1111110""}]}"
End Sub

Private Sub Member(ByRef grp As String, ByVal mat As Long, ByVal sec As Long, ByVal a As Long, ByVal b As Long)
    Dim e As Long
    e = AddElem(mat, sec, a, b)
    grp = grp & IIf(grp = "", "", ",") & e
End Sub

' girder outline (KIRIS_INPUT), X from the girder axis, Y from the soffit, shifted by xs, ys
Private Function GirderPts(ByVal xs As Double, ByVal ys As Double) As String
    Dim k(1 To 12) As Double, i As Long, yA As Double, yB As Double, yC As Double, yD As Double, p As Variant, s As String
    For i = 1 To 12: k(i) = kb(i): Next i
    yA = Hg - k(7): yB = yA - k(8): yC = yB - k(9): yD = k(12)
    p = Array(k(3) / 2, 0#, k(3) / 2, yD, k(2) / 2, yD + k(11), k(2) / 2, yC, k(2) / 2 + k(4), yB, k(1) / 2, yA, k(1) / 2, Hg, _
              -k(1) / 2, Hg, -k(1) / 2, yA, -(k(2) / 2 + k(4)), yB, -k(2) / 2, yC, -k(2) / 2, yD + k(11), -k(3) / 2, yD, -k(3) / 2, 0#)
    For i = 0 To UBound(p) Step 2
        s = s & IIf(i = 0, "", ", ") & JN(p(i) + xs) & ", " & JN(p(i + 1) + ys)
    Next i
    GirderPts = s
End Function

' girder properties from the outline (polygon formulas): area, centroid above the soffit, I about the
' horizontal and vertical centroid axes, perimeter
Private Sub GirderPolygon()
    Dim a() As String, i As Long, n As Long, px() As Double, py() As Double, q As Long, c As Double
    Dim sa As Double, sy As Double, sxx As Double, syy As Double, per As Double
    a = Split(GirderPts(0, 0), ",")
    n = (UBound(a) + 1) / 2
    ReDim px(1 To n): ReDim py(1 To n)
    For i = 1 To n
        px(i) = Val(a(2 * i - 2)): py(i) = Val(a(2 * i - 1))
    Next i
    For i = 1 To n
        q = IIf(i = n, 1, i + 1)
        c = px(i) * py(q) - px(q) * py(i)
        sa = sa + c
        sy = sy + (py(i) + py(q)) * c
        sxx = sxx + (py(i) ^ 2 + py(i) * py(q) + py(q) ^ 2) * c
        syy = syy + (px(i) ^ 2 + px(i) * px(q) + px(q) ^ 2) * c
        per = per + Sqr((px(q) - px(i)) ^ 2 + (py(q) - py(i)) ^ 2)
    Next i
    polyArea = Abs(sa) / 2
    gCz = Abs(sy / (3 * sa))
    gIy = Abs(sxx) / 12 - polyArea * gCz ^ 2
    gIz = Abs(syy) / 12
    gPer = per
    deckArea = nG * polyArea + deckW * tc * eS / eG
End Sub

' three stiffness lines of a composite part (section properties in the MCT layout)
Private Function StiffLines(ByVal a As Double, ByVal aRaw As Double, ByVal asy As Double, ByVal asz As Double, ByVal ixx As Double, ByVal iyy As Double, ByVal izz As Double, _
        ByVal hw As Double, ByVal hwB As Double, ByVal cTop As Double, ByVal cBot As Double, ByVal per As Double, ByVal cy As Double, ByVal cz As Double) As String
    StiffLines = "       " & JN(a) & ", " & JN(aRaw) & ", " & JN(asy) & ", " & JN(asz) & ", " & JN(ixx) & ", " & JN(iyy) & ", " & JN(izz) & vbCrLf & _
        "       " & JN(hw) & ", " & JN(hw) & ", " & JN(cTop) & ", " & JN(cBot) & ", 0, 0, " & JN(per) & ", 0, " & JN(cy) & ", " & JN(cz) & vbCrLf & _
        "       " & JN(-hw) & ", " & JN(hw) & ", " & JN(hwB) & ", " & JN(-hwB) & ", " & JN(cTop) & ", " & JN(cTop) & ", " & JN(-cBot) & ", " & JN(-cBot) & vbCrLf
End Function

' Composite deck section as MCT (COMPOSITE-GEN): part 1 = girders, part 2 = slab. E values give the part ratio.
' Civil NX refuses this section type with zero stiffness lines (live 2026-11-05), so the properties are
' calculated here: A, I and centroids from the polygons; whole-section J, Asy, Asz from DeckProps (Civil NX keeps them as sent).
Private Function DeckMct() As String
    Dim s As String, d As Long
    s = "*UNIT" & vbCrLf & "   KN   , M, KJ, C" & vbCrLf & "*SECT-PSCVALUE" & vbCrLf
    For d = 1 To nDk
        UseType d
        s = s & DeckOne(d)
    Next d
    DeckMct = s & "*ENDDATA" & vbCrLf
End Function

Private Function DeckOne(ByVal d As Long) As String
    Dim s As String, j As Long, y0 As Double, nr As Double, sum2 As Double, yj As Double
    Dim a1 As Double, iy1 As Double, iz1 As Double, asy1 As Double, asz1 As Double, j1 As Double, hw As Double, hwB As Double, per1 As Double
    Dim a2 As Double, iy2 As Double, iz2 As Double, j2 As Double, per2 As Double, cz2 As Double
    Dim aT As Double, czT As Double, iyT As Double, izT As Double, st1 As String
    y0 = -(Hg + tc) / 2
    nr = eS / eG
    For j = 0 To nG - 1
        yj = (j - (nG - 1) / 2#) * sG
        sum2 = sum2 + yj ^ 2
    Next j
    a1 = nG * polyArea: iy1 = nG * gIy: iz1 = nG * gIz + polyArea * sum2
    asz1 = nG * kb(2) * Hg: asy1 = 0.94 * a1
    j1 = nG * (kb(1) * kb(7) ^ 3 + kb(2) * Hg ^ 3 * 0 + kb(3) * kb(12) ^ 3 + Hg * kb(2) ^ 3) / 3
    hw = (nG - 1) / 2# * sG + kb(1) / 2: hwB = (nG - 1) / 2# * sG + kb(3) / 2: per1 = nG * gPer
    a2 = deckW * tc: iy2 = deckW * tc ^ 3 / 12: iz2 = tc * deckW ^ 3 / 12: per2 = 2 * (deckW + tc): cz2 = Hg + tc / 2
    j2 = deckW * tc ^ 3 / 3
    aT = a1 + nr * a2
    czT = (a1 * gCz + nr * a2 * cz2) / aT
    iyT = iy1 + a1 * (gCz - czT) ^ 2 + nr * (iy2 + a2 * (cz2 - czT) ^ 2)
    izT = iz1 + nr * iz2
    st1 = StiffLines(aT, a1 + a2, tAsy(d), tAsz(d), tJ(d), iyT, izT, hw, hwB, Hg + tc - czT, czT, per2, deckW / 2, czT)
    s = s & " SECT=" & DeckSec(d) & ", COMPOSITE-GEN, USTYAPI" & IIf(d = 1, "", "_" & d) & ", CT, 0, 0, 0, 0, 0, 0, YES, NO, CP_G, YES, YES, 1" & vbCrLf & "       PART=1" & vbCrLf
    s = s & StiffLines(a1, a1, asy1, asz1, j1, iy1, iz1, hw, hwB, Hg - gCz, gCz, per1, deckW / 2, gCz) & st1
    s = s & "       1, YES, YES, " & JN(eG) & ", 0.2" & vbCrLf
    For j = 0 To nG - 1
        s = s & "       OPOLY=" & GirderPts((j - (nG - 1) / 2#) * sG, y0) & vbCrLf
    Next j
    s = s & "       PART=2" & vbCrLf & StiffLines(a2, a2, 5 / 6 * a2, a2, j2, iy2, iz2, deckW / 2, deckW / 2, tc / 2, tc / 2, per2, deckW / 2, cz2) & st1
    s = s & "       1, NO, YES, " & JN(eS) & ", 0.2" & vbCrLf
    s = s & "       OPOLY=" & JN(-deckW / 2) & ", " & JN(y0 + Hg + tc) & ", " & JN(-deckW / 2) & ", " & JN(y0 + Hg) & ", " & JN(deckW / 2) & ", " & JN(y0 + Hg) & ", " & JN(deckW / 2) & ", " & JN(y0 + Hg + tc) & vbCrLf
    DeckOne = s
End Function

' ============================================================================
'  DECK SECTION SOLVER: J, Asy, Asz of the whole composite deck (girders + slab, one
'  merged section, slab scaled by Es / Eg). Civil NX keeps the composite stiffness
'  lines as sent, so these values are calculated here from the real outline.
'  Finite volumes on half the deck (symmetry), grid lines on every section vertex:
'    torsion: div((1/n) grad phi) = -2, phi = 0 on the boundary, J = 2 int phi
'    shear:   div(n grad chi) = -n s / I, free boundary, 1 / As = int tau^2 / n
'  Two grids, Richardson extrapolation. Checked against a numpy solution (2026-11-05).
' ============================================================================
Private Sub DeckProps()
    Dim j1 As Double, y1 As Double, z1 As Double, j2 As Double, y2 As Double, z2 As Double, d As Long
    dkLog = ""
    For d = 1 To nDk
        UseType d
        Application.StatusBar = "MIDAS_KOPRU: deck section properties, type " & d & " of " & nDk & "..."
        SolveGrid SOLVE_H, j1, y1, z1
        SolveGrid SOLVE_H / 2, j2, y2, z2
        tJ(d) = 2 * j2 - j1: tAsy(d) = 2 * y2 - y1: tAsz(d) = 2 * z2 - z1
        dkLog = dkLog & tName(d) & ": " & tNG(d) & " x " & tSG(d) & " m, H " & tHg(d) & ", deck " & tW(d) & " m, joint " & tJoint(d) & " m, J " & Format(tJ(d), "0.0000") & _
                ", Asy " & Format(tAsy(d), "0.000") & ", Asz " & Format(tAsz(d), "0.000") & vbCrLf
    Next d
End Sub

' sorted grid lines through every break point, no cell wider than h
Private Function GridLines(pts() As Double, ByVal cnt As Long, ByVal h As Double) As Double()
    Dim i As Long, j As Long, t As Double, n As Long, o() As Double, c As Long, k As Long
    For i = 1 To cnt - 1
        For j = i + 1 To cnt
            If pts(j) < pts(i) Then t = pts(i): pts(i) = pts(j): pts(j) = t
        Next j
    Next i
    ReDim o(1 To 20000)
    c = 1: o(1) = pts(1)
    For i = 2 To cnt
        If pts(i) - o(c) > 0.000001 Then
            t = o(c)
            n = -Int(-(pts(i) - t) / h + 0.000000001)
            For k = 1 To n
                c = c + 1: o(c) = t + (pts(i) - t) * k / n
            Next k
        End If
    Next i
    ReDim Preserve o(1 To c)
    GridLines = o
End Function

Private Sub SolveGrid(ByVal h As Double, ByRef jOut As Double, ByRef asyOut As Double, ByRef aszOut As Double)
    Dim nr As Double, yp(1 To 7) As Double, wp(1 To 7) As Double, g() As Double, nPos As Long
    Dim k1 As Double, k2 As Double, k3 As Double, i As Long, j As Long, q As Long, w As Double, half As Double
    nr = eS / eG: half = deckW / 2
    k1 = kb(1): k2 = kb(2): k3 = kb(3)
    yp(1) = 0: yp(2) = kb(12): yp(3) = kb(12) + kb(11): yp(6) = Hg - kb(7): yp(5) = yp(6) - kb(8): yp(4) = yp(5) - kb(9): yp(7) = Hg
    wp(1) = k3 / 2: wp(2) = k3 / 2: wp(3) = k2 / 2: wp(4) = k2 / 2: wp(5) = k2 / 2 + kb(4): wp(6) = k1 / 2: wp(7) = k1 / 2
    ReDim g(0 To nG - 1)
    For i = 0 To nG - 1: g(i) = (i - (nG - 1) / 2#) * sG: Next i

    Dim px() As Double, cnt As Long, X() As Double, Y() As Double
    ReDim px(1 To 2 + 14 * nG)
    px(1) = 0: px(2) = half: cnt = 2
    For i = 0 To nG - 1
        For q = 1 To 7
            If g(i) + wp(q) >= 0 And g(i) + wp(q) <= half Then cnt = cnt + 1: px(cnt) = g(i) + wp(q)
            If g(i) - wp(q) >= 0 And g(i) - wp(q) <= half Then cnt = cnt + 1: px(cnt) = g(i) - wp(q)
        Next q
    Next i
    X = GridLines(px, cnt, h)
    ReDim px(1 To 8)
    For q = 1 To 7: px(q) = yp(q): Next q
    px(8) = Hg + tc
    Y = GridLines(px, 8, h)

    Dim nx As Long, ny As Long, nn() As Double, id() As Long, m As Long, xc As Double, yc As Double, dmin As Double
    nx = UBound(X) - 1: ny = UBound(Y) - 1
    ReDim nn(1 To nx, 1 To ny): ReDim id(0 To nx + 1, 0 To ny + 1)
    For j = 1 To ny
        yc = (Y(j) + Y(j + 1)) / 2
        If yc > Hg Then
            For i = 1 To nx: nn(i, j) = nr: Next i
        Else
            For q = 1 To 6
                If yc <= yp(q + 1) Then Exit For
            Next q
            If yp(q + 1) - yp(q) > 0 Then w = wp(q) + (wp(q + 1) - wp(q)) * (yc - yp(q)) / (yp(q + 1) - yp(q)) Else w = wp(q)
            For i = 1 To nx
                xc = (X(i) + X(i + 1)) / 2
                dmin = 1E+30
                For q = 0 To nG - 1
                    If Abs(xc - g(q)) < dmin Then dmin = Abs(xc - g(q))
                Next q
                If dmin <= w Then nn(i, j) = 1
            Next i
        End If
    Next j
    For j = 1 To ny
        For i = 1 To nx
            If nn(i, j) > 0 Then m = m + 1: id(i, j) = m
        Next i
    Next j

    Dim nb() As Long, gT() As Double, gS() As Double, dT() As Double, dS() As Double, dA() As Double
    Dim ar() As Double, cx() As Double, cy() As Double, cn() As Double
    ReDim nb(1 To 4, 1 To m): ReDim gT(1 To 4, 1 To m): ReDim gS(1 To 4, 1 To m)
    ReDim dT(1 To m): ReDim dS(1 To m): ReDim dA(1 To m): ReDim ar(1 To m): ReDim cx(1 To m): ReDim cy(1 To m): ReDim cn(1 To m)
    Dim p As Long, a As Long, b As Long, f As Long, L As Double, dp As Double, dq As Double, n1 As Double, n2 As Double, ct As Double, cs As Double
    Dim di As Variant, dj As Variant
    di = Array(1, -1, 0, 0): dj = Array(0, 0, 1, -1)
    For j = 1 To ny
        For i = 1 To nx
            p = id(i, j)
            If p > 0 Then
                n1 = nn(i, j)
                ar(p) = (X(i + 1) - X(i)) * (Y(j + 1) - Y(j))
                cx(p) = (X(i) + X(i + 1)) / 2: cy(p) = (Y(j) + Y(j + 1)) / 2: cn(p) = n1
                For f = 0 To 3
                    a = i + di(f): b = j + dj(f)
                    If di(f) <> 0 Then
                        L = Y(j + 1) - Y(j): dp = (X(i + 1) - X(i)) / 2
                    Else
                        L = X(i + 1) - X(i): dp = (Y(j + 1) - Y(j)) / 2
                    End If
                    If id(a, b) > 0 Then
                        n2 = nn(a, b)
                        If di(f) <> 0 Then dq = (X(a + 1) - X(a)) / 2 Else dq = (Y(b + 1) - Y(b)) / 2
                        ct = L / (dp * n1 + dq * n2)
                        cs = L / (dp / n1 + dq / n2)
                        nb(f + 1, p) = id(a, b): gT(f + 1, p) = ct: gS(f + 1, p) = cs
                        dT(p) = dT(p) + ct: dS(p) = dS(p) + cs: dA(p) = dA(p) + cs
                    ElseIf a = 0 Then
                        dA(p) = dA(p) + L * n1 / dp          ' centre line: Asy is antisymmetric (chi = 0)
                    Else
                        dT(p) = dT(p) + L / (n1 * dp)        ' free boundary: phi = 0
                    End If
                Next f
            End If
        Next i
    Next j

    Dim sA As Double, sZ As Double, cz As Double, iy As Double, iz As Double, rhs() As Double, sol() As Double, e As Double
    ReDim rhs(1 To m): ReDim sol(1 To m)
    For p = 1 To m
        sA = sA + cn(p) * ar(p): sZ = sZ + cn(p) * ar(p) * cy(p)
    Next p
    cz = sZ / sA
    For p = 1 To m
        iy = iy + 2 * cn(p) * ar(p) * (cy(p) - cz) ^ 2
        iz = iz + 2 * cn(p) * ar(p) * cx(p) ^ 2
    Next p

    For p = 1 To m: rhs(p) = 2 * ar(p): Next p
    Pcg nb, gT, dT, rhs, sol, m
    e = 0
    For p = 1 To m: e = e + sol(p) * ar(p): Next p
    jOut = 4 * e

    For p = 1 To m: rhs(p) = cn(p) * (cy(p) - cz) / iy * ar(p): sol(p) = 0: Next p
    Pcg nb, gS, dS, rhs, sol, m
    e = 0
    For p = 1 To m: e = e + sol(p) * rhs(p): Next p
    aszOut = 1 / (2 * e)

    For p = 1 To m: rhs(p) = cn(p) * cx(p) / iz * ar(p): sol(p) = 0: Next p
    Pcg nb, gS, dA, rhs, sol, m
    e = 0
    For p = 1 To m: e = e + sol(p) * rhs(p): Next p
    asyOut = 1 / (2 * e)
End Sub

' conjugate gradients, Jacobi preconditioner; matrix = diag d, off-diagonals -g to neighbours nb
Private Sub Pcg(nb() As Long, g() As Double, d() As Double, b() As Double, x() As Double, ByVal m As Long)
    Dim r() As Double, z() As Double, p() As Double, ap() As Double, i As Long, k As Long, it As Long
    Dim rz As Double, rzNew As Double, pap As Double, al As Double, be As Double, b2 As Double, r2 As Double, s As Double
    ReDim r(1 To m): ReDim z(1 To m): ReDim p(1 To m): ReDim ap(1 To m)
    For i = 1 To m
        r(i) = b(i): z(i) = r(i) / d(i): p(i) = z(i)
        rz = rz + r(i) * z(i): b2 = b2 + b(i) * b(i)
    Next i
    For it = 1 To 30000
        pap = 0
        For i = 1 To m
            s = d(i) * p(i)
            For k = 1 To 4
                If nb(k, i) > 0 Then s = s - g(k, i) * p(nb(k, i))
            Next k
            ap(i) = s: pap = pap + p(i) * s
        Next i
        If pap <= 0 Then Exit For
        al = rz / pap
        r2 = 0: rzNew = 0
        For i = 1 To m
            x(i) = x(i) + al * p(i)
            r(i) = r(i) - al * ap(i)
            z(i) = r(i) / d(i)
            r2 = r2 + r(i) * r(i): rzNew = rzNew + r(i) * z(i)
        Next i
        If r2 < 0.0000000000000001 * b2 Then Exit For
        be = rzNew / rz: rz = rzNew
        For i = 1 To m: p(i) = z(i) + be * p(i): Next i
        If it Mod 200 = 0 Then DoEvents
    Next it
    If it > 30000 Then Err.Raise vbObjectError + 40, , "Deck section solver did not converge."
End Sub

Private Function SectHead2(ByVal pt As String, ByVal vOpt As Long, ByVal vOff As Double) As String
    SectHead2 = """OFFSET_PT"":" & QS(pt) & ",""OFFSET_CENTER"":0,""USER_OFFSET_REF"":0,""HORZ_OFFSET_OPT"":0,""USERDEF_OFFSET_YI"":0,""VERT_OFFSET_OPT"":" & vOpt & ",""USERDEF_OFFSET_ZI"":" & JN(vOff) & ",""USE_SHEAR_DEFORM"":true,""USE_WARPING_EFFECT"":false"
End Function

Private Function DbSect(ByVal no As Long, ByVal nm As String, ByVal head As String, ByVal shp As String, ByVal sz As String) As String
    DbSect = QS(CStr(no)) & ":{""SECTTYPE"":""DBUSER"",""SECT_NAME"":" & QS(nm) & ",""SECT_BEFORE"":{" & head & ",""SHAPE"":" & QS(shp) & ",""DATATYPE"":2,""SECT_I"":{""vSIZE"":[" & sz & "]}}}"
End Function

' RTI: the MIDAS EN04(RC) database = Eurocode Ecm. Static and seismic: user concrete with the AASHTO Ec of the
' workbook (MPa -> kN/m2); Poisson, thermal coefficient and unit weight as the database, only E changes
Private Function MatJson(ByVal no As Long, ByVal cls As String, ByVal eMpa As Double, ByVal den As Double) As String
    Dim prm As String
    If rtiMode Then
        prm = "{""P_TYPE"":1,""STANDARD"":""EN04(RC)"",""CODE"":"""",""DB"":" & QS(EnDb(cls)) & "}"
    Else
        prm = "{""P_TYPE"":2,""ELAST"":" & JN(eMpa * 1000#) & ",""POISN"":0.2,""THERMAL"":0.00001,""DEN"":" & JN(den) & ",""MASS"":0}"
    End If
    MatJson = QS(CStr(no)) & ":{""TYPE"":""CONC"",""NAME"":" & QS(cls) & ",""HE_SPEC"":0,""HE_COND"":0,""bMASS_DENS"":false,""DAMP_RAT"":0.05,""PARAM"":[" & prm & "]}"
End Function

Private Sub BuildPayloads()
    Dim s As String, i As Long, sz As String
    Set payloads = New Collection
    Add "db/UNIT", "{""Assign"":{""1"":{""FORCE"":""KN"",""DIST"":""M"",""HEAT"":""J"",""TEMPER"":""C""}}}"
    Add "db/STYP", "{""Assign"":{""1"":{""STYP"":0,""MASS"":1,""GRAV"":" & JN(GRAV_STYP) & ",""TEMP"":0,""bALIGNBEAM"":false,""bALIGNSLAB"":false,""bMASSOFFSET"":false,""bSELFWEIGHT"":true,""SMASS"":1}}}"
    Add "db/PJCF", "{""Assign"":{""1"":{""PROJECT"":" & QS(Pj(projName)) & ",""USER"":" & QS(Pj(Application.UserName)) & ",""CLIENT"":" & QS(Pj(clientName)) & _
        ",""TITLE"":" & QS(Pj(bridgeName)) & ",""COMMENT"":" & QS(Pj(nSp & " spans; " & SCRIPT_ID & " " & SCRIPT_VERSION & IIf(seisMode, " SEISMIC", IIf(rtiMode, " RTI", "")))) & "}}}"
    Dim sMat As String
    If rtiMode Then sMat = "," & RailMat()
    Add "db/MATL", "{""Assign"":{" & MatJson(1, clsG, eG, denG) & "," & MatJson(2, clsS, eS, denS) & "," & MatJson(3, clsO, eO, denS) & sMat & "}}"
    s = ""
    For i = 1 To nDk
        s = s & DbSect(SlabSec(i), "DOSEME" & IIf(i = 1, "", "_" & i), SectHead2("CT", 0, 0), "SB", JN(tc) & "," & JN(tW(i)) & ",0,0,0,0,0,0,0,0") & ","
    Next i
    ' cap nodes sit at the cap centroid (owner 2026-11-09), so the section is centred on them
    If capRect Then
        s = s & DbSect(SEC_CAP, "BASLIK", SectHead2("CC", 0, 0), "SB", JN(capS(1)) & "," & JN(capS(4)) & ",0,0,0,0,0,0,0,0")
    Else
        s = s & DbSect(SEC_CAP, "BASLIK", SectHead2("CC", 0, 0), "UDT", JN(capS(1)) & "," & JN(capS(2)) & "," & JN(capS(3)) & "," & JN(capS(4)) & "," & JN(capS(5)) & ",0,0,0,0,0")
    End If
    For i = 1 To nPS
        If psTyp(i) = "SR" Then
            sz = JN(psEn(i)) & ",0,0,0,0,0,0,0,0,0"
        ElseIf psTyp(i) = "B" Then
            sz = JN(psBoy(i)) & "," & JN(psEn(i)) & "," & JN(psT(i)) & "," & JN(psT(i)) & "," & JN(psEn(i) - psT(i)) & "," & JN(psT(i)) & ",0,0,0,0"
        Else
            sz = JN(psBoy(i)) & "," & JN(psEn(i)) & ",0,0,0,0,0,0,0,0"
        End If
        s = s & "," & DbSect(SEC_PIER0 + i - 1, "KOLON_" & i, SectHead2("CC", 0, 0), psTyp(i), sz)
    Next i
    If rtiMode Then s = s & "," & RailSect()
    Add "db/SECT", "{""Assign"":{" & s & "}}"
    Add "db/CO_S", SectColorJson()
    s = ""
    For i = 1 To nCnt
        s = s & IIf(i = 1, "", ",") & QS(CStr(i)) & ":{""X"":" & JN(nX(i)) & ",""Y"":" & JN(nY(i)) & ",""Z"":" & JN(nZ(i)) & "}"
    Next i
    Add "db/NODE", "{""Assign"":{" & s & "}}"
    s = ""
    For i = 1 To eCnt
        s = s & IIf(i = 1, "", ",") & QS(CStr(i)) & ":{""TYPE"":""BEAM"",""MATL"":" & eMat(i) & ",""SECT"":" & eSec(i) & _
            ",""NODE"":[" & eN1(i) & "," & eN2(i) & ",0,0,0,0,0,0],""ANGLE"":" & JN(IIf(pierAng.Exists(CStr(i)), pierAng(CStr(i)), 0)) & ",""STYPE"":0}"
    Next i
    Add "db/ELEM", "{""Assign"":{" & s & "}}"
    Dim sAll As String
    If rtiMode Then
        AddRailGeometry
        sAll = "," & GrupAll()
    End If
    Add "db/GRUP", "{""Assign"":{""1"":{""NAME"":""SG_TABLIYE"",""P_TYPE"":0,""E_LIST"":[" & grpDeck & "]}," & _
        """2"":{""NAME"":""SG_BASLIK"",""P_TYPE"":0,""E_LIST"":[" & grpCap & "]},""3"":{""NAME"":""SG_KOLON"",""P_TYPE"":0,""E_LIST"":[" & grpPier & "]}" & sAll & "}}"
    If rtiMode Then
        Add "db/BNGR", "{""Assign"":{" & BngrRti() & "}}"
    Else
        Add "db/BNGR", "{""Assign"":{""1"":{""NAME"":""BG_MESNET"",""AUTOTYPE"":0}}}"
    End If
    Add "db/CONS", "{""Assign"":{" & consJson & "}}"
    Add "db/RIGD", "{""Assign"":{" & RigdAll() & "}}"
    If rtiMode Then AddRailLinks
    Add "db/ELNK", "{""Assign"":{" & elnkJson & "}}"
    AddLoadPayloads
End Sub

Private Sub WritePayloads(ByVal dirPath As String)
    Dim i As Long, p As Variant
    MakeDir dirPath
    For i = 1 To payloads.Count
        p = payloads(i)
        WriteText dirPath & "\" & Format(i, "00") & "_" & Replace(p(0), "db/", "") & ".json", CStr(p(1))
    Next i
    WriteText dirPath & "\00_SECT_deck.mct", DeckMct()
    WriteText dirPath & "\summary.txt", Summary()
End Sub

Private Function SendAll(Optional ByVal saveModel As Boolean = True) As String
    Dim rep As String, i As Long, p As Variant, resp As String, t0 As Single, cleanRep As String, mctPath As String, a As Double
    t0 = Timer
    If Not GetKey() Then SendAll = "FAIL: no MAPI key.": Exit Function
    Set http = Nothing
    If Not Api("GET", "db/NODE", "", resp) Then SendAll = "FAIL: Civil NX does not answer (db/NODE): " & Left$(resp, 300): Exit Function
    If Not ModelOk(rep) Then SendAll = rep: Exit Function
    If Not ModelTypeOk(rep) Then SendAll = rep: Exit Function
    rep = ""
    ' an empty model needs no clean pass (owner 2026-11-05)
    If ModelIsEmpty() Then
        ProgReset 1 + payloads.Count + 4
        cleanRep = "Clean pass skipped: the open model is empty."
    Else
        ProgReset 2 * (UBound(CleanTables()) + 1) + 1 + payloads.Count + 4
        cleanRep = CleanModel()
    End If
    If Left$(cleanRep, 7) = "STOPPED" Then SendAll = cleanRep & vbCrLf & "Nothing was written.": Exit Function
    ' doc/IMPORTMXT replaces the whole model content, so it runs before every JSON table
    mctPath = Environ$("TEMP") & "\midas_kopru_sect.mct"
    WriteText mctPath, DeckMct()
    Prog "deck section (MCT import)"
    If Not Api("POST", "doc/IMPORTMXT", "{""Argument"":" & QS(mctPath) & "}", resp) Then
        SendAll = cleanRep & vbCrLf & "FAIL: MCT import of the deck section (" & mctPath & "): " & Left$(resp, 300) & vbCrLf & "Nothing else was written."
        Exit Function
    End If
    If Not Api("GET", "db/SECT", "", resp) Then resp = ""
    If InStr(resp, "USTYAPI") = 0 Then SendAll = cleanRep & vbCrLf & "FAIL: deck section USTYAPI is not in the model after the MCT import. Nothing else was written.": Exit Function
    For i = 1 To payloads.Count
        p = payloads(i)
        Prog "table " & i & "/" & payloads.Count & "  " & p(0) & "  " & Format$(Len(CStr(p(1))) / 1024#, "0") & " kB"
        If Not Api("PUT", CStr(p(0)), CStr(p(1)), resp) Then
            SendAll = cleanRep & vbCrLf & "FAIL at step " & i & " (" & p(0) & "):" & vbCrLf & Left$(resp, 600) & vbCrLf & vbCrLf & "Model is incomplete; nothing saved."
            Exit Function
        End If
    Next i
    rep = "Read-back:" & vbCrLf
    Prog "read-back nodes"
    If Api("GET", "db/NODE", "", resp) Then rep = rep & "  nodes " & CountOf(resp, """X"":") & " / " & nCnt + xNo & IIf(CountOf(resp, """X"":") = nCnt + xNo, "", "  FAIL") & vbCrLf
    Prog "read-back elements"
    If Api("GET", "db/ELEM", "", resp) Then rep = rep & "  elements " & CountOf(resp, """TYPE"":") & " / " & eCnt + reCnt & IIf(CountOf(resp, """TYPE"":") = eCnt + reCnt, "", "  FAIL") & vbCrLf
    Prog "read-back sections"
    If Api("GET", "db/STLD", "", resp) Then rep = rep & "  load cases " & CountOf(resp, """NAME"":") & " / " & CaseList().Count & IIf(CountOf(resp, """NAME"":") = CaseList().Count, "", "  FAIL") & vbCrLf
    If Api("GET", "db/LCOM-GEN", "", resp) Then rep = rep & "  combinations " & CountOf(resp, """iTYPE"":") & " / " & combNo & IIf(CountOf(resp, """iTYPE"":") = combNo, "", "  FAIL") & vbCrLf
    If Api("GET", "db/SECT", "", resp) Then
        UseType 1
        a = SectArea(resp, CStr(DeckSec(1)))
        rep = rep & "  deck section area read " & Format(a, "0.0000") & " m2 (girders " & Format(nG * polyArea, "0.0000") & ", transformed " & Format(deckArea, "0.0000") & ")" & IIf(a > 0, "", "  FAIL: no stiffness in the composite section") & vbCrLf
    End If
    If seisMode Then rep = rep & SecfCheck()
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
    SendAll = cleanRep & vbCrLf & "All " & payloads.Count & " tables sent (" & Format(Timer - t0, "0") & " s)." & vbCrLf & rep & vbCrLf & SCRIPT_ID & " " & SCRIPT_VERSION & ", " & bridgeName
End Function

Private Function ModelIsEmpty() As Boolean
    Dim t As Variant, resp As String
    For Each t In Array("NODE", "ELEM", "SECT", "MATL", "STLD")
        If Not Api("GET", "db/" & t, "", resp) Then Exit Function
        If TableHasRecords(resp, CStr(t)) Then Exit Function
    Next t
    ModelIsEmpty = True
End Function

' do not wipe a superstructure model by mistake
Private Function ModelOk(ByRef msg As String) As Boolean
    Dim p As String, txt As String
    ModelOk = True
    If Not Api("GET", "db/PJCF", "", p) Then Exit Function
    If InStr(p, "midas-ustyapi-build") = 0 Then Exit Function
    txt = "The model open in Civil NX is a superstructure (MIDAS_USTYAPI) model."
    If askUser Then
        If UiMsg(txt & vbCrLf & vbCrLf & "WIPE it anyway?", vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) = vbYes Then Exit Function
    End If
    msg = "STOPPED: " & txt & " Nothing was changed. Open the right .mcb, or use an empty file."
    ModelOk = False
End Function

Private Function AxisBearingText() As String
    Dim i As Long, mn As Double, mx As Double
    If nSp < 2 Then AxisBearingText = "-": Exit Function
    mn = ePier(1): mx = ePier(1)
    For i = 2 To nSp - 1
        If ePier(i) < mn Then mn = ePier(i)
        If ePier(i) > mx Then mx = ePier(i)
    Next i
    AxisBearingText = Format(mn, "0.000") & IIf(mx - mn > 0.0005, " - " & Format(mx, "0.000"), "") & " m (calculated)"
End Function

Private Function Summary() As String
    Dim s As String, i As Long, f As Long, tot As Double
    For i = 1 To nSp
        tot = tot + spL(i)
        If fixLeft(i) Then f = f + 1
    Next i
    s = SCRIPT_ID & " " & SCRIPT_VERSION & vbCrLf & bridgeName & ": " & nSp & " spans, length " & Format(tot, "0.00") & " m, " & nSp - 1 & " piers" & vbCrLf & _
        dkLog & "Slab " & Format(tc, "0.000") & " m, sarkiklik " & sark & " m; concrete " & clsG & " / " & clsS & " / " & clsO & IIf(rtiMode, " (MIDAS EN04 database, Ecm)", " (AASHTO Ec " & Format(eG, "0") & " / " & Format(eS, "0") & " / " & Format(eO, "0") & " MPa)") & vbCrLf & _
        "Bearings: kv " & Format(kEB, "0") & ", free ky " & Format(kyB, "0.0") & ", kh " & Format(kFix, "0") & ", link " & linkLen & " m; " & f & " spans fixed left, " & nSp - f & " right" & vbCrLf & _
        "Cap " & IIf(capRect, "rectangle " & capS(1) & " x " & capS(4), "inverted T H " & capS(1)) & ", length " & capL & " m, centroid " & Format(capYc, "0.000") & " m below its top, pier top at the cap bottom; " & nPS & " pier section(s)" & vbCrLf & _
        "Axis to bearing " & AxisBearingText() & IIf(capRect, "", ", seat on the " & IIf(seatOnLedge(1), "ledges", "web")) & vbCrLf & _
        "Model: " & nCnt & " nodes, " & eCnt & " elements, " & nLink & " bearings, element " & meshDx & " m" & vbCrLf & _
        IIf(nAxXY > 0, "Plan: " & nAxXY & " axis points from G_AKS_XY (DXF), largest offset from the A1-A2 chord " & Format(maxOff, "0.00") & " m", "Plan: straight") & vbCrLf & SkewText() & vbCrLf & _
        IIf(seisMode, "SEISMIC MODEL: " & CaseList().Count & " cases, " & combNo & " combinations; train " & Format(qvk, "0") & " kN/m static, reduced stiffness in EVERY case", "Loads: " & CaseList().Count & " cases, " & combNo & " combinations, " & nFW & " braking windows; rail level " & Format(railH, "0.00") & " m; alpha " & alphaLM & ", Wg " & wG & ", Wv " & wV & ", TU " & tuPos & "/" & tuNeg) & vbCrLf & _
        IIf(seisMode, "EQ " & EQ_FUNC & ", Ritz " & RITZ_X & "/" & RITZ_Y & "/" & RITZ_Z, IIf(hasCF, "Centrifugal: " & nCfSpans & " spans on a curve, up to " & Format(cfMax, "0.0") & " kN/m", "Centrifugal: none") & "; static model: no eigenvalue, spectra or seismic stiffness")
    If warnText <> "" Then s = "WARNINGS:" & vbCrLf & warnText & vbCrLf & s
    If rtiMode Then s = s & vbCrLf & "RTI: " & nTrk & " rail line(s), " & rN(1) & " rail nodes per line, " & nWin & " train windows + " & IIf(rtiNL, nLL & " nonlinear live load stages + ", "") & "TU+, TU-, END; " & bgUsed.Count & " spring groups, " & lgUsed.Count & " load groups; embankment " & RTI_EMB & " m"
    Summary = s
End Function

' ============================================================================
'  STEP B: LOADS, SPECTRA, COMBINATIONS (owner 2026-11-06)
' ============================================================================
Private Sub ReadLoadData(wb As Workbook)
    Dim y As Worksheet, k As Long
    Set y = wb.Worksheets("2_YUKLER")
    wG = Num(BelowHeader(y, "Wg").Value, "2_YUKLER Wg (wind on structure)")
    wV = Num(BelowHeader(y, "Wv").Value, "2_YUKLER Wv (wind on train)")
    alphaLM = Num(BelowHeader(y, ChrW(945)).Value, "2_YUKLER alpha")
    qsk = Num(LabelValue(y, 7, "Qsk", 16), "2_YUKLER Qsk x alpha (nosing)")
    tuPos = Num(NmRange(wb, "G_SICAKLIK_MAKS").Cells(1).Value, "G_SICAKLIK_MAKS")
    tuNeg = Num(NmRange(wb, "G_SICAKLIK_MIN").Cells(1).Value, "G_SICAKLIK_MIN")
    vDes = Num(NmRange(wb, "G_TASARIM_HIZI").Cells(1).Value, "G_TASARIM_HIZI")
    qvk = Num(NmRange(wb, "G_HAT_DUSEY_YUKU").Cells(1).Value, "G_HAT_DUSEY_YUKU")
    ' rail level above the slab top: no input cell (owner 2026-11-06), ballast + protection concrete + sleeper and rail
    railH = Num(NmRange(wb, "G_BALAST_YUKSEKLIGI").Cells(1).Value, "G_BALAST_YUKSEKLIGI") + _
            Num(NmRange(wb, "G_KORUMA_BETONU_YUKSEKLIGI").Cells(1).Value, "G_KORUMA_BETONU_YUKSEKLIGI") + RAIL_OVER_BALLAST
    If NumOr0(NmRange(wb, "G_KOPRU_ONEM_SINIFI").Cells(1).Value) <> 1 Then warnText = warnText & "- KOS is not 1: EQ cases still use " & EQ_FUNC & "." & vbCrLf
    ReadCombTables wb.Worksheets("1_GIRIS")
    If seisMode Then ReadSpectra wb.Worksheets("SPECTRUM")
    If rtiMode Then
        If InStr(UCase$(CStr(NmRange(wb, "G_HAT_TIPI").Cells(1).Value)), "SLAB") > 0 Then Err.Raise vbObjectError + 61, , "RTI: the track law is set for ballast only. Slab-track values are not defined yet."
    End If
    phiByElem = False
    For k = 2 To nDk
        If tPhi(k) <> tPhi(1) Then phiByElem = True
    Next k
End Sub

' every curve of the SPECTRUM sheet: name above 'T', period column, Sae(T) column
Private Sub ReadSpectra(ws As Worksheet)
    Dim c As Range, hr As Long, col As Long, r As Long, s As String, prev As Double, t As Double, nP As Long
    Set c = ws.UsedRange.Find(What:="Sae(T)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
    If c Is Nothing Then Err.Raise vbObjectError + 50, , "SPECTRUM: header Sae(T) not found."
    hr = c.Row
    Set specNames = New Collection: Set specJson = New Collection
    For col = 1 To ws.UsedRange.Columns.Count + ws.UsedRange.Column
        If Trim$(CStr(ws.Cells(hr, col).Value)) = "T" And Trim$(CStr(ws.Cells(hr, col + 1).Value)) = "Sae(T)" Then
            s = "": prev = -1: nP = 0
            r = hr + 1
            Do While Not IsEmpty(ws.Cells(r, col).Value)
                t = Num(ws.Cells(r, col).Value, "SPECTRUM period")
                If t > prev Then
                    s = s & IIf(s = "", "", ",") & "{""PERIOD"":" & JN(t) & ",""VALUE"":" & JN(Num(ws.Cells(r, col + 1).Value, "SPECTRUM value")) & "}"
                    prev = t: nP = nP + 1
                End If
                r = r + 1
            Loop
            If nP < 5 Then Err.Raise vbObjectError + 51, , "SPECTRUM: curve at column " & col & " has " & nP & " points."
            specNames.Add Trim$(CStr(ws.Cells(hr - 1, col).Value))
            specJson.Add s
        End If
    Next col
    For r = 1 To specNames.Count
        If specNames(r) = EQ_FUNC Then Exit Sub
    Next r
    Err.Raise vbObjectError + 52, , "SPECTRUM: curve " & EQ_FUNC & " not found."
End Sub

Private Sub ReadCombTables(ws As Worksheet)
    Dim hc As Range, hr As Long, c As Long, t As String, cDC As Long, cLL As Long, cWA As Long, cWS As Long, cWL As Long, cTU As Long, cEQ As Long
    Dim r As Long, lab As String, g As Range, v As String
    Set hc = FindLabel(ws, 3, "Y" & ChrW(252) & "k Komb")
    hr = hc.Row
    For c = 4 To 40
        t = Trim$(CStr(ws.Cells(hr, c).Value))
        If Left$(t, 2) = "DC" And cDC = 0 Then cDC = c
        If Left$(t, 2) = "LL" And cLL = 0 Then cLL = c
        If Left$(t, 2) = "WA" And cWA = 0 Then cWA = c
        If t = "WS" Then cWS = c
        If t = "WL" Then cWL = c
        If t = "TU" Then cTU = c
        If t = "EQ" Then cEQ = c
    Next c
    If cDC * cLL * cWA * cWS * cWL * cTU * cEQ = 0 Then Err.Raise vbObjectError + 8, , "Tablo 14 header columns (DC, LL, WA, WS, WL, TU, EQ) not all found."
    Set facRows = CreateObject("Scripting.Dictionary")
    facRows.CompareMode = vbBinaryCompare
    For r = hr + 1 To hr + 20
        lab = Replace(Trim$(CStr(ws.Cells(r, 3).Value)), " ", "")
        If lab <> "" Then facRows(lab) = Array(Fac(ws.Cells(r, cDC).Value), Fac(ws.Cells(r, cLL).Value), Fac(ws.Cells(r, cWA).Value), _
                                             Fac(ws.Cells(r, cWS).Value), Fac(ws.Cells(r, cWL).Value), Fac(ws.Cells(r, cTU).Value))
        If lab = "TKDY" Then
            v = Trim$(CStr(ws.Cells(r, cEQ).Value))
            If InStr(v, "/") = 0 Then Err.Raise vbObjectError + 53, , "Tablo 14 TKDY row: EQ must read like 1.00 / 0.30."
            eqMain = Val(Replace(Left$(v, InStr(v, "/") - 1), ",", "."))
            eqOther = Val(Replace(Mid$(v, InStr(v, "/") + 1), ",", "."))
        End If
    Next r
    If Not facRows.Exists("TKDY") Then Err.Raise vbObjectError + 10, , "Tablo 14 row 'TKDY' not found."
    Set g = FindLabel(ws, 7, "DC", hr)
    dcMax = Num(ws.Cells(g.Row, 14).Value, "Tablo 15 DC max"): dcMin = Num(ws.Cells(g.Row, 17).Value, "Tablo 15 DC min")
    dcMax4 = Num(ws.Cells(g.Row + 1, 14).Value, "Tablo 15 DC max STR-4")
    Set g = FindLabel(ws, 7, "DW", hr)
    dwMax = Num(ws.Cells(g.Row, 14).Value, "Tablo 15 DW max"): dwMin = Num(ws.Cells(g.Row, 17).Value, "Tablo 15 DW min")
End Sub

Private Function Fac(ByVal v As Variant) As Double
    Dim t As String
    If IsNumeric(v) And Not IsEmpty(v) Then Fac = CDbl(v): Exit Function
    t = LCase$(Trim$(CStr(v)))
    If t = "" Or Left$(t, 2) = "--" Or t = "-" Then Fac = -1: Exit Function
    If t = "gp" Or t = ChrW(947) & "p" Then Fac = -2: Exit Function
    If InStr(t, "/") > 0 Then t = Trim$(Left$(t, InStr(t, "/") - 1))
    Fac = Val(Replace(t, ",", "."))
End Function

Private Function BelowHeader(ws As Worksheet, ByVal header As String) As Range
    ' first cell equal to header whose cell below holds a number
    Dim c As Range, first As String
    Set c = ws.UsedRange.Find(What:=header, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
    If Not c Is Nothing Then
        first = c.Address
        Do
            If Not IsError(c.Offset(1, 0).Value) Then
                If IsNumeric(c.Offset(1, 0).Value) And Not IsEmpty(c.Offset(1, 0).Value) Then Set BelowHeader = c.Offset(1, 0): Exit Function
            End If
            Set c = ws.UsedRange.FindNext(c)
            If c Is Nothing Then Exit Do
        Loop While c.Address <> first
    End If
    Err.Raise vbObjectError + 54, , ws.Name & ": no number below the header '" & header & "'."
End Function

' axis name left of span i (the name of its IVME case)
Private Function AxisName(ByVal i As Long) As String
    AxisName = IIf(i = 1, "A1", "P" & (i - 1))
End Function

' centrifugal force per track and metre: same formula as 2_YUKLER Tablo 17 (L180, U180), r = span radius, Lf = span
Private Function CentForce(ByVal r As Double, ByVal Lf As Double) As Double
    Dim v As Double, f As Double, a As Double, b As Double, v2 As Double
    v = IIf(vDes > 300, 300, vDes)
    If vDes <= 120 Or Lf <= 2.88 Then
        f = 1
    Else
        f = 1 - (v - 120) * (814 / v + 1.75) * (1 - Sqr(2.88 / Lf)) / 1000
        If f < 0.35 Then f = 0.35
    End If
    v2 = IIf(vDes < 120, vDes, 120)
    a = vDes ^ 2 / (127 * r) * f
    b = v2 ^ 2 / (127 * r)
    CentForce = IIf(a > b, a, b) * alphaLM * qvk
End Function

Private Sub BuildLoads()
    If rtiMode Then
        hasCF = False: nCfSpans = 0: cfMax = 0: nFW = 0: combNo = 0
        BuildRail
        Exit Sub
    End If
    xNo = 0: reCnt = 0
    Dim i As Long, e As Long, d As Long, n As Double, w As Double, qcf As Double, k As Long, m As Long
    Dim pos As Double, ctr As Double, nW As Long, a As Long, b As Long, lenM As Double
    Set loadItems = CreateObject("Scripting.Dictionary")
    Set loadIds = CreateObject("Scripting.Dictionary")
    hasCF = False: nCfSpans = 0: cfMax = 0
    For i = 1 To nSp
        d = spType(i): n = tNG(d): w = tW(d)
        qcf = 0
        If spR(i) <> 0 And Not seisMode Then
            qcf = CentForce(Abs(spR(i)), tLc(d)) * NTracks(d) * Sgn(spR(i))   ' R > 0 = right turn, outward = +LY
            hasCF = True: nCfSpans = nCfSpans + 1
            If Abs(qcf) > cfMax Then cfMax = Abs(qcf)
        End If
        For e = dkA(i) To dkB(i)
            AddLoad e, "DW_1", "GZ", -1, "LY", 0, -tQ(d, Q_BAL) * n
            AddLoad e, "DW_1", "GZ", 1, "LY", w / 2 - tKL(d) / 2, -tQ(d, Q_KERB) * n / 2
            AddLoad e, "DW_1", "GZ", 1, "LY", -(w / 2 - tKR(d) / 2), -tQ(d, Q_KERB) * n / 2
            AddLoad e, "DW_1", "GZ", 1, "LY", w / 2, -(tQ(d, Q_RAIL) + tQ(d, Q_PANEL)) * n / 2
            AddLoad e, "DW_1", "GZ", 1, "LY", -w / 2, -(tQ(d, Q_RAIL) + tQ(d, Q_PANEL)) * n / 2
            AddLoad e, "DW_2", "GZ", -1, "LY", 0, -(tQ(d, Q_PROT) + tQ(d, Q_TRK)) * n
            If seisMode Then
                AddLoad e, MASS_CASE, "GZ", -1, "LY", 0, -qvk * NTracks(d)
            Else
                AddLoad e, "WS-Y", "LY", -1, "LY", 0, wG
                AddLoad e, "WL", "LY", 1, "LZ", railH + WL_HEIGHT, wV
            End If
            If qcf <> 0 Then AddLoad e, "CF", "LY", 1, "LZ", railH + CF_HEIGHT, qcf
        Next e
        If seisMode Then GoTo NextSpan
        ' traction: the bearing-to-bearing elements, at most TRAC_L centred on the span
        m = mnB(i) - mnA(i) + 1
        ctr = m * dxS(i) / 2: nW = 0
        For k = 1 To m
            pos = (k - 0.5) * dxS(i)
            If Abs(pos - ctr) + dxS(i) / 2 <= TRAC_L / 2 + 0.05 Then
                AddLoad mnA(i) + k - 1, "IVME " & AxisName(i), "LX", 1, "LZ", railH, TRAC_Q * alphaLM
                nW = nW + 1
            End If
        Next k
        If nW = 0 Then Err.Raise vbObjectError + 55, , "Span " & i & ": element longer than the " & TRAC_L & " m traction length."
        ' nosing: spread over the middle element(s) of the span
        If m Mod 2 = 0 Then
            a = mnA(i) + m / 2 - 1: b = a + 1
        Else
            a = mnA(i) + (m - 1) / 2: b = a
        End If
        lenM = (b - a + 1) * dxS(i)
        For e = a To b
            AddLoad e, "MOVE", "LY", 1, "LZ", railH, qsk / lenM
        Next e
NextSpan:
    Next i
    If seisMode Then nFW = 0: Exit Sub
    BrakeWindows
    For k = 1 To nFW
        For i = fwA(k) To fwB(k)
            For e = dkA(i) To dkB(i)
                AddLoad e, "FREN " & k, "LX", 1, "LZ", railH, BRAKE_Q * alphaLM
            Next e
        Next i
    Next k
End Sub

' braking windows, number from the bridge length: whole spans up to BRAKE_L, each window overlaps the one before
' by two spans, so every pier has both its spans and a neighbouring span loaded in one window: the most critical
' placement for each pier with few cases (owner 2026-11-06). One span of overlap left the window-edge piers 2-4 %
' low in the 33-span run, because the free bearings also drag on the next pier.
Private Sub BrakeWindows()
    Dim a As Long, b As Long, s As Double
    nFW = 0
    ReDim fwA(1 To nSp): ReDim fwB(1 To nSp)
    a = 1
    Do
        s = spL(a): b = a
        Do While b < nSp
            If s + spL(b + 1) > BRAKE_L Then Exit Do
            b = b + 1: s = s + spL(b)
        Loop
        If b = a And b < nSp Then
            warnText = warnText & "- Spans " & a & " and " & a + 1 & " together exceed " & BRAKE_L & " m: braking windows hold one span, the pier between them is not fully covered." & vbCrLf
            b = a + 1
        End If
        nFW = nFW + 1: fwA(nFW) = a: fwB(nFW) = b
        If b >= nSp Then Exit Do
        a = IIf(b - 1 > a, b - 1, b)
    Loop
End Sub

Private Function NTracks(ByVal d As Long) As Long
    NTracks = IIf(tH2(d) > 0, 2, 1)
End Function

Private Sub AddLoad(ByVal el As Long, ByVal lc As String, ByVal dirn As String, ByVal eccType As Long, ByVal eccDir As String, ByVal ecc As Double, ByVal q As Double)
    Dim id As Long, s As String
    id = loadIds(el) + 1: loadIds(el) = id
    s = "{""ID"":" & id & ",""LCNAME"":" & QS(lc) & ",""GROUP_NAME"":"""",""CMD"":""BEAM"",""TYPE"":""UNILOAD"",""DIRECTION"":" & QS(dirn) & _
        ",""USE_PROJECTION"":false,""USE_ECCEN"":" & IIf(eccType >= 0, "true", "false")
    If eccType >= 0 Then s = s & ",""ECCEN_TYPE"":" & eccType & ",""ECCEN_DIR"":" & QS(eccDir) & ",""I_END"":" & JN(ecc) & ",""J_END"":" & JN(ecc) & ",""USE_J_END"":false"
    s = s & ",""D"":[0,1,0,0],""P"":[" & JN(q) & "," & JN(q) & ",0,0],""USE_ADDITIONAL"":false,""ADDITIONAL_I_END"":0,""ADDITIONAL_J_END"":0,""USE_ADDITIONAL_J_END"":false}"
    If loadItems.Exists(el) Then loadItems(el) = loadItems(el) & "," & s Else loadItems(el) = s
End Sub

Private Function CaseList() As Collection
    Dim c As New Collection, i As Long, a As Variant, v As Variant
    If rtiMode Then
        c.Add "FREN|D|Braking on track 1"
        c.Add "IVME|D|Traction"
        c.Add "TU+|T|Temperature Effect (+)"
        c.Add "TU-|T|Temperature Effect (-)"
        If rtiNL Then c.Add "LLS|D|LM71 static: nonlinear live load stages"
        Set CaseList = c
        Exit Function
    End If
    If seisMode Then
        c.Add "DC|D|Self Weight"
        c.Add "DW_1|DC|Ballast, kerb, railing, panel, cross beams"
        c.Add "DW_2|DW|Protection concrete, rails"
        c.Add MASS_CASE & "|L|Uniform train load: seismic mass and TKDY combinations"
        Set CaseList = c
        Exit Function
    End If
    a = Array("DC|D|Self Weight", "DW_1|DC|Ballast, kerb, railing, panel, cross beams", "DW_2|DW|Protection concrete, rails", "FR|FR|Friction Force", _
        "WS-X|W|Horizontal Wind Load (Long.)", "WS-Y|W|Horizontal Wind Load (Trans.)", "WS-Z|W|Vertical Wind Load", "WL|WL|Wind Load on Live Load", _
        "TU+|T|Temperature Effect (+)", "TU-|T|Temperature Effect (-)", "PL|L|Pedestrian Live Load", "MOVE|L|Nosing")
    For Each v In a: c.Add v: Next v
    For i = 1 To nFW
        c.Add "FREN " & i & "|BRK|Train Braking - spans " & fwA(i) & " to " & fwB(i)
    Next i
    If hasCF Then c.Add "CF|CF|Centrifugal Force"
    For i = 1 To nSp
        c.Add "IVME " & AxisName(i) & "|BRK|Train Acceleration - " & AxisName(i)
    Next i
    Set CaseList = c
End Function

Private Function StldJson() As String
    Dim c As Collection, i As Long, p As Variant, s As String
    Set c = CaseList()
    For i = 1 To c.Count
        p = Split(c(i), "|")
        s = s & IIf(i = 1, "", ",") & QS(CStr(i)) & ":{""NO"":" & i & ",""NAME"":" & QS(CStr(p(0))) & ",""TYPE"":" & QS(CStr(p(1))) & ",""DESC"":" & QS(CStr(p(2))) & "}"
    Next i
    StldJson = s
End Function

Private Function SortedLongKeys(d As Object) As Variant
    Dim a() As Long, i As Long, j As Long, t As Long, k As Variant
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

Private Sub AddLoadPayloads()
    Dim s As String, i As Long, e As Long, k As Variant, cnt As Long, d As Long
    If rtiMode Then
        AddRtiPayloads
        Exit Sub
    End If
    Add "db/STLD", "{""Assign"":{" & StldJson() & "}}"
    Add "db/BODF", "{""Assign"":{""1"":{""LCNAME"":""DC"",""GROUP_NAME"":"""",""FV"":[0,0,-1]}}}"
    s = "": cnt = 0
    For Each k In SortedLongKeys(loadItems)
        s = s & IIf(s = "", "", ",") & QS(CStr(k)) & ":{""ITEMS"":[" & loadItems(CLng(k)) & "]}"
        cnt = cnt + 1
        If cnt = BMLD_CHUNK Then Add "db/BMLD", "{""Assign"":{" & s & "}}": s = "": cnt = 0
    Next k
    If s <> "" Then Add "db/BMLD", "{""Assign"":{" & s & "}}"
    ' cross beams: their weight at the deck node of each bearing line
    s = ""
    For i = 1 To nBrg
        d = spType(brgSpan(i))
        s = s & IIf(s = "", "", ",") & QS(CStr(brgNode(i))) & ":{""ITEMS"":[{""ID"":1,""LCNAME"":""DW_1"",""GROUP_NAME"":"""",""FZ"":" & JN(-tQ(d, Q_ENL) * tNG(d)) & "}]}"
    Next i
    Add "db/CNLD", "{""Assign"":{" & s & "}}"
    If seisMode Then
        AddSeismicPayloads
        Exit Sub
    End If
    ' temperature: deck only (owner 2026-11-06)
    s = ""
    For i = 1 To nSp
        For e = dkA(i) To dkB(i)
            s = s & IIf(s = "", "", ",") & QS(CStr(e)) & ":{""ITEMS"":[{""ID"":1,""LCNAME"":""TU+"",""GROUP_NAME"":"""",""TEMP"":" & JN(tuPos) & "},{""ID"":2,""LCNAME"":""TU-"",""GROUP_NAME"":"""",""TEMP"":" & JN(tuNeg) & "}]}"
        Next e
    Next i
    Add "db/ETMP", "{""Assign"":{" & s & "}}"
    AddMoving
    AddCombs
End Sub

' LM71 lanes, vehicle, moving-load case and dynamic factor
Private Sub AddMoving()
    Dim s As String, i As Long, e As Long
    Add "db/MVCD", "{""Assign"":{""1"":{""CODE"":""EUROCODE""}}}"
    s = LaneJson("LANE1", 1)
    If AnyTwoTracks() Then s = s & "," & LaneJson("LANE2", 2)
    Add "db/LLAN", "{""Assign"":{" & s & "}}"
    Add "db/MVHL", "{""Assign"":{""1"":{""MVLD_CODE"":11,""VEHICLE_LOAD_NAME"":""Load Model 71"",""VEHICLE_LOAD_NUM"":1,""VEHICLE_TYPE_NAME"":""Load Model 71"",""VEH_EUROCODE"":{""SUB_TYPE"":23,""W1"":80,""DD1"":0,""D1"":0.8,""W2"":80,""DD2"":0,""D2"":0.8,""V_LOAD_FACTOR"":" & JN(alphaLM) & ",""LONGI_DIST"":false,""ECCEN_VERT_LOAD"":false}}}}"
    Add "db/MVLDeu", "{""Assign"":{""1"":{""LCNAME"":""LM71"",""OPT_AUTO_OPTIMIZE"":false,""TYPE_LOADMODEL"":5,""DESC"":"""",""OPT_COMB"":1,""SCALE_FACTOR1"":0.8,""SCALE_FACTOR2"":0.7,""SCALE_FACTOR3"":0.6," & _
        """OPT_PSI_FACTOR"":true,""MULTI_FACTOR1"":1,""MULTI_FACTOR2"":1,""MULTI_FACTOR3"":0.75,""SUB_LOAD_LIST"":[{""TYPE"":2,""NAME"":""Load Model 71"",""SCALE_FACTOR"":1," & _
        """MIN_LOAD_LANE_TYPE"":0,""MAX_LOAD_LANE_TYPE"":" & IIf(AnyTwoTracks(), 2, 1) & ",""SLN_LIST"":[""LANE1""" & IIf(AnyTwoTracks(), ",""LANE2""", "") & "]}]}}}"
    If phiByElem Then
        s = ""
        For i = 1 To nSp
            For e = dkA(i) To dkB(i)
                s = s & IIf(s = "", "", ",") & QS(CStr(e)) & ":{""INPUT_TYPE"":1,""DYN_FACTOR"":" & JN(tPhi(spType(i))) & "}"
            Next e
        Next i
        Add "db/DYNF", "{""Assign"":{" & s & "}}"
    Else
        Add "db/DYFG", "{""Assign"":{""1"":{""INPUT_TYPE"":0,""LENGTH"":" & JN(tLc(1)) & ",""MAINTAIN_TYPE"":1,""OPT_REDUCE_EFF"":false,""HEIGHT_COVER"":0,""DYN_FACTOR"":" & JN(tPhi(1)) & "}}}"
    End If
End Sub

' seismic model: stiffness factors with no boundary group, so they reach the eigenvalue analysis too
Private Sub AddSeismicPayloads()
    Dim s As String, i As Long
    ' TKDY effective stiffness: bending x EQ_STIFF (deck, cap, piers), EQ_STIFF_SLAB (connection slab)
    For i = 1 To nDk
        s = s & SecfSeis(DeckSec(i), EQ_STIFF, 3) & "," & SecfSeis(SlabSec(i), EQ_STIFF_SLAB, 1) & ","
    Next i
    s = s & SecfSeis(SEC_CAP, EQ_STIFF, 1)
    For i = 1 To nPS
        s = s & "," & SecfSeis(SEC_PIER0 + i - 1, EQ_STIFF, 1)
    Next i
    Add "db/SECF", "{""Assign"":{" & s & "}}"
    Add "db/LTOM", "{""Assign"":{""1"":{""DIR"":""XYZ"",""bNODAL"":true,""bBEAM"":true,""bFLOOR"":true,""bPRES"":true,""GRAV"":" & JN(GRAV_STYP) & _
        ",""vLC"":[{""LCNAME"":""DW_1"",""FACTOR"":1},{""LCNAME"":""DW_2"",""FACTOR"":1},{""LCNAME"":" & QS(MASS_CASE) & ",""FACTOR"":" & JN(facRows("TKDY")(1)) & "}]}}}"
    Add "db/EIGV", "{""Assign"":{""1"":{""TYPE"":""RITZ"",""bINCNL"":false,""iGNUM"":0,""vRITZ"":[{""KIND"":""GROUND"",""GROUND"":""ACCX"",""iNOG"":" & RITZ_X & "},{""KIND"":""GROUND"",""GROUND"":""ACCY"",""iNOG"":" & RITZ_Y & "},{""KIND"":""GROUND"",""GROUND"":""ACCZ"",""iNOG"":" & RITZ_Z & "}]}}}"
    s = ""
    For i = 1 To specNames.Count
        s = s & IIf(i = 1, "", ",") & QS(CStr(i)) & ":{""NAME"":" & QS(CStr(specNames(i))) & ",""iTYPE"":1,""iMETHOD"":0,""SCALE"":1,""GRAV"":" & JN(GRAV_STYP) & ",""DRATIO"":0.05,""DESC"":"""",""aFUNC"":[" & specJson(i) & "]}"
    Next i
    Add "db/SPFC", "{""Assign"":{" & s & "}}"
    Add "db/SPLC", "{""Assign"":{" & RsCase(1, "EQ-X-(+)", 0) & "," & RsCase(2, "EQ-Y-(+)", 90) & "," & RsCase(3, "EQ-X-(-)", 180) & "," & RsCase(4, "EQ-Y-(-)", 270) & "}}"
    AddSeisCombs
End Sub

Private Function AnyTwoTracks() As Boolean
    Dim d As Long
    For d = 1 To nDk
        If tH2(d) > 0 Then AnyTwoTracks = True: Exit Function
    Next d
End Function

Private Function SecfSeis(ByVal sec As Long, ByVal f As Double, ByVal part As Long) As String
    SecfSeis = QS(CStr(sec)) & ":{""ITEMS"":[{""ID"":1,""GROUP_NAME"":"""",""AREA_SF"":1,""ASY_SF"":1,""ASZ_SF"":1,""IXX_SF"":1,""IYY_SF"":" & JN(f) & ",""IZZ_SF"":" & JN(f) & _
        ",""WGT_SF"":1,""W_SF"":1,""IPART"":" & part & ",""bDiffIJ"":false}]}"
End Function

Private Function RsCase(ByVal no As Long, ByVal nm As String, ByVal ang As Long) As String
    RsCase = QS(CStr(no)) & ":{""NAME"":" & QS(nm) & ",""DIR"":""XY"",""ANGLE"":" & ang & ",""SCALE"":1,""PMFT"":1,""bDAMP"":false,""INTERP"":""LOG"",""DESC"":"""",""COMTYPE"":""CQC"",""bADDSIGN"":true,""iSIGNTYPE"":0,""bMODE"":false,""aFUNCNAME"":[" & QS(EQ_FUNC) & "]}"
End Function

' lane t on every deck element of the spans that have that track; ECC = minus the offset to the left
Private Function LaneJson(ByVal nm As String, ByVal t As Long) As String
    Dim i As Long, e As Long, s As String, d As Long, v As Double
    For i = 1 To nSp
        d = spType(i)
        If t = 1 Or tH2(d) > 0 Then
            If t = 1 Then
                v = tW(d) / 2 - tKL(d) - tH1(d) / 2
            Else
                v = tW(d) / 2 - tKL(d) - tH1(d) - tH2(d) / 2
            End If
            For e = dkA(i) To dkB(i)
                s = s & IIf(s = "", "", ",") & "{""ELEM"":" & e & ",""ECC"":" & JN(-v) & ",""FACT"":0,""SPAN_START"":false,""ECCEN_VERT_LOAD"":0,""CENT_F"":0.5}"
            Next e
        End If
    Next i
    LaneJson = QS(CStr(t)) & ":{""COMMON"":{""LL_NAME"":" & QS(nm) & ",""LOAD_DIST"":""LANE"",""GROUP_NAME"":"""",""SKEW_START"":0,""SKEW_END"":0,""MOVING"":""BOTH""," & _
        """WHEEL_SPACE"":1.5,""WIDTH"":2,""OPT_AUTO_LANE"":false,""ALLOW_WIDTH"":3},""LANE_ITEMS"":[" & s & "]}"
End Function

' ---- combinations: Tablo 14 / 15 of 1_GIRIS ----
Private Function It(ByVal an As String, ByVal lc As String, ByVal f As Double) As String
    If f < 0 Then Exit Function
    It = ";" & an & "|" & lc & "|" & JN(f)
End Function

Private Function Comb(ByVal nm As String, ByVal iType As Long, ByVal items As String) As String
    Dim p As Variant, i As Long, s As String, q2 As Variant
    If Left$(items, 1) = ";" Then items = Mid$(items, 2)
    p = Split(items, ";")
    For i = 0 To UBound(p)
        q2 = Split(p(i), "|")
        s = s & IIf(i = 0, "", ",") & "{""ANAL"":" & QS(CStr(q2(0))) & ",""LCNAME"":" & QS(CStr(q2(1))) & ",""FACTOR"":" & q2(2) & "}"
    Next i
    combNo = combNo + 1
    Comb = QS(CStr(combNo)) & ":{""NO"":" & combNo & ",""NAME"":" & QS(nm) & ",""ACTIVE"":""ACTIVE"",""bCB"":" & IIf(InStr(";" & items, ";CB|") > 0, "true", "false") & ",""iTYPE"":" & iType & ",""DESC"":"""",""vCOMB"":[" & s & "]}"
End Function

' mode 0 = service (table values), 1 = max, 2 = min
Private Function TableItems(ByVal lab As String, ByVal mode As Long, ByVal strNo As Long) As String
    Dim f As Variant, fDC As Double, fDW As Double, s As String
    If Not facRows.Exists(lab) Then Err.Raise vbObjectError + 10, , "Tablo 14 row '" & lab & "' not found."
    f = facRows(lab)
    fDC = f(0): fDW = f(0)
    If fDC = -2 Then
        If mode = 2 Then
            fDC = dcMin: fDW = dwMin
        Else
            fDC = IIf(strNo = 4, dcMax4, dcMax): fDW = dwMax
        End If
    End If
    s = It("ST", "DC", fDC) & It("ST", "DW_1", fDC) & It("ST", "DW_2", fDW) & It("ST", "FR", f(2)) & _
        It("ST", "WS-X", f(3)) & It("ST", "WS-Y", f(3)) & It("ST", "WS-Z", f(3)) & It("ST", "WL", f(4)) & It("CB", "TU", f(5)) & _
        It("ST", "PL", f(1)) & It("ST", "MOVE", f(1)) & It("MV", "LM71", f(1)) & It("CB", "IVME/FREN", f(1))
    If hasCF Then s = s & It("ST", "CF", f(1))
    TableItems = s
End Function

Private Function EqItems(ByVal dead As Double, ByVal ll As Double, ByVal mainC As String, ByVal otherC As String) As String
    EqItems = It("ST", "DC", dead) & It("ST", "DW_1", dead) & It("ST", "DW_2", dead) & It("ST", MASS_CASE, ll) & It("RS", mainC, eqMain) & It("RS", otherC, eqOther)
End Function

' TKDY: Tablo 14 dead and LL factors; H-TKDY: minimum dead load (Tablo 15 DC min), no live load, as ALTYAPI
Private Sub AddSeisCombs()
    Dim t As String, k As Long, f As Variant, pr As Variant, q As Variant, envD As String
    combNo = 0
    f = facRows("TKDY")
    pr = Array("X-1|EQ-X-(+)|EQ-Y-(+)", "Y-1|EQ-Y-(+)|EQ-X-(+)", "X-2|EQ-X-(+)|EQ-Y-(-)", "Y-2|EQ-Y-(-)|EQ-X-(+)", _
               "X-3|EQ-X-(-)|EQ-Y-(+)", "Y-3|EQ-Y-(+)|EQ-X-(-)", "X-4|EQ-X-(-)|EQ-Y-(-)", "Y-4|EQ-Y-(-)|EQ-X-(-)")
    For k = 0 To 7
        q = Split(pr(k), "|")
        t = t & IIf(k = 0, "", ",") & Comb("TKDY-EQ-" & q(0), 0, EqItems(f(0), f(1), CStr(q(1)), CStr(q(2))))
        envD = envD & It("CB", "TKDY-EQ-" & q(0), 1)
    Next k
    For k = 0 To 7
        q = Split(pr(k), "|")
        t = t & "," & Comb("H-TKDY-EQ-" & q(0), 0, EqItems(dcMin, -1, CStr(q(1)), CStr(q(2))))
        envD = envD & It("CB", "H-TKDY-EQ-" & q(0), 1)
    Next k
    Add "db/LCOM-GEN", "{""Assign"":{" & t & "}}"
    Add "db/LCOM-GEN", "{""Assign"":{" & Comb("ENV_DEP", 1, envD) & "}}"
End Sub

Private Sub AddCombs()
    Dim t As String, a As String, i As Long, k As Long, f As Variant, pr As Variant, q As Variant
    Dim mins As Variant, maxs As Variant, envS As String, envR As String
    combNo = 0
    For i = 1 To nSp
        a = a & It("ST", "IVME " & AxisName(i), 1)
    Next i
    ' combinations used by other combinations go first
    t = Comb("IVME", 1, a)
    t = t & "," & Comb("TU", 1, It("ST", "TU+", 1) & It("ST", "TU-", 1))
    a = ""
    For i = 1 To nFW
        a = a & It("ST", "FREN " & i, 1)
    Next i
    t = t & "," & Comb("FREN", 1, a)
    Add "db/LCOM-GEN", "{""Assign"":{" & t & "}}"
    Add "db/LCOM-GEN", "{""Assign"":{" & Comb("IVME/FREN", 0, It("CB", "FREN", 1) & It("CB", "IVME", 1)) & "}}"
    t = ""
    For k = 1 To 4
        t = t & IIf(t = "", "", ",") & Comb("Servis " & k, 0, TableItems("Servis-" & k, 0, k))
        envS = envS & It("CB", "Servis " & k, 1)
    Next k
    mins = Array("STR-1-Min.", "STR-2-Min.", "STR-3-Min", "STR-4-Min.", "STR-5-Min")
    maxs = Array("STR-1-Maks.", "STR-2-Maks.", "STR-3-Maks.", "STR-4-Maks.", "STR-5-Maks")
    For k = 1 To 5
        t = t & "," & Comb(CStr(mins(k - 1)), 0, TableItems("STR-" & k, 2, k))
        t = t & "," & Comb(CStr(maxs(k - 1)), 0, TableItems("STR-" & k, 1, k))
        envR = envR & It("CB", CStr(mins(k - 1)), 1) & It("CB", CStr(maxs(k - 1)), 1)
    Next k
    Add "db/LCOM-GEN", "{""Assign"":{" & t & "}}"
    t = Comb("ENV_SER", 1, envS)
    t = t & "," & Comb("ENV_STR", 1, envR)
    Add "db/LCOM-GEN", "{""Assign"":{" & t & "}}"
    Add "db/LCOM-GEN", "{""Assign"":{" & Comb("ENV_ALL", 1, It("CB", "ENV_SER", 1) & It("CB", "ENV_STR", 1)) & "}}"
End Sub

' ============================================================================
'  RTI BRANCH: rail layer and stages (owner decisions 2026-11-05 / 06)
'  One rail line per track (2 rails), rail nodes above track nodes tied rigidly to the deck
'  nodes, embankment RTI_EMB each side on fixed ground nodes. Each rail node has a rigid-free
'  link (longitudinal free) and two multilinear track springs in boundary groups: unloaded
'  U<t>_<k> and loaded L<t>_<k>, k = segment between window ends.
'  Stages: braking window (BRAKE_L, track 1) with head and tail on every axis, both directions;
'  traction (TRAC_L, rear of the window, last track); loads on the bridge part only, loaded
'  spring under each train; then TU+, TU-, and END = loaded springs on the bridge for LM71.
'  Live load gate (owner 2026-11-06): moving load LM71 on END by default (track springs act
'  linearly there). Nonlinear adds static LM71 stages before END: whole-span trains, both
'  tracks, x alpha x Phi, loaded springs under the train, from zero, so they sum like FREN + IVME.
' ============================================================================
Private Function TrackOff(ByVal d As Long, ByVal t As Long) As Double
    If t = 1 Then
        TrackOff = tW(d) / 2 - tKL(d) - tH1(d) / 2
    Else
        If tH2(d) <= 0 Then Err.Raise vbObjectError + 60, , "RTI: every span needs both tracks when one span has two."
        TrackOff = tW(d) / 2 - tKL(d) - tH1(d) - tH2(d) / 2
    End If
End Function

Private Function NewX(ByVal x As Double, ByVal y As Double, ByVal z As Double) As Long
    xNo = xNo + 1
    xX(xNo) = x: xY(xNo) = y: xZ(xNo) = z
    NewX = NUM_LIMIT + xNo - 1
End Function

Private Sub RailPt(ByVal t As Long, ByVal x As Double, ByVal y As Double, ByVal deckNode As Long, ByVal sp As Long)
    Dim n As Long, g As Long, r As Long
    rN(t) = rN(t) + 1: n = rN(t)
    g = NewX(x, y, 0)
    r = NewX(x, y, railH)
    rGr(t, n) = g: rNode(t, n) = r: rAng(t, n) = -phiS(sp): rOn(t, n) = (deckNode > 0)
    If deckNode > 0 Then
        Rigid deckNode, CStr(g)
    Else
        consJson = consJson & IIf(consJson = "", "", ",") & QS(CStr(g)) & ":{""ITEMS"":[{""ID"":1,""GROUP_NAME"":""BG_RAY"",""CONSTRAINT"":""1111110""}]}"
    End If
    If n = 1 Then
        rSt(t, n) = 0
    Else
        rSt(t, n) = rSt(t, n - 1) + Sqr((x - xX(rGr(t, n - 1) - NUM_LIMIT + 1)) ^ 2 + (y - xY(rGr(t, n - 1) - NUM_LIMIT + 1)) ^ 2)
        reCnt = reCnt + 1
        rE1(reCnt) = rNode(t, n - 1): rE2(reCnt) = r: rET(reCnt) = t: rEI(reCnt) = n - 1
    End If
End Sub

Private Sub BuildRail()
    Dim t As Long, i As Long, e As Long, d As Long, k As Long, m As Long, mx As Long, v As Double, n1 As Long, bx As Double, by As Double
    nTrk = IIf(AnyTwoTracks(), 2, 1)
    m = CLng(RTI_EMB / meshDx)
    mx = eCnt + 2 * m + 20
    ReDim rNode(1 To 2, 1 To mx): ReDim rGr(1 To 2, 1 To mx): ReDim rSt(1 To 2, 1 To mx): ReDim rTr(1 To 2, 1 To mx)
    ReDim rAng(1 To 2, 1 To mx): ReDim rOn(1 To 2, 1 To mx)
    ReDim xX(1 To 4 * mx): ReDim xY(1 To 4 * mx): ReDim xZ(1 To 4 * mx)
    ReDim rE1(1 To 2 * mx): ReDim rE2(1 To 2 * mx): ReDim rET(1 To 2 * mx): ReDim rEI(1 To 2 * mx)
    ReDim rFirst(1 To 2, 1 To nSp): ReDim rLast(1 To 2, 1 To nSp)
    xNo = 0: reCnt = 0: rN(1) = 0: rN(2) = 0
    For t = 1 To nTrk
        v = TrackOff(spType(1), t)
        n1 = eN1(dkA(1))
        bx = nX(n1) - v * cyS(1): by = nY(n1) + v * cxS(1)
        For k = m To 1 Step -1
            RailPt t, bx - k * meshDx * cxS(1), by - k * meshDx * cyS(1), 0, 1
        Next k
        For i = 1 To nSp
            d = spType(i): v = TrackOff(d, t)
            rFirst(t, i) = rN(t) + 1
            For e = dkA(i) To dkB(i)
                RailPt t, nX(eN1(e)) - v * cyS(i), nY(eN1(e)) + v * cxS(i), eN1(e), i
            Next e
            RailPt t, nX(eN2(dkB(i))) - v * cyS(i), nY(eN2(dkB(i))) + v * cxS(i), eN2(dkB(i)), i
            rLast(t, i) = rN(t)
        Next i
        v = TrackOff(spType(nSp), t)
        n1 = eN2(dkB(nSp))
        bx = nX(n1) - v * cyS(nSp): by = nY(n1) + v * cxS(nSp)
        For k = 1 To m
            RailPt t, bx + k * meshDx * cxS(nSp), by + k * meshDx * cyS(nSp), 0, nSp
        Next k
        For i = 1 To rN(t)
            rTr(t, i) = ((IIf(i < rN(t), rSt(t, IIf(i < rN(t), i + 1, i)) - rSt(t, i), 0)) + IIf(i > 1, rSt(t, i) - rSt(t, IIf(i > 1, i - 1, i)), 0)) / 2
        Next i
    Next t
    BuildWindows
    BuildRailLoads
    nLL = 0
    If rtiNL Then BuildLiveStages
End Sub

' trains on whole spans: one span, two spans, to each end of the bridge. Axles at the middle of the
' first span of the train, and of the last. On the old model this list reaches 98 % of the moving
' load peaks in tension and compression; away from the peaks it stays lower (trains end at axes).
Private Sub BuildLiveStages()
    Dim i As Long, e As Long, seen As Object, q As Double, nt As Long, phi As Double, s As Double, c As Double, l As Double, tot As Double, lSum As Double, grp As String, hit As Long
    Set seen = CreateObject("Scripting.Dictionary")
    ReDim lpA(1 To 8 * nSp): ReDim lpB(1 To 8 * nSp): ReDim lpX(1 To 8 * nSp)
    For i = 1 To nSp
        LiveTrain i, i, seen
        LiveTrain i, IIf(i < nSp, i + 1, i), seen
        LiveTrain i, nSp, seen
        LiveTrain 1, i, seen, i
    Next i
    For i = 1 To nSp
        nt = IIf(tH2(spType(i)) > 0, 2, 1)
        phi = tPhi(spType(i))
        q = LM71_Q * nt * alphaLM * phi
        tot = 0
        For e = dkA(i) To dkB(i)
            DeckLoad e, "LS_" & i, -q
            tot = tot + ElemLen(e)
        Next e
        ' axles over the uniform load: the same total on the elements around the middle of the span
        c = tot / 2: s = 0: lSum = 0: hit = 0
        For e = dkA(i) To dkB(i)
            l = ElemLen(e)
            If Abs(s + l / 2 - c) <= LM71_AXL / 2 Then lSum = lSum + l
            If s <= c And c <= s + l Then hit = e
            s = s + l
        Next e
        q = (LM71_AX - LM71_Q * LM71_AXL) * nt * alphaLM * phi
        grp = "LX_" & i
        If lSum = 0 Then
            DeckLoad hit, grp, -q / ElemLen(hit)
        Else
            s = 0
            For e = dkA(i) To dkB(i)
                l = ElemLen(e)
                If Abs(s + l / 2 - c) <= LM71_AXL / 2 Then DeckLoad e, grp, -q / lSum
                s = s + l
            Next e
        End If
    Next i
End Sub

' train on spans a to b with its axles on span a and on span b (x = one span only)
Private Sub LiveTrain(ByVal a As Long, ByVal b As Long, seen As Object, Optional ByVal x As Long = 0)
    If x = 0 Then x = a
    LiveAdd a, b, x, seen
    If b > a Then LiveAdd a, b, b, seen
End Sub

Private Sub LiveAdd(ByVal a As Long, ByVal b As Long, ByVal x As Long, seen As Object)
    Dim k As String
    k = a & "_" & b & "_" & x
    If seen.Exists(k) Then Exit Sub
    seen(k) = 1
    nLL = nLL + 1
    lpA(nLL) = a: lpB(nLL) = b: lpX(nLL) = x
End Sub

Private Function ElemLen(ByVal e As Long) As Double
    ElemLen = Sqr((nX(eN2(e)) - nX(eN1(e))) ^ 2 + (nY(eN2(e)) - nY(eN1(e))) ^ 2)
End Function

' uniform vertical load on a deck element, case LLS
Private Sub DeckLoad(ByVal el As Long, ByVal grp As String, ByVal q As Double)
    Dim id As Long, s As String
    lgUsed(grp) = 1
    id = 1
    If railItems.Exists(el) Then id = 2 + (Len(railItems(el)) - Len(Replace(railItems(el), """ID"":", ""))) \ 5 - 1
    s = "{""ID"":" & id & ",""LCNAME"":""LLS"",""GROUP_NAME"":" & QS(grp) & ",""CMD"":""BEAM"",""TYPE"":""UNILOAD"",""DIRECTION"":""GZ"",""USE_PROJECTION"":false,""USE_ECCEN"":false," & _
        """D"":[0,1,0,0],""P"":[" & JN(q) & "," & JN(q) & ",0,0],""USE_ADDITIONAL"":false,""ADDITIONAL_I_END"":0,""ADDITIONAL_J_END"":0,""USE_ADDITIONAL_J_END"":false}"
    If railItems.Exists(el) Then railItems(el) = railItems(el) & "," & s Else railItems(el) = s
End Sub

Private Function AxisSt(ByVal t As Long, ByVal a As Long) As Double
    If a = 0 Then
        AxisSt = rSt(t, rFirst(t, 1))
    Else
        AxisSt = rSt(t, rLast(t, a))
    End If
End Function

Private Function AxisLabel(ByVal a As Long) As String
    If a = 0 Then
        AxisLabel = "A1"
    ElseIf a = nSp Then
        AxisLabel = "A2"
    Else
        AxisLabel = "P" & a
    End If
End Function

' position a: even = axis a / 2, odd = middle of the span between two axes
Private Function PosSt(ByVal t As Long, ByVal a As Long) As Double
    If a Mod 2 = 0 Then
        PosSt = AxisSt(t, a \ 2)
    Else
        PosSt = (AxisSt(t, a \ 2) + AxisSt(t, a \ 2 + 1)) / 2
    End If
End Function

Private Function PosLabel(ByVal a As Long) As String
    If a Mod 2 = 0 Then
        PosLabel = AxisLabel(a \ 2)
    Else
        PosLabel = "S" & (a \ 2 + 1)
    End If
End Function

' a, b braking train on track 1; c, d traction zone on the last track; e, f the whole train on that track
' br / tr: the window carries braking / traction (one track: one of them, two tracks: both)
Private Sub AddWin(ByVal nm As String, ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double, ByVal e As Double, ByVal f As Double, ByVal sgn1 As Long, Optional ByVal br As Boolean = True, Optional ByVal tr As Boolean = True)
    ' a window that does not reach the bridge carries no load
    If b <= rSt(1, rFirst(1, 1)) + 1 Or a >= rSt(1, rLast(1, nSp)) - 1 Then Exit Sub
    If Not br And (d <= rSt(nTrk, rFirst(nTrk, 1)) + 0.001 Or c >= rSt(nTrk, rLast(nTrk, nSp)) - 0.001) Then Exit Sub
    nWin = nWin + 1
    wNm(nWin) = nm: wA(nWin) = a: wB(nWin) = b: wC(nWin) = c: wD(nWin) = d: wE(nWin) = e: wF(nWin) = f: wBr(nWin) = br: wTr(nWin) = tr: wS(nWin) = sgn1
End Sub

' two tracks: one window, braking train on track 1 and traction on track 2. One track (owner 2026-11-09):
' the same train position twice, once braking only and once traction only (name + "_T"), never both
Private Sub WinPair(ByVal nm As String, ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double, ByVal e As Double, ByVal f As Double, ByVal sgn1 As Long)
    If nTrk > 1 Then
        AddWin nm, a, b, c, d, e, f, sgn1
    Else
        AddWin nm, a, b, c, d, e, f, sgn1, True, False
        AddWin nm & "_T", a, b, c, d, e, f, sgn1, False, True
    End If
End Sub

Private Sub BuildWindows()
    Dim a As Long, s1 As Double, s2 As Double, tl As Double, tt As Long
    tt = nTrk
    nWin = 0
    ReDim wNm(1 To 16 * (nSp + 1)): ReDim wA(1 To 16 * (nSp + 1)): ReDim wB(1 To 16 * (nSp + 1)): ReDim wC(1 To 16 * (nSp + 1)): ReDim wD(1 To 16 * (nSp + 1)): ReDim wE(1 To 16 * (nSp + 1)): ReDim wF(1 To 16 * (nSp + 1)): ReDim wBr(1 To 16 * (nSp + 1)): ReDim wTr(1 To 16 * (nSp + 1)): ReDim wS(1 To 16 * (nSp + 1))
    tl = TRAC_L
    ' window ends on every axis and every mid-span: the mid-span positions close the gap to the
    ' every-element run of the old model (rail compression and deck movement, owner 2026-11-06)
    For a = 0 To 2 * nSp
        s1 = PosSt(1, a): s2 = PosSt(tt, a)
        ' train moving +: head at the position, then tail at the position; traction at the rear
        WinPair "F_BAS_" & PosLabel(a), s1 - BRAKE_L, s1, s2 - BRAKE_L, s2 - BRAKE_L + tl, s2 - BRAKE_L, s2, 1
        WinPair "F_SON_" & PosLabel(a), s1, s1 + BRAKE_L, s2, s2 + tl, s2, s2 + BRAKE_L, 1
    Next a
    For a = 2 * nSp To 0 Step -1
        s1 = PosSt(1, a): s2 = PosSt(tt, a)
        ' train moving -
        WinPair "B_BAS_" & PosLabel(a), s1, s1 + BRAKE_L, s2 + BRAKE_L - tl, s2 + BRAKE_L, s2, s2 + BRAKE_L, -1
        WinPair "B_SON_" & PosLabel(a), s1 - BRAKE_L, s1, s2 - tl, s2, s2 - BRAKE_L, s2, -1
    Next a
End Sub

Private Function SortedD(d As Object) As Double()
    Dim a() As Double, i As Long, j As Long, v As Double, k As Variant
    ReDim a(1 To d.Count)
    For Each k In d.Keys
        i = i + 1: a(i) = CDbl(d(k))
    Next k
    For i = 2 To UBound(a)
        v = a(i): j = i - 1
        Do While j >= 1
            If a(j) <= v Then Exit Do
            a(j + 1) = a(j): j = j - 1
        Loop
        a(j + 1) = v
    Next i
    SortedD = a
End Function

Private Sub AddCut(d As Object, ByVal s As Double)
    d(Format(Round(s * 1000), "0")) = s
End Sub

' segment number of a station on track t
Private Function SegOf(ByVal t As Long, ByVal s As Double) As Long
    Dim k As Long, c() As Double
    If t = 1 Then c = cut1 Else c = cut2
    For k = 1 To UBound(c)
        If s < c(k) Then SegOf = k: Exit Function
    Next k
    SegOf = UBound(c) + 1
End Function

' middle station of segment k of track t
Private Function SegMid(ByVal t As Long, ByVal k As Long) As Double
    Dim c() As Double
    If t = 1 Then c = cut1 Else c = cut2
    If k = 1 Then
        SegMid = c(1) - 1000
    ElseIf k > UBound(c) Then
        SegMid = c(UBound(c)) + 1000
    Else
        SegMid = (c(k - 1) + c(k)) / 2
    End If
End Function

Private Sub BuildRailLoads()
    Dim d1 As Object, d2 As Object, i As Long, t As Long, tt As Long, e As Long, k As Long, n As Long, q As Double, nm As String, s As String
    Dim s0 As Double, s1 As Double, le As Double, pa As Double, pb As Double, c() As Double, j As Long, ra As Double, rb As Double
    Const EPS As Double = 0.001
    Set d1 = CreateObject("Scripting.Dictionary"): Set d2 = CreateObject("Scripting.Dictionary")
    tt = nTrk
    For i = 1 To nWin
        AddCut d1, wA(i) - EPS: AddCut d1, wB(i) + EPS
        If tt = 1 Then
            AddCut d1, wC(i) - EPS: AddCut d1, wD(i) + EPS
        Else
            AddCut d2, wC(i) - EPS: AddCut d2, wD(i) + EPS
            AddCut d2, wE(i) - EPS: AddCut d2, wF(i) + EPS
        End If
    Next i
    AddCut d1, rSt(1, rFirst(1, 1)) - EPS: AddCut d1, rSt(1, rLast(1, nSp)) + EPS
    If nTrk = 2 Then AddCut d2, rSt(2, rFirst(2, 1)) - EPS: AddCut d2, rSt(2, rLast(2, nSp)) + EPS
    cut1 = SortedD(d1)
    If nTrk = 2 Then cut2 = SortedD(d2)
    Set bgUsed = CreateObject("Scripting.Dictionary")
    Set lgUsed = CreateObject("Scripting.Dictionary")
    Set railItems = CreateObject("Scripting.Dictionary")
    Set fnIds = CreateObject("Scripting.Dictionary")
    For t = 1 To nTrk
        For i = 1 To rN(t)
            k = SegOf(t, rSt(t, i))
            bgUsed("U" & t & "_" & k) = 1: bgUsed("L" & t & "_" & k) = 1
            FnId rTr(t, i)
        Next i
    Next t
    ' each bridge rail element is cut at the window ends inside it and every piece is loaded in the group
    ' of its own segment (partial UNILOAD, D = distance ratios), so a window carries exactly its length
    For e = 1 To reCnt
        t = rET(e): i = rEI(e)
        If rOn(t, i) And rOn(t, i + 1) Then
            s0 = rSt(t, i): s1 = rSt(t, i + 1): le = s1 - s0
            If t = 1 Then c = cut1 Else c = cut2
            pa = s0
            Do While pa < s1 - 0.000001
                pb = s1
                For j = 1 To UBound(c)
                    If c(j) > pa + 0.000001 And c(j) < pb Then pb = c(j): Exit For
                Next j
                k = SegOf(t, (pa + pb) / 2)
                ra = (pa - s0) / le: rb = (pb - s0) / le
                If t = 1 Then
                    RailLoad e, "FREN", "FB_" & k, BRAKE_Q * alphaLM, ra, rb
                    RailLoad e, "FREN", "RB_" & k, -BRAKE_Q * alphaLM, ra, rb
                End If
                If t = tt Then
                    RailLoad e, "IVME", "FT_" & k, TRAC_Q * alphaLM, ra, rb
                    RailLoad e, "IVME", "RT_" & k, -TRAC_Q * alphaLM, ra, rb
                End If
                pa = pb
            Loop
        End If
    Next e
End Sub

' force-deformation function of a tributary length: id of the unloaded law; loaded = id + 1
Private Function FnId(ByVal trib As Double) As Long
    Dim k As String
    k = Format(Round(trib * 1000), "0")
    If Not fnIds.Exists(k) Then fnIds(k) = 2 * fnIds.Count + 1
    FnId = fnIds(k)
End Function

Private Sub RailLoad(ByVal e As Long, ByVal lc As String, ByVal grp As String, ByVal q As Double, ByVal ra As Double, ByVal rb As Double)
    Dim el As Long, id As Long, s As String
    el = NUM_LIMIT + e - 1
    lgUsed(grp) = 1
    id = 1
    If railItems.Exists(el) Then id = 2 + (Len(railItems(el)) - Len(Replace(railItems(el), """ID"":", ""))) \ 5 - 1
    s = "{""ID"":" & id & ",""LCNAME"":" & QS(lc) & ",""GROUP_NAME"":" & QS(grp) & ",""CMD"":""BEAM"",""TYPE"":""UNILOAD"",""DIRECTION"":""LX"",""USE_PROJECTION"":false,""USE_ECCEN"":false," & _
        """D"":[" & JN(ra) & "," & JN(rb) & ",0,0],""P"":[" & JN(q) & "," & JN(q) & ",0,0],""USE_ADDITIONAL"":false,""ADDITIONAL_I_END"":0,""ADDITIONAL_J_END"":0,""USE_ADDITIONAL_J_END"":false}"
    If railItems.Exists(el) Then railItems(el) = railItems(el) & "," & s Else railItems(el) = s
End Sub

Private Sub AddRailGeometry()
    Dim s As String, i As Long, c As Long
    For i = 1 To xNo
        s = s & IIf(s = "", "", ",") & QS(CStr(NUM_LIMIT + i - 1)) & ":{""X"":" & JN(xX(i)) & ",""Y"":" & JN(xY(i)) & ",""Z"":" & JN(xZ(i)) & "}"
        c = c + 1
        If c = 1000 Then Add "db/NODE", "{""Assign"":{" & s & "}}": s = "": c = 0
    Next i
    If s <> "" Then Add "db/NODE", "{""Assign"":{" & s & "}}"
    s = "": c = 0
    For i = 1 To reCnt
        s = s & IIf(s = "", "", ",") & QS(CStr(NUM_LIMIT + i - 1)) & ":{""TYPE"":""BEAM"",""MATL"":" & MAT_RAIL & ",""SECT"":" & SEC_RAIL & ",""NODE"":[" & rE1(i) & "," & rE2(i) & ",0,0,0,0,0,0],""ANGLE"":0,""STYPE"":0}"
        c = c + 1
        If c = 1000 Then Add "db/ELEM", "{""Assign"":{" & s & "}}": s = "": c = 0
    Next i
    If s <> "" Then Add "db/ELEM", "{""Assign"":{" & s & "}}"
End Sub

Private Function BngrRti() As String
    Dim s As String, n As Long, k As Variant
    s = """1"":{""NAME"":""BG_MESNET"",""AUTOTYPE"":0},""2"":{""NAME"":""BG_RAY"",""AUTOTYPE"":0}"
    n = 2
    For Each k In bgUsed.Keys
        n = n + 1
        s = s & "," & QS(CStr(n)) & ":{""NAME"":" & QS(CStr(k)) & ",""AUTOTYPE"":0}"
    Next k
    BngrRti = s
End Function

Private Sub AddRailLinks()
    Dim s As String, t As Long, i As Long, k As Long, c As Long, no As Long, f As Long, kk As Variant, trib As Double
    ' force-deformation functions: unloaded and loaded law per tributary length
    For Each kk In fnIds.Keys
        trib = CDbl(kk) / 1000#
        f = fnIds(kk)
        s = s & IIf(s = "", "", ",") & MlfcJson(f, "U_" & kk, trib * TRK_Q_UNL) & "," & MlfcJson(f + 1, "L_" & kk, trib * TRK_Q_LOAD)
    Next kk
    Add "db/MLFC", "{""Assign"":{" & s & "}}"
    ' links run down from the rail node to the track node, like the bearing links: ANGLE -phi then
    ' turns the spring along the track (an upward link put it 2 phi off, measured 2026-11-09)
    s = "": no = nLink
    For t = 1 To nTrk
        For i = 1 To rN(t)
            k = SegOf(t, rSt(t, i))
            f = FnId(rTr(t, i))
            no = no + 1
            s = s & IIf(s = "", "", ",") & QS(CStr(no)) & ":{""NODE"":[" & rNode(t, i) & "," & rGr(t, i) & "],""ANGLE"":" & JN(rAng(t, i)) & ",""LINK"":""GEN"",""R_S"":[true,true,false,true,false,true],""SDR"":[0,0,0,0,0,0],""bSHEAR"":false,""DR"":[0.5,0.5],""BNGR_NAME"":""BG_RAY""}"
            no = no + 1
            s = s & "," & QS(CStr(no)) & ":{""NODE"":[" & rNode(t, i) & "," & rGr(t, i) & "],""ANGLE"":" & JN(rAng(t, i)) & ",""LINK"":""MULTILINEAR"",""DIR"":2,""MLFC"":" & f & ",""DRENDI"":0.5,""BNGR_NAME"":" & QS("U" & t & "_" & k) & "}"
            no = no + 1
            s = s & "," & QS(CStr(no)) & ":{""NODE"":[" & rNode(t, i) & "," & rGr(t, i) & "],""ANGLE"":" & JN(rAng(t, i)) & ",""LINK"":""MULTILINEAR"",""DIR"":2,""MLFC"":" & (f + 1) & ",""DRENDI"":0.5,""BNGR_NAME"":" & QS("L" & t & "_" & k) & "}"
            c = c + 3
            If c >= 900 Then Add "db/ELNK", "{""Assign"":{" & s & "}}": s = "": c = 0
        Next i
    Next t
    If s <> "" Then Add "db/ELNK", "{""Assign"":{" & s & "}}"
End Sub

Private Function MlfcJson(ByVal id As Long, ByVal nm As String, ByVal f As Double) As String
    MlfcJson = QS(CStr(id)) & ":{""NAME"":" & QS(nm) & ",""TYPE"":""FORCE"",""SYMM"":true,""FUNC_ID"":0,""ITEMS"":[{""X"":0,""Y"":0},{""X"":" & JN(TRK_U0) & ",""Y"":" & JN(f) & "},{""X"":1,""Y"":" & JN(f) & "}]}"
End Function

Private Sub AddRtiPayloads()
    Dim s As String, i As Long, e As Long, k As Variant, cnt As Long, n As Long
    Add "db/STLD", "{""Assign"":{" & StldJson() & "}}"
    s = """1"":{""NAME"":""LG_TUP""},""2"":{""NAME"":""LG_TUN""}"
    n = 2
    For Each k In lgUsed.Keys
        n = n + 1
        s = s & "," & QS(CStr(n)) & ":{""NAME"":" & QS(CStr(k)) & "}"
    Next k
    Add "db/LDGR", "{""Assign"":{" & s & "}}"
    s = "": cnt = 0
    For Each k In SortedLongKeys(railItems)
        s = s & IIf(s = "", "", ",") & QS(CStr(k)) & ":{""ITEMS"":[" & railItems(CLng(k)) & "]}"
        cnt = cnt + 1
        If cnt = BMLD_CHUNK Then Add "db/BMLD", "{""Assign"":{" & s & "}}": s = "": cnt = 0
    Next k
    If s <> "" Then Add "db/BMLD", "{""Assign"":{" & s & "}}"
    ' deck temperature only, no rail temperature (UIC 774-3 1.4.2, continuous welded rail)
    s = ""
    For i = 1 To nSp
        For e = dkA(i) To dkB(i)
            s = s & IIf(s = "", "", ",") & QS(CStr(e)) & ":{""ITEMS"":[{""ID"":1,""LCNAME"":""TU+"",""GROUP_NAME"":""LG_TUP"",""TEMP"":" & JN(tuPos) & "},{""ID"":2,""LCNAME"":""TU-"",""GROUP_NAME"":""LG_TUN"",""TEMP"":" & JN(tuNeg) & "}]}"
        Next e
    Next i
    Add "db/ETMP", "{""Assign"":{" & s & "}}"
    AddMoving
    AddStages
End Sub

' stage list: active boundary and load groups of every stage, sent as changes from the stage before
Private Sub AddStages()
    Dim prevB As Object, prevL As Object, curB As Object, curL As Object, s As String, i As Long, n As Long, t As Long, tt As Long, k As Long, m As Double
    Dim nSeg As Long, win1 As Boolean, win2 As Boolean, nm As String, kind As Long, u As Variant, tot As Long, act As String, j As Long, llNm As String
    Const EPS As Double = 0.002
    Set prevB = CreateObject("Scripting.Dictionary"): Set prevL = CreateObject("Scripting.Dictionary")
    tt = nTrk
    tot = nWin + nLL + 3
    For i = 1 To tot
        Set curB = CreateObject("Scripting.Dictionary"): Set curL = CreateObject("Scripting.Dictionary")
        kind = 0
        If i = nWin + 1 Then kind = 1      ' TU+
        If i = nWin + 2 Then kind = 2      ' TU-
        If i > nWin + 2 And i < tot Then kind = 4   ' nonlinear live load
        If i = tot Then kind = 3           ' END
        j = i - nWin - 2
        For t = 1 To nTrk
            ' not IIf: it evaluates both sides, and cut2 is empty on a single-track bridge
            If t = 1 Then nSeg = UBound(cut1) + 1 Else nSeg = UBound(cut2) + 1
            For k = 1 To nSeg
                m = SegMid(t, k)
                If bgUsed.Exists("U" & t & "_" & k) Then
                    win1 = False
                    If kind = 0 Then
                        If t = 1 Then win1 = (m >= wA(i) - EPS And m <= wB(i) + EPS)
                        If t = tt And tt > 1 Then win1 = (m >= wE(i) - EPS And m <= wF(i) + EPS)
                    ElseIf kind = 3 Then
                        win1 = (m >= rSt(t, rFirst(t, 1)) - EPS And m <= rSt(t, rLast(t, nSp)) + EPS)
                    ElseIf kind = 4 Then
                        win1 = (m >= AxisSt(t, lpA(j) - 1) - EPS And m <= AxisSt(t, lpB(j)) + EPS)
                    End If
                    curB(IIf(win1, "L", "U") & t & "_" & k) = 1
                End If
                ' loads apart from the springs: a short segment can hold an element midpoint but no rail node
                If kind = 0 Then
                    If t = 1 And wBr(i) And m >= wA(i) - EPS And m <= wB(i) + EPS Then
                        nm = IIf(wS(i) > 0, "FB_", "RB_") & k
                        If lgUsed.Exists(nm) Then curL(nm) = 1
                    End If
                    If t = tt And wTr(i) And m >= wC(i) - EPS And m <= wD(i) + EPS Then
                        nm = IIf(wS(i) > 0, "FT_", "RT_") & k
                        If lgUsed.Exists(nm) Then curL(nm) = 1
                    End If
                End If
            Next k
        Next t
        If kind = 1 Then curL("LG_TUP") = 1
        If kind = 2 Then curL("LG_TUN") = 1
        llNm = ""
        If kind = 4 Then
            For k = lpA(j) To lpB(j)
                curL("LS_" & k) = 1
            Next k
            curL("LX_" & lpX(j)) = 1
            llNm = "LL_" & lpA(j) & "_" & lpB(j) & "_" & lpX(j)
        End If
        act = ""
        If i = 1 Then
            act = """ACT_ELEM"":[{""GRUP_NAME"":""SG_TUM"",""AGE"":0}],"
            curB("BG_MESNET") = 1: curB("BG_RAY") = 1
        Else
            curB("BG_MESNET") = 1: curB("BG_RAY") = 1
        End If
        act = act & ActJson(curB, prevB, "ACT_BNGR", "BNGR_NAME", "POS", "ORIGINAL", "DACT_BNGR", False)
        act = act & ActJson(curL, prevL, "ACT_LOAD", "LOAD_NAME", "DAY", "FIRST", "DACT_LOAD", True)
        If Right$(act, 1) = "," Then act = Left$(act, Len(act) - 1)
        nm = Choose(kind + 1, IIf(i <= nWin, wNm(IIf(i <= nWin, i, 1)), ""), "TU+", "TU-", "END", llNm)
        s = s & IIf(s = "", "", ",") & QS(CStr(i)) & ":" & StageJ(nm, act)
        n = n + 1
        If n = 20 Then Add "db/STAG", "{""Assign"":{" & s & "}}": s = "": n = 0
        Set prevB = curB: Set prevL = curL
    Next i
    If s <> "" Then Add "db/STAG", "{""Assign"":{" & s & "}}"
    Add "db/STCT", "{""Assign"":{""1"":{""bLAST_FINAL"":true,""FINAL_STAGE"":""END"",""iINC_NLA"":0,""iNLA_TYPE"":1,""bINC_PDL"":false,""iITER"":30,""TOL"":0.01,""bINC_TDE"":false," & _
        """vEREC"":[{""LTYPECC"":""FREN+IVME"",""EREC"":""L"",""vLCNAME"":[""IVME"",""FREN""]},{""LTYPECC"":""TU+"",""EREC"":""T"",""vLCNAME"":[""TU+""]},{""LTYPECC"":""TU-"",""EREC"":""T"",""vLCNAME"":[""TU-""]}" & IIf(rtiNL, ",{""LTYPECC"":""LLS"",""EREC"":""L"",""vLCNAME"":[""LLS""]}", "") & "]," & _
        """CPFC"":""INTERNAL"",""bCONV"":false,""bTRUSS"":false,""bBEAM"":false,""bCHANGE_CABLE"":false,""bAPPLY_IMF"":false,""bCAMBER"":false,""bSD"":false,""iBSC"":0}}}"
End Sub

Private Function StageJ(ByVal nm As String, ByVal acts As String) As String
    StageJ = "{""NAME"":" & QS(nm) & ",""DURATION"":0,""bSV_RSLT"":true,""bSV_STEP"":false,""bLOAD_STEP"":false,""ADD_STEP"":[]" & IIf(acts = "", "", "," & acts) & "}"
End Function

' activation and deactivation lists between two stages
Private Function ActJson(cur As Object, prev As Object, ByVal keyA As String, ByVal keyN As String, ByVal keyP As String, ByVal valP As String, ByVal keyD As String, ByVal dObj As Boolean) As String
    Dim a As String, d As String, k As Variant
    For Each k In cur.Keys
        If Not prev.Exists(k) Then a = a & IIf(a = "", "", ",") & "{" & QS(keyN) & ":" & QS(CStr(k)) & "," & QS(keyP) & ":" & QS(valP) & "}"
    Next k
    For Each k In prev.Keys
        If Not cur.Exists(k) Then
            If dObj Then
                d = d & IIf(d = "", "", ",") & "{" & QS(keyN) & ":" & QS(CStr(k)) & "," & QS(keyP) & ":" & QS(valP) & "}"
            Else
                d = d & IIf(d = "", "", ",") & QS(CStr(k))
            End If
        End If
    Next k
    If a <> "" Then ActJson = QS(keyA) & ":[" & a & "],"
    If d <> "" Then ActJson = ActJson & QS(keyD) & ":[" & d & "],"
End Function

' ============================================================================
'  ENGINE (copied from MIDAS_USTYAPI)
' ============================================================================
Private Function OpenWb(ByVal p As String, ByRef opened As Boolean) As Workbook
    Dim w As Workbook, nm As String
    nm = Mid$(p, InStrRev(p, "\") + 1)
    For Each w In Application.Workbooks
        If StrComp(w.Name, nm, vbTextCompare) = 0 Then Set OpenWb = w: Exit Function
    Next w
    Application.ScreenUpdating = False
    Set OpenWb = Application.Workbooks.Open(Filename:=p, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Application.ScreenUpdating = True
    opened = True
End Function

Private Sub ReadGirder(wb As Workbook, ByVal typ As String, Optional ByVal isY As Boolean = False)
    Dim t(1 To 12) As Double
    Dim r As Range, hr As Long, rr As Long, cc As Long, i As Long, nm As Variant, col As Long, found As Boolean
    Set r = NmRange(wb, "G_KIRIS_TABLOSU")
    ' the header row (Ad, b2, ...) may sit just above the named range
    hr = r.Row - 2: If hr < 1 Then hr = 1
    Set r = r.Worksheet.Range(r.Worksheet.Cells(hr, r.Column), r.Cells(r.Rows.Count, r.Columns.Count))
    For hr = 1 To r.Rows.Count
        For cc = 1 To r.Columns.Count
            If Trim$(CStr(r.Cells(hr, cc).Value)) = "b2" Then GoTo HaveHeader
        Next cc
    Next hr
    Err.Raise vbObjectError + 4, , "KIRIS_INPUT: header row with b2 not found."
HaveHeader:
    For rr = hr + 1 To r.Rows.Count
        If StrComp(Trim$(CStr(r.Cells(rr, 1).Value)), Trim$(typ), vbTextCompare) = 0 Then found = True: Exit For
    Next rr
    If Not found Then Err.Raise vbObjectError + 5, , "KIRIS_INPUT: girder type '" & typ & "' not found."
    nm = Array("b2", "b3", "b4", "b5", "b6", "b7", "h2", "h3", "h4", "h5", "h6", "h7")
    For i = 0 To 11
        col = 0
        For cc = 1 To r.Columns.Count
            If Trim$(CStr(r.Cells(hr, cc).Value)) = nm(i) Then col = cc: Exit For
        Next cc
        If col = 0 Then Err.Raise vbObjectError + 6, , "KIRIS_INPUT: column " & nm(i) & " not found."
        t(i + 1) = NumOr0(r.Cells(rr, col).Value) / 100#
    Next i
    If isY Then
        For i = 1 To 12: kbY(i) = t(i): Next i
        If Abs(t(7) + t(8) + t(9) + t(10) + t(11) + t(12) - Hg) > 0.001 Then Err.Raise vbObjectError + 12, , "KIRIS_INPUT: '" & typ & "' height differs from the span girder."
        Exit Sub
    End If
    For i = 1 To 12: kb(i) = t(i): Next i
    Dim hs As Double: hs = kb(7) + kb(8) + kb(9) + kb(10) + kb(11) + kb(12)
    If Abs(hs - Hg) > 0.001 Then warnText = warnText & "- KIRIS_INPUT h2..h7 = " & hs & " m, G_KIRIS_YUKSEKLIGI = " & Hg & " m (KIRIS_INPUT used)." & vbCrLf
    Hg = hs
    If Abs(kb(1) / 2 - (kb(2) / 2 + kb(4) + kb(5))) > 0.0005 Then warnText = warnText & "- KIRIS_INPUT: b2/2 <> b3/2 + b5 + b6." & vbCrLf
    If Abs(kb(3) / 2 - (kb(2) / 2 + kb(6))) > 0.0005 Then warnText = warnText & "- KIRIS_INPUT: b4/2 <> b3/2 + b7." & vbCrLf
End Sub

' db/CO_S (Section Color): W = wireframe (2D), HF = hidden fill (3D surface), HE = hidden edge (3D outline)
Private Function SectColorJson() As String
    Dim i As Long, s As String
    For i = 1 To nDk
        s = s & ColorItem(DeckSec(i), 45, 110, 200, i) & "," & ColorItem(SlabSec(i), 60, 140, 75, i) & ","
    Next i
    s = s & ColorItem(SEC_CAP, 205, 125, 40, 1)
    For i = 1 To nPS
        s = s & "," & ColorItem(SEC_PIER0 + i - 1, 150, 150, 150, i)
    Next i
    SectColorJson = "{""Assign"":{" & s & "}}"
End Function

Private Function ColorItem(ByVal no As Long, ByVal r As Long, ByVal g As Long, ByVal bl As Long, ByVal typeNo As Long) As String
    Dim f As Double, v(1 To 3) As Long, t(1 To 3) As Long, k As Long
    f = 1 - SECT_SHADE_STEP * (typeNo - 1)
    If f < 0.4 Then f = 0.4
    v(1) = CLng(r * f): v(2) = CLng(g * f): v(3) = CLng(bl * f)
    For k = 1 To 3
        t(k) = CLng(v(k) + (255 - v(k)) * SECT_FILL_TINT)
    Next k
    ColorItem = QS(CStr(no)) & ":{""W_R"":" & v(1) & ",""W_G"":" & v(2) & ",""W_B"":" & v(3) & _
        ",""HF_R"":" & t(1) & ",""HF_G"":" & t(2) & ",""HF_B"":" & t(3) & _
        ",""HE_R"":" & v(1) & ",""HE_G"":" & v(2) & ",""HE_B"":" & v(3) & ",""bBLEMD"":false,""FACT"":0.5}"
End Function

Private Function AddNode(ByVal x As Double, ByVal y As Double, ByVal z As Double) As Long
    Dim k As String
    k = JN(Round(x, 6)) & "|" & JN(Round(y, 6)) & "|" & JN(Round(z, 6))
    If nodeKey.Exists(k) Then AddNode = nodeKey(k): Exit Function
    nCnt = nCnt + 1
    If nCnt > UBound(nX) Then ReDim Preserve nX(1 To nCnt * 2): ReDim Preserve nY(1 To nCnt * 2): ReDim Preserve nZ(1 To nCnt * 2)
    nX(nCnt) = Round(x, 6): nY(nCnt) = Round(y, 6): nZ(nCnt) = Round(z, 6)
    nodeKey(k) = nCnt
    AddNode = nCnt
End Function

Private Function AddElem(ByVal mat As Long, ByVal sec As Long, ByVal a As Long, ByVal b As Long) As Long
    eCnt = eCnt + 1
    If eCnt > UBound(eMat) Then ReDim Preserve eMat(1 To eCnt * 2): ReDim Preserve eSec(1 To eCnt * 2): ReDim Preserve eN1(1 To eCnt * 2): ReDim Preserve eN2(1 To eCnt * 2)
    eMat(eCnt) = mat: eSec(eCnt) = sec: eN1(eCnt) = a: eN2(eCnt) = b
    AddElem = eCnt
End Function

Private Sub AddX(d As Object, ByVal v As Double)
    Dim k As Variant
    For Each k In d.Keys
        If Abs(CDbl(d(k)) - v) < 0.0005 Then Exit Sub
    Next k
    d(CStr(d.Count)) = v
End Sub

Private Function SortedKeys(d As Object) As Variant
    Dim a() As Double, i As Long, j As Long, t As Double, k As Variant
    ReDim a(1 To d.Count)
    For Each k In d.Keys: i = i + 1: a(i) = d(k): Next k
    For i = 1 To UBound(a) - 1
        For j = i + 1 To UBound(a)
            If a(j) < a(i) Then t = a(i): a(i) = a(j): a(j) = t
        Next j
    Next i
    SortedKeys = a
End Function

Private Function EnDb(ByVal cls As String) As String
    Dim a As Variant, i As Long, c As String
    c = UCase$(Replace(Trim$(cls), " ", ""))
    a = Array("C12/15", "C16/20", "C20/25", "C25/30", "C30/37", "C35/45", "C40/50", "C45/55", "C50/60", "C55/67", "C60/75", "C70/85", "C80/95", "C90/105")
    For i = 0 To UBound(a)
        If Split(a(i), "/")(0) = c Or a(i) = c Then EnDb = a(i): Exit Function
    Next i
    Err.Raise vbObjectError + 11, , "Concrete class '" & cls & "' has no EN04(RC) database name."
End Function

Private Sub Add(ByVal ep As String, ByVal body As String)
    payloads.Add Array(ep, body)
End Sub

Private Function CleanTables() As Variant
    CleanTables = Array("LCOM-CONC", "LCOM-GEN", "BCCT", "DYNF", "SPLC", "SPFC", "ETMP", "CNLD", "STCT", "TDPL", "TDCS", "TDNA", "TDGR", "TDNT", "CSCS", "STAG", _
        "DYFG", "MVLDeu", "MVHL", "LLAN", "MVCD", "THGC", "EIGV", "ACTL", "LTOM", "BMLD", "BODF", "STMP", "LDGR", "STLD", "ESSF", "SECF", "MLFC", _
        "NSPR", "ELNK", "RIGD", "CONS", "BNGR", "GRUP", "ELEM", "NODE", "SECT", "TMAT", "TDME", "TDMT", "MATL")
End Function

' The project comment of a model built by this macro says which type it is (ends SEISMIC or not).
Private Function ModelTypeOk(ByRef msg As String) As Boolean
    Dim pj As String, was As String, txt As String
    ModelTypeOk = True
    If Not Api("GET", "db/PJCF", "", pj) Then Exit Function
    If InStr(pj, SCRIPT_ID) = 0 Then Exit Function
    was = IIf(InStr(pj, " SEISMIC") > 0, "SEISMIC", IIf(InStr(pj, " RTI") > 0, "RTI", "STATIC"))
    If was = ModelName() Then Exit Function
    txt = "The model open in Civil NX is a " & was & " model, but you are building the " & ModelName() & " model."
    If askUser Then
        If UiMsg(txt & vbCrLf & vbCrLf & "WIPE it anyway?", vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) = vbYes Then Exit Function
    End If
    msg = "STOPPED: " & txt & " Nothing was changed. Open the right .mcb, or use an empty file."
    ModelTypeOk = False
End Function

Private Function ModelName() As String
    ModelName = IIf(seisMode, "SEISMIC", IIf(rtiMode, "RTI", "STATIC"))
End Function

' stiffness factors of the seismic model: one record per section, no boundary group
Private Function SecfCheck() As String
    Dim resp As String, n As Long, nWant As Long, sReport As String
    nWant = 2 * nDk + 1 + nPS
    If Not Api("GET", "db/SECF", "", resp) Then SecfCheck = "  stiffness factors: FAIL (db/SECF not readable)" & vbCrLf: Exit Function
    n = CountOf(resp, """IYY_SF"":")
    sReport = "  stiffness factors " & n & " / " & nWant
    If n <> nWant Then sReport = sReport & "  FAIL"
    SecfCheck = sReport & vbCrLf
End Function

Private Function CleanModel() As String
    Dim t As Variant, resp As String, remain As String, nTry As Long
    ' pass 1: keyless DELETE of every table; pass 2: what is still there is deleted BY KEY
    ' (Civil NX answers 200 to a keyless DELETE of db/ACTL but keeps the record; with
    ' {"Argument":["1"]} it is removed - tested live 2026-11-03).
    Dim leftTbl As Collection, keys As String
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
    ' ChrW keeps this file ASCII: 9608 full block, 9617 light shade
    Application.StatusBar = "MIDAS_KOPRU  [" & String(filled, ChrW(9608)) & String(PROG_BAR_WIDTH - filled, ChrW(9617)) & _
        "] " & pct & "%   (" & progStep & "/" & progTotal & ")  " & stepName
    On Error GoTo 0
    DoEvents
End Sub

Private Function TopKeys(ByVal resp As String) As String
    Dim i As Long, n As Long, ch As String, depth As Long, j As Long, out As String, expectKey As Boolean
    n = Len(resp)
    i = 1
    Do While i <= n
        ch = Mid$(resp, i, 1)
        If ch = """" Then
            j = i + 1
            Do While j <= n
                If Mid$(resp, j, 1) = "\" Then
                    j = j + 2
                ElseIf Mid$(resp, j, 1) = """" Then
                    Exit Do
                Else
                    j = j + 1
                End If
            Loop
            If depth = 2 And expectKey Then out = out & IIf(out = "", "", ",") & Mid$(resp, i, j - i + 1)
            expectKey = False
            i = j + 1
        Else
            Select Case ch
                Case "{": depth = depth + 1: expectKey = (depth = 2)
                Case "}": depth = depth - 1
                Case ",": expectKey = (depth = 2)
                Case "[": depth = depth + 1
                Case "]": depth = depth - 1
            End Select
            i = i + 1
        End If
    Loop
    TopKeys = out
End Function

Private Function TableHasRecords(ByVal resp As String, ByVal tbl As String) As Boolean
    TableHasRecords = (InStr(1, resp, """" & tbl & """:{""", vbBinaryCompare) > 0)
End Function

Private Function Api(ByVal method As String, ByVal ep As String, ByVal body As String, ByRef resp As String) As Boolean
    On Error GoTo Bad
    If http Is Nothing Then Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open method, API_BASE & ep, False
    http.SetTimeouts 30000, 60000, 600000, 600000
    http.SetRequestHeader "MAPI-Key", mKey
    http.SetRequestHeader "Content-Type", "application/json"
    If method = "GET" Then http.Send Else http.Send body
    resp = http.ResponseText
    If http.Status = 404 Then mKey = ""
    Api = (http.Status >= 200 And http.Status < 300 And InStr(1, resp, """error""", vbTextCompare) = 0)
    Exit Function
Bad:
    resp = "HTTP error: " & Err.Description
    Set http = Nothing
    Api = False
End Function

Private Function GetKey() As Boolean
    Dim cellKey As String
    On Error Resume Next
    cellKey = Trim$(CStr(ThisWorkbook.Worksheets(KEY_SHEET).Range(KEY_CELL).Value))
    On Error GoTo 0
    If cellKey <> "" Then mKey = cellKey
    If mKey = "" Then mKey = Environ$("MIDAS_MAPI_KEY")
    If mKey = "" Then mKey = UiInputS("MIDAS MAPI key (kept in memory for this Excel session only):", MSG_TITLE)
    GetKey = (mKey <> "")
End Function

Private Function SectArea(ByVal resp As String, ByVal key As String) As Double
    Dim p As Long, q As Long
    p = InStr(1, resp, """" & key & """:{")
    If p = 0 Then Exit Function
    p = InStr(p, resp, """AREA"":")
    If p = 0 Then Exit Function
    q = p + 7
    SectArea = Val(Mid$(resp, q, 30))
End Function

Private Function CountOf(ByVal s As String, ByVal what As String) As Long
    CountOf = (Len(s) - Len(Replace(s, what, ""))) / Len(what)
End Function

Private Function FitBox(ByVal txt As String, ByVal tag As String, ByVal reserve As Long) As String
    FitBox = txt
    Exit Function
    Dim lim As Long, p As String
    lim = 1000 - reserve
    If Len(txt) <= lim Then FitBox = txt: Exit Function
    On Error Resume Next
    p = ThisWorkbook.Path & "\MIDAS"
    MakeDir p
    If Err.Number <> 0 Or Len(Dir$(p, vbDirectory)) = 0 Then Err.Clear: p = Environ$("TEMP")
    p = p & "\" & tag & ".txt"
    WriteText p, txt
    On Error GoTo 0
    FitBox = Left$(txt, lim - Len(p) - 30) & " ..." & vbCrLf & "(full text: " & p & ")"
End Function

Private Function NmRange(wb As Workbook, ByVal nm As String) As Range
    On Error Resume Next
    Set NmRange = wb.Names(nm).RefersToRange
    On Error GoTo 0
    If NmRange Is Nothing Then Err.Raise vbObjectError + 20, , "Name " & nm & " missing in " & wb.Name & "."
End Function

Private Function Num(ByVal v As Variant, ByVal what As String) As Double
    If IsError(v) Then Err.Raise vbObjectError + 21, , what & ": error value."
    If IsEmpty(v) Or Not IsNumeric(v) Then Err.Raise vbObjectError + 21, , what & ": not a number ('" & CStr(v) & "')."
    Num = CDbl(v)
End Function

Private Function NumOr0(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsNumeric(v) And Not IsEmpty(v) Then NumOr0 = CDbl(v)
End Function

Private Function FindLabel(ws As Worksheet, ByVal col As Long, ByVal startText As String, Optional ByVal fromRow As Long = 1) As Range
    Dim r As Long, last As Long, t As String
    last = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    For r = fromRow To last
        t = Trim$(CStr(IIf(IsError(ws.Cells(r, col).Value), "", ws.Cells(r, col).Value)))
        If StrComp(Left$(t, Len(startText)), startText, vbTextCompare) = 0 Then Set FindLabel = ws.Cells(r, col): Exit Function
    Next r
    Err.Raise vbObjectError + 22, , ws.Name & ": label '" & startText & "' not found in column " & col & "."
End Function

Private Function LabelValue(ws As Worksheet, ByVal col As Long, ByVal startText As String, ByVal valCol As Long) As Variant
    LabelValue = ws.Cells(FindLabel(ws, col, startText).Row, valCol).Value
End Function

Private Function JN(ByVal x As Double) As String
    Dim s As String
    If Abs(x) < 0.000000000001 Then JN = "0": Exit Function
    s = Trim$(Str$(x))
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    JN = s
End Function

Private Function QS(ByVal s As String) As String
    Dim i As Long, c As Long, o As String
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        Select Case c
            Case 34: o = o & "\"""
            Case 92: o = o & "\\"
            Case Is < 32: o = o & "\u" & Right$("000" & Hex$(c), 4)
            Case Is > 126: o = o & "\u" & Right$("000" & Hex$(c), 4)
            Case Else: o = o & ChrW(c)
        End Select
    Next i
    QS = """" & o & """"
End Function

Private Function Pj(ByVal s As String) As String
    Dim i As Long, c As Long, o As String, tr As Variant, k As Long
    tr = Array(199, "C", 231, "c", 208, "G", 286, "G", 287, "g", 304, "I", 305, "i", 214, "O", 246, "o", 350, "S", 351, "s", 220, "U", 252, "u", 194, "A", 226, "a", 206, "I", 238, "i", 219, "U", 251, "u")
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        If c >= 32 And c <= 126 Then
            o = o & ChrW(c)
        Else
            For k = 0 To UBound(tr) Step 2
                If tr(k) = c Then o = o & tr(k + 1): Exit For
            Next k
            If k > UBound(tr) Then o = o & "?"
        End If
    Next i
    Pj = Left$(o, 80)
End Function

Private Sub MakeDir(ByVal p As String)
    If Len(Dir$(p, vbDirectory)) = 0 Then MkDir p
End Sub

Private Sub WriteText(ByVal p As String, ByVal txt As String)
    Dim f As Integer: f = FreeFile
    Open p For Output As #f
    Print #f, txt;
    Close #f
End Sub

