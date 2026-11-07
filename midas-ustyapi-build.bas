Attribute VB_Name = "MIDAS_USTYAPI"
Option Explicit
' ============================================================================
'  MIDAS_USTYAPI - builds the superstructure grillage model (STAGED design model or
'  DYNAMIC model, reference models/USTYAPI_DYN.mct) in MIDAS Civil NX through the MAPI
'  REST API. Runs from the 1-3A workbook (INT-LDS-DYN).
' ----------------------------------------------------------------------------
'  Button macro: MIDAS_USTYAPI_OLUSTUR
'    0. asks STAGED or DYNAMIC (dynamic = no stages, no tendons, static-based combinations)
'    1. asks for the span (Hesap Acikligi 1..6 of G_HESAP_ACIKLIK)
'    2. reads the span sheets of this workbook (4<letter>_KDME, 4<letter>_KILIF): sarkiklik (G_SARKIKLIK)
'       and the elastomer stiffnesses (1_RAPOR, column C labels
'       "Elastomer Yanal Kesme Rijitligi" / "Elastomer Eksenel Rijitligi",
'       value in column S)
'    3. reads 1-2: G_ names of INPUT, KIRIS_INPUT, 2_YUKLER (wind Wg / Wv,
'       alpha), 1_GIRIS Tablo 14 + 15 (combination factors)
'    4. shows a summary; Yes = CLEAN PASS on the model open in Civil NX (every
'       table this macro writes is deleted, then read back empty; anything left
'       -> stop, nothing written), then writes every table into that model,
'       reads back counts + the girder area and saves it (doc/SAVE)
'  Clean only: MIDAS_USTYAPI_TEMIZLE (asks first).
'  Dry run (no dialogs, no HTTP): MIDAS_USTYAPI_PAYLOAD span, outDir [, dynamicModel]
'    writes NN_<TABLE>.json files + summary.txt.
'
'  Construction stages (owner 2026-11-02): CS1 girders + bearings + self-weight +
'  pretensioned strands (4<letter>_KILIF, fpi from 4<letter>_KDME), CS2 deck, CS3 DW_1 +
'  DW_2, CS4 long term; AASHTO creep/shrinkage, ACI 209 strength gain, Magura relaxation
'  (KL 45); girder age at transfer = day f'c(t) reaches the 4-7 f'ci; combinations use the stage
'  results. End zones: girder "<type> - Y" over the length in the 1-2 INPUT cell right
'  of Sarkiklik (cm from the girder end).
'  Not built (no API endpoint): PSC design parameters, DynaGen, load
'  contribution for nonlinear case (LOADCONST), perform-analysis entries.
'  MAPI key: cell INPUT!T3 of this workbook (owner 2026-11-03), else environment
'  variable MIDAS_MAPI_KEY, else asked once per Excel session (InputBox, memory only).
'
'  Asked in pop-ups (button macro): strip spacing, link length, fixed-spring
'  stiffness. From geometry: end strip width = sarkiklik/2, extra girder
'  node at sarkiklik + H/2 from each girder end.
'  Constants below that are NOT inputs of the workbooks (taken from the
'  reference model; change here if needed): WL_ECC, NOSING_Q, NOSING_LEN,
'  NOSING_ECC, RAIL_HALF_GAUGE, HSLM data.
' ============================================================================

Private Const SCRIPT_ID As String = "midas-ustyapi-build"
Private Const SCRIPT_VERSION As String = "2026-11-07b"
Private Const SCRIPT_CHANGELOG As String = "Staged model uses AASHTO Ec (INPUT column C), dynamic keeps Eurocode Ecm; CF case loaded with the largest Qtk of 2_YUKLER Tablo 17; Inputs and messages through the macro window; Span sheets named 4A_, 4B_ ...; Reads the span sheets 4<letter>_KDME and 4<letter>_KILIF in this workbook, no 4-7 file dialog; Mass at the centroid of section (STYP bMASSOFFSET false); self-weight mass in X, Y and Z (was Z only)"

Private Const API_BASE As String = "https://moa-engineers.midasit.com:443/civil/"
Private Const MSG_TITLE As String = "MIDAS_USTYAPI"
Private Const KEY_SHEET As String = "INPUT"   ' MAPI key cell (owner 2026-11-03)
Private Const KEY_CELL As String = "T3"

' asked in pop-ups by the button macro (owner 2026-11-03); these are the proposed defaults
Private Const DEF_STRIP_DX As Double = 1.5      ' nominal spacing of transverse slab strips (m)
Private Const DEF_LINK_LEN As Double = 0.35     ' elastic link length below the bearing node (m)
Private Const DEF_K_FIX As Double = 100000#     ' fixed bearing spring (kN/m)
Private Const WL_ECC As Double = 3.5            ' lever of the train wind above the girder node (m)
Private Const NOSING_Q As Double = 100#         ' EN 1991-2 6.5.2 Qsk (kN)
Private Const NOSING_LEN As Double = 3#         ' nosing spread length (m)
Private Const NOSING_ECC As Double = 1.1        ' nosing lever above the girder node (m)
Private Const RAIL_OVER_BALLAST As Double = 0.16 ' rail level above ballast + protection concrete (as MIDAS_KOPRU)
Private Const CF_HEIGHT As Double = 1.8         ' centrifugal force above rail level (2_YUKLER C160)
Private Const RAIL_HALF_GAUGE As Double = 0.7175
Private Const GRAV As Double = 9.806          ' load-to-mass conversion (LTOM, reference model)
Private Const GRAV_STYP As Double = 9.807     ' structure type gravity (STYP, reference model)
Private Const BMLD_CHUNK As Long = 60
' section colours (owner 2026-11-03): no|R,G,B base colour = 2D wireframe + 3D hidden edge; 3D hidden fill = the
' base lightened by SECT_FILL_TINT toward white (same method as the culvert macro). Change the list to recolour.
Private Const SECT_COLORS As String = "2|60,140,75;3|150,150,150;4|90,90,90;5|120,140,165;6|45,110,200;7|45,110,200;8|25,70,150;" & _
    "9|40,150,170;10|25,105,125;11|40,150,170;12|25,105,125;13|140,165,140;14|150,150,150;15|225,130,35"
Private Const SECT_FILL_TINT As Double = 0.45
Private Const N_MODES As Long = 20         ' eigen modes (owner 2026-11-03; 20th mode ~35 Hz > 30 Hz of EN 1991-2 6.4.6)

' construction stages (owner 2026-11-02: proposal accepted; not in the workbooks)
' girder age at transfer (CS1 start) = ageG, computed in ReadStrands: the day the ACI 209 curve
' f(t) = t / (A + B t) f'c reaches the 4-7 ratio f'ci / f'c (owner 2026-11-03: 'the day where it hits 0.75 fc')
Private Const ACI_A_G As Double = 1          ' ACI 209R-92: steam-cured, type I cement (precast girder)
Private Const ACI_B_G As Double = 0.95
Private Const ACI_A_S As Double = 4          ' ACI 209R-92: moist-cured, type I cement (cast-in-place slab)
Private Const ACI_B_S As Double = 0.85
Private Const DUR_CS1 As Double = 120        ' girder storage until the deck is cast (AASHTO C5.9.3.4.2c t = 120 d; owner 2026-11-03)
Private Const RELAX_KL As Double = 45        ' Magura relaxation KL, low-relaxation strand (AASHTO C5.9.3.4.2c-1)
Private Const AGE_SLAB As Double = 3         ' slab concrete age when it becomes composite (CS2)
Private Const DUR_CS2 As Double = 30
Private Const DUR_CS3 As Double = 30         ' DW_1 + DW_2
Private Const DUR_CS4 As Double = 10000      ' long term

' ---- inputs ----
Private spanNo As Long, Lc As Double, nG As Long, sG As Double, Hg As Double
Private kSol As Double, h1 As Double, h2 As Double, kSag As Double
Private tc As Double, tMin As Double, enlB As Double, enlH As Double
Private disB As Double, disL As Double   ' outer end beam (dis enleme) beyond the girder ends (1-2 G_DIS_ENLEME_*)
Private eG As Double, eS As Double, clsG As String, clsS As String
Private denG As Double, denS As Double   ' unit weights for the staged model's user materials (kN/m3)
Private qBal As Double, qKor As Double, qBor As Double, qKork As Double, qCep As Double, qRay As Double
Private sark As Double, kyB As Double, kEB As Double
Private wG As Double, wV As Double, alphaLM As Double, tuPos As Double, tuNeg As Double, qtk As Double, railH As Double
Private kb(1 To 12) As Double   ' b2..b7, h2..h7 (m)
Private kbY(1 To 12) As Double  ' end-zone girder "<type> - Y" (section over the support)
Private zoneL As Double         ' end-zone length from each girder end (1-2 INPUT, cell right of Sarkiklik, cm)
' strands (4<letter>_KILIF + 4<letter>_KDME)
Private nRow As Long, rowZ() As Double, rowN() As Long
Private nDb As Long, dbX() As Double, dbN() As Long          ' debond stations (m from girder end), debonded strands per row
Private ageG As Double, fciRatio As Double   ' girder age at transfer (d), 4-7 f'ci / f'c
Private nLLst As Long, llX() As Double       ' 4-7 Tablo 27 live-load stations (m from the bearing)
Private fpi As Double, aStr As Double, dStr As Double, epS As Double, fpuS As Double, humid As Double
Private fckG As Double, fckS As Double
Private projName As String, clientName As String, bridgeName As String
Private warnText As String
' dynamic model (owner 2026-11-05): no stages, no tendons; without CSCS MIDAS drops the Hambly option and uses the
' merged-section J, so the build measures it on cantilevers and sets a SECF factor per girder section
Private dynMode As Boolean, askUser As Boolean
Private jMer(7 To 12) As Double, jHam(7 To 12) As Double, jFac(7 To 12) As Double
Private Const DYN_TAG As String = "@@GIRDER_SECF@@"
' asked values + geometry derived from them (owner 2026-11-03)
Private stripDx As Double, linkLen As Double, kFix As Double
Private dxAct As Double, endStripW As Double, endExtra As Double
' Hambly grillage (owner 2026-11-03): diaphragm slab flange width, overhangs, edge-girder slab widths
Private diaW As Double, ovL As Double, ovR As Double
Private visW As Double                      ' drawn width of the first/last strip (visual only, section 14)
Private polyPts As String, polyPtsY As String    ' polygon as MCT text (x, y, ...)

' ---- combination factors (-1 = omitted, -2 = gp) ----
Private facRows As Object       ' label -> Array(fDC, fLL, fWA, fWS, fWL, fTU)
Private dcMax As Double, dcMin As Double, dcMax4 As Double, dwMax As Double, dwMin As Double

' ---- model ----
Private nodeKey As Object
Private nX() As Double, nY() As Double, nZ() As Double, nCnt As Long
Private eMat() As Long, eSec() As Long, eN1() As Long, eN2() As Long, eCnt As Long
Private gEl() As Long, nSt As Long, st() As Double
Private loadItems As Object, loadIds As Object
Private crossE As String, crossN As Object
Private polyArea As Double, polyJson As String, polyPer As Double
Private polyAreaY As Double, polyJsonY As String, polyPerY As Double
Private girdE As String, girdN As Object, deckE As String, deckN As Object
Private tdnaJson As String, tdcsJson As String, tdplJson As String, nTdn As Long, nTdnGroup As Long
Private Lg As Double, yL As Double, yR As Double, nTr As Long, tr(1 To 2) As Double, trG(1 To 2) As Long
Private rigdJson As String, elnkJson As String, consJson As String, nsprJson As String

Private payloads As Collection
Private mKey As String
Private http As Object
Private progStep As Long, progTotal As Long   ' status-bar progress bar (like the culvert build)
Private Const PROG_BAR_WIDTH As Long = 24

' ============================================================================
Public Sub MIDAS_USTYAPI_OLUSTUR()
    RunBuild 0, "", False
End Sub

Public Sub MIDAS_USTYAPI_PAYLOAD(ByVal spanIndex As Long, ByVal outDir As String, Optional ByVal dynamicModel As Boolean = False)
    RunBuild spanIndex, outDir, True, dynamicModel
End Sub

' No dialogs: writes into the open model at once and saves it, report text to reportFile.
' For a scripted run; the key is passed in memory only (never stored).
Public Sub MIDAS_USTYAPI_BUILD(ByVal spanIndex As Long, ByVal reportFile As String, Optional ByVal apiKey As String = "", Optional ByVal saveModel As Boolean = True, _
                              Optional ByVal stripSpacing As Double = DEF_STRIP_DX, Optional ByVal linkLength As Double = DEF_LINK_LEN, Optional ByVal fixedSpring As Double = DEF_K_FIX, Optional ByVal dynamicModel As Boolean = False)
    Dim rep As String
    On Error GoTo Fail
    If apiKey <> "" Then mKey = apiKey
    warnText = ""
    stripDx = stripSpacing: linkLen = linkLength: kFix = fixedSpring
    dynMode = dynamicModel: askUser = False
    spanNo = spanIndex
    ReadInputs ThisWorkbook
    BuildModel
    BuildPayloads
    rep = SendAll(saveModel)
    Application.StatusBar = False
    WriteText reportFile, rep
    Exit Sub
Fail:
    Dim d As String: d = Err.Description
    On Error Resume Next
    Application.StatusBar = False
    WriteText reportFile, "FAIL: " & d
End Sub

Private Sub RunBuild(ByVal spanIndex As Long, ByVal outDir As String, ByVal dry As Boolean, Optional ByVal dynamicModel As Boolean = False)
    Dim wb12 As Workbook, msg As String, rep As String
    On Error GoTo Fail
    warnText = ""
    stripDx = DEF_STRIP_DX: linkLen = DEF_LINK_LEN: kFix = DEF_K_FIX
    Set wb12 = ThisWorkbook
    dynMode = dynamicModel: askUser = Not dry
    If Not dry Then dynMode = (UiIdx("model", 1) = 2)

    If spanIndex = 0 Then spanIndex = UiIdx("span", 0)
    If spanIndex = 0 Then spanIndex = AskSpan(wb12)
    If spanIndex = 0 Then Exit Sub
    spanNo = spanIndex

    If Not dry Then
        If Not AskModelValues() Then Exit Sub
    End If

    Application.StatusBar = "MIDAS_USTYAPI: reading inputs..."
    ReadInputs wb12

    BuildModel
    BuildPayloads

    If dry Then
        WritePayloads outDir
        Application.StatusBar = False
        Exit Sub
    End If

    msg = FitBox(Summary(), "summary", 160) & vbCrLf & vbCrLf & _
          "YES = WIPE the open Civil NX model, rebuild it as the " & IIf(dynMode, "DYNAMIC", "STAGED") & " model and SAVE it." & IIf(dynMode, " The J measurement overwrites the open file during the build.", "") & vbCrLf & _
          "NO = only write the JSON payloads (MIDAS\payload)."
    Select Case UiMsg(msg, vbYesNoCancel + vbQuestion, MSG_TITLE)
        Case vbNo
            WritePayloads ThisWorkbook.Path & "\MIDAS\payload"
            UiMsg "Payloads written to " & ThisWorkbook.Path & "\MIDAS\payload", vbInformation, MSG_TITLE
            Application.StatusBar = False
            Exit Sub
        Case vbCancel
            Application.StatusBar = False
            Exit Sub
    End Select

    rep = SendAll()
    Application.StatusBar = False
    UiMsg FitBox(rep, "report", 0), IIf(InStr(rep, "FAIL") > 0, vbExclamation, vbInformation), MSG_TITLE
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

' ============================================================================
'  INPUTS
' ============================================================================
' pop-ups (owner 2026-11-03): strip spacing, elastic link length, fixed-spring stiffness.
' Application.InputBox Type 1 = number in the Excel locale (decimal comma is fine); Cancel stops.
Private Function AskModelValues() As Boolean
    stripDx = UiNum("strip", DEF_STRIP_DX)
    linkLen = UiNum("link", DEF_LINK_LEN)
    kFix = UiNum("kfix", DEF_K_FIX)
    AskModelValues = True
End Function

Private Function AskSpan(wb As Workbook) As Long
    Dim r As Range, i As Long, t As String, v As Variant
    Set r = NmRange(wb, "G_HESAP_ACIKLIK")
    For i = 1 To r.Cells.Count
        If IsNumeric(r.Cells(i).Value) And Not IsEmpty(r.Cells(i).Value) Then
            t = t & "  " & i & " :  " & r.Cells(i).Value & " m" & vbCrLf
        End If
    Next i
    v = UiInput("Which span (Hesap Acikligi)?" & vbCrLf & t, MSG_TITLE, 1, True)
    If VarType(v) = vbBoolean Then Exit Function
    i = CLng(v)
    If i < 1 Or i > r.Cells.Count Then Err.Raise vbObjectError + 1, , "Span number " & i & " does not exist."
    AskSpan = i
End Function

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

Private Sub ReadInputs(wb12 As Workbook)
    Dim k As Long, gcol As Long, r As Range, s As Double, c As Range
    k = spanNo
    projName = CStr(NmRange(wb12, "G_PROJE_ADI").Cells(1).Value)
    clientName = CStr(NmRange(wb12, "G_IDARE_ADI").Cells(1).Value)
    bridgeName = CStr(NmRange(wb12, "G_KOPRU_ADI").Cells(1).Value)

    Lc = Num(NmRange(wb12, "G_HESAP_ACIKLIK").Cells(k).Value, "G_HESAP_ACIKLIK span " & k)
    nG = CLng(Num(NmRange(wb12, "G_KIRIS_SAYISI").Cells(k).Value, "G_KIRIS_SAYISI span " & k))
    sG = Num(NmRange(wb12, "G_KIRIS_ARALIK").Cells(k).Value, "G_KIRIS_ARALIK span " & k)
    Hg = Num(NmRange(wb12, "G_KIRIS_YUKSEKLIGI").Cells(k).Value, "G_KIRIS_YUKSEKLIGI span " & k)
    If nG < 2 Then Err.Raise vbObjectError + 2, , "G_KIRIS_SAYISI must be >= 2."
    gcol = CLng(Num(NmRange(wb12, "G_KIRIS_GABARI").Cells(k).Value, "G_KIRIS_GABARI span " & k))
    kSol = Num(NmRange(wb12, "G_GABARI_KALDIRIM_SOL").Cells(gcol).Value, "G_GABARI_KALDIRIM_SOL")
    h1 = Num(NmRange(wb12, "G_GABARI_HAT_1").Cells(gcol).Value, "G_GABARI_HAT_1")
    h2 = NumOr0(NmRange(wb12, "G_GABARI_HAT_2").Cells(gcol).Value)
    kSag = Num(NmRange(wb12, "G_GABARI_KALDIRIM_SAG").Cells(gcol).Value, "G_GABARI_KALDIRIM_SAG")

    ReadGirder wb12, CStr(NmRange(wb12, "G_KIRIS_TIPI").Cells(k).Value)
    ReadGirder wb12, Trim$(CStr(NmRange(wb12, "G_KIRIS_TIPI").Cells(k).Value)) & " - Y", True
    zoneL = Num(NmRange(wb12, "G_SARKIKLIK").Cells(1).Offset(0, 1).Value, "End-zone length (1-2 INPUT, cell right of Sarkiklik, cm from the girder end)") / 100#
    fckG = ClassFck(CStr(NmRange(wb12, "G_KIRIS_BETONU").Cells(1, 1).Value))
    fckS = ClassFck(CStr(NmRange(wb12, "G_TABLIYE_BETONU").Cells(1, 1).Value))

    tMin = Num(NmRange(wb12, "G_MIN_DOSEME_KALINLIGI").Cells(1).Value, "G_MIN_DOSEME_KALINLIGI")
    tc = Num(NmRange(wb12, "G_TABLIYE_KALINLIGI").Cells(1).Offset(0, 1).Value, "Tabliye Kalinligi (INPUT column C)")
    enlB = Num(NmRange(wb12, "G_ENLEME_KALINLIGI").Cells(1).Value, "G_ENLEME_KALINLIGI")
    enlH = ReadEnlemeH(wb12)
    disB = Num(NmRange(wb12, "G_DIS_ENLEME_KALINLIGI").Cells(1).Value, "G_DIS_ENLEME_KALINLIGI")
    disL = Num(NmRange(wb12, "G_DIS_ENLEME_BOYU").Cells(1).Value, "G_DIS_ENLEME_BOYU")
    If disB <= 0 Or disB > 2 Or disL <= 0 Then Err.Raise vbObjectError + 13, , "Dis enleme " & disB & " x " & disL & " m is out of range."
    ' cross-beam heights in 1-2 are measured from the deck top with the minimum slab G_MIN_DOSEME_KALINLIGI (0.25);
    ' the girders carry the averaged slab tc (0.28375), so both beams get the slab difference added: their bottom
    ' then stays where 1-2 puts it (owner 2026-11-03). The OUTER end beam is deeper: deck top down to the bearing
' bottom = girder H + tc + bearing (link) height from the pop-up (owner 2026-11-03; 1-2 G_DIS_ENLEME_YUKSEKLIGI not used)
    enlH = enlH - tMin + tc

    ' concrete E (owner 2026-11-07): staged model = AASHTO Ec (INPUT column C), it feeds the AASHTO girder
    ' design; dynamic model = Eurocode Ecm (column D). The same E sets the composite modular ratio eG / eS.
    Dim eCol As Long, eTxt As String
    eCol = IIf(dynMode, 3, 2): eTxt = IIf(dynMode, "Ecm (INPUT column D)", "AASHTO Ec (INPUT column C)")
    Set r = NmRange(wb12, "G_KIRIS_BETONU"): clsG = CStr(r.Cells(1, 1).Value)
    eG = Num(r.Cells(1, eCol).Value, "Kiris betonu " & eTxt)
    Set r = NmRange(wb12, "G_TABLIYE_BETONU"): clsS = CStr(r.Cells(1, 1).Value)
    eS = Num(r.Cells(1, eCol).Value, "Tabliye betonu " & eTxt)
    denG = Num(NmRange(wb12, "G_BIRIM_AGIRLIK_ONGERME").Cells(1).Value, "G_BIRIM_AGIRLIK_ONGERME")
    denS = Num(NmRange(wb12, "G_BIRIM_AGIRLIK_BETONARME").Cells(1).Value, "G_BIRIM_AGIRLIK_BETONARME")

    qBal = Num(NmRange(wb12, "G_YUK_BALAST").Cells(k).Value, "G_YUK_BALAST span " & k)
    qKor = Num(NmRange(wb12, "G_YUK_KORUMA_BETONU").Cells(k).Value, "G_YUK_KORUMA_BETONU span " & k)
    qBor = Num(NmRange(wb12, "G_YUK_BORDUR").Cells(k).Value, "G_YUK_BORDUR span " & k)
    qKork = Num(NmRange(wb12, "G_YUK_KORKULUK").Cells(k).Value, "G_YUK_KORKULUK span " & k)
    qCep = Num(NmRange(wb12, "G_YUK_CEPHE_PANELI").Cells(k).Value, "G_YUK_CEPHE_PANELI span " & k)
    qRay = Num(NmRange(wb12, "G_YUK_RAY").Cells(k).Value, "G_YUK_RAY span " & k)

    wG = Num(BelowHeader(wb12.Worksheets("2_YUKLER"), "Wg").Value, "2_YUKLER Wg (wind on structure)")
    wV = Num(BelowHeader(wb12.Worksheets("2_YUKLER"), "Wv").Value, "2_YUKLER Wv (wind on train)")
    alphaLM = Num(BelowHeader(wb12.Worksheets("2_YUKLER"), ChrW(945)).Value, "2_YUKLER alpha")
    ' centrifugal force (owner 2026-11-06): largest Qtk of Tablo 17, so the model carries the value of the report
    qtk = MaxQtk(wb12.Worksheets("2_YUKLER"))
    railH = Num(NmRange(wb12, "G_BALAST_YUKSEKLIGI").Cells(1).Value, "G_BALAST_YUKSEKLIGI") + _
            Num(NmRange(wb12, "G_KORUMA_BETONU_YUKSEKLIGI").Cells(1).Value, "G_KORUMA_BETONU_YUKSEKLIGI") + RAIL_OVER_BALLAST
    If qtk <= 0 Then warnText = warnText & "- Tablo 17 (2_YUKLER) shows no Qtk: fill INPUT B19 (R) and B20 (Lf). The CF case stays empty." & vbCrLf

    ' uniform temperature (owner 2026-11-03, option A): dT = the 1-2 INPUT values, reference temperature 0
    tuPos = Num(NmRange(wb12, "G_SICAKLIK_MAKS").Cells(1).Value, "G_SICAKLIK_MAKS")
    tuNeg = Num(NmRange(wb12, "G_SICAKLIK_MIN").Cells(1).Value, "G_SICAKLIK_MIN")
    If tuPos <= 0 Or tuNeg >= 0 Then warnText = warnText & "- Temperature: G_SICAKLIK_MAKS = " & tuPos & ", G_SICAKLIK_MIN = " & tuNeg & " (expected > 0 and < 0)." & vbCrLf
    ' the model is always built square (owner 2026-11-05: warn, do not stop)
    Dim skewV As Variant
    On Error Resume Next
    skewV = wb12.Names("G_VEREVLIK").RefersToRange.Cells(1).Value
    On Error GoTo 0
    If IsNumeric(skewV) Then
        If CDbl(skewV) <> 0 Then warnText = warnText & "- Skew G_VEREVLIK = " & skewV & ": this macro ignores skew, the model is built SQUARE." & vbCrLf
    End If

    ReadCombTables wb12.Worksheets("1_GIRIS")

    ' ---- 4-7 ----
    sark = Num(NmRange(wb12, "G_SARKIKLIK").Cells(1).Value, "G_SARKIKLIK") / 100#
    kyB = Num(LabelValue(ChapterSheet(wb12, "KDME"), 3, "Elastomer Yanal Kesme Rijitli", 19), "Elastomer Yanal Kesme Rijitligi")
    kEB = Num(LabelValue(ChapterSheet(wb12, "KDME"), 3, "Elastomer Eksenel Rijitli", 19), "Elastomer Eksenel Rijitligi")
    If kyB <= 0 Or kEB <= 0 Or sark <= 0 Then Err.Raise vbObjectError + 3, , "4-7 bearing values must be > 0."
    If zoneL <= sark Or zoneL >= (Lc + 2 * sark) / 2 Then Err.Raise vbObjectError + 13, , "End-zone length " & zoneL & " m must lie between the bearing (" & sark & " m) and midspan."
    ReadStrands wb12
End Sub

' "C45" or "C45/55" -> 45 (MPa)
Private Function ClassFck(ByVal cls As String) As Double
    Dim c As String: c = UCase$(Replace(Trim$(cls), " ", ""))
    If Left$(c, 1) = "C" Then
        c = Mid$(c, 2): If InStr(c, "/") > 0 Then c = Left$(c, InStr(c, "/") - 1)
        ClassFck = Val(c)
    End If
    If ClassFck <= 0 Then Err.Raise vbObjectError + 14, , "Concrete class '" & cls & "' not understood."
End Function

' 4-7 1_RAPOR Tablo 27 (live-load stations): header row with 'Yeri' in column H and 'V (kN)' in column N;
' the rows below it hold the station x (m from the bearing) in column K until column D is empty
' span sheet of the 4A chapter, e.g. 4A_KDME
Private Function ChapterSheet(wb As Workbook, ByVal sfx As String) As Worksheet
    Dim nm As String
    nm = "4" & Chr$(64 + spanNo) & "_" & sfx
    On Error Resume Next
    Set ChapterSheet = wb.Worksheets(nm)
    On Error GoTo 0
    If ChapterSheet Is Nothing Then Err.Raise vbObjectError + 21, , "Sheet " & nm & " not found: fill the span in INPUT and press the chapter button."
End Function

Private Sub ReadLLStations(wb As Workbook)
    Dim ws As Worksheet, r As Long, hdr As Long, last As Long
    Set ws = ChapterSheet(wb, "KDME")
    last = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row
    For r = 1 To last
        If Left$(Trim$(CStr(IIf(IsError(ws.Cells(r, 8).Value), "", ws.Cells(r, 8).Value))), 4) = "Yeri" And _
           Left$(Trim$(CStr(IIf(IsError(ws.Cells(r, 14).Value), "", ws.Cells(r, 14).Value))), 6) = "V (kN)" Then hdr = r: Exit For
    Next r
    If hdr = 0 Then Err.Raise vbObjectError + 16, , "4-7 1_RAPOR: Tablo 27 header ('Yeri' in column H, 'V (kN)' in column N) not found."
    nLLst = 0: ReDim llX(1 To 30)
    For r = hdr + 1 To hdr + 30
        If Trim$(CStr(IIf(IsError(ws.Cells(r, 4).Value), "", ws.Cells(r, 4).Value))) = "" Then Exit For
        nLLst = nLLst + 1
        llX(nLLst) = Num(ws.Cells(r, 11).Value, "4-7 1_RAPOR Tablo 27 station x (column K, row " & r & ")")
    Next r
    If nLLst = 0 Then Err.Raise vbObjectError + 16, , "4-7 1_RAPOR: Tablo 27 has no stations below row " & hdr & "."
End Sub

' 4-7: strand rows + debonding from KILIFLAMA (B3 "SIRA", O3 "Halatlar"; C5:C14 row height cm,
' O5:O14 strands per row, D4:N4 stations in m from the bearing, D5:N14 debonded strands per row
' at each station), fpi / strand area / humidity from 1_RAPOR, fpu / Ep / diameter from INPUT
Private Sub ReadStrands(wb As Workbook)
    Dim ws As Worksheet, r As Long, j As Long, n As Long, d As Double
    Set ws = ChapterSheet(wb, "KILIF")
    If UCase$(Trim$(CStr(ws.Range("B3").Value))) <> "SIRA" Or LCase$(Trim$(CStr(ws.Range("O3").Value))) <> "halatlar" Then _
        Err.Raise vbObjectError + 15, , "4-7 KILIFLAMA: layout changed (B3 'SIRA' / O3 'Halatlar' expected)."
    nDb = 0
    For j = 4 To 14
        If IsNumeric(ws.Cells(4, j).Value) And Not IsEmpty(ws.Cells(4, j).Value) Then nDb = j - 3 Else Exit For
    Next j
    If nDb < 1 Then Err.Raise vbObjectError + 15, , "4-7 KILIFLAMA: no stations in D4:N4."
    ReadLLStations wb
    ReDim dbX(1 To nDb): ReDim rowZ(1 To 10): ReDim rowN(1 To 10): ReDim dbN(1 To 10, 1 To nDb)
    For j = 1 To nDb: dbX(j) = sark + CDbl(ws.Cells(4, j + 3).Value): Next j
    nRow = 0
    For r = 5 To 14
        n = CLng(NumOr0(ws.Cells(r, 15).Value))
        If n > 0 Then
            nRow = nRow + 1
            rowZ(nRow) = Num(ws.Cells(r, 3).Value, "KILIFLAMA C" & r & " (strand row height)") / 100#
            rowN(nRow) = n
            For j = 1 To nDb
                dbN(nRow, j) = CLng(NumOr0(ws.Cells(r, j + 3).Value))
                If dbN(nRow, j) > n Then Err.Raise vbObjectError + 15, , "4-7 KILIFLAMA row " & r & ": more debonded than strands."
                If j > 1 Then
                    If dbN(nRow, j) > dbN(nRow, j - 1) Then Err.Raise vbObjectError + 15, , "4-7 KILIFLAMA row " & r & ": debonded count grows along the girder."
                End If
            Next j
            If dbN(nRow, nDb) <> 0 Then Err.Raise vbObjectError + 15, , "4-7 KILIFLAMA row " & r & ": strands still debonded at the last station."
        End If
    Next r
    If nRow = 0 Then Err.Raise vbObjectError + 15, , "4-7 KILIFLAMA: no strands in O5:O14."
    fpi = Num(LabelValue(ChapterSheet(wb, "KDME"), 3, "1 Halattaki ilk gerilme", 15), "4-7 1_RAPOR fpi (1 Halattaki ilk gerilme)")
    humid = Num(LabelValue(ChapterSheet(wb, "KDME"), 3, "Ortalama ba", 15), "4-7 1_RAPOR Ortalama bagil nem")
    ' f'ci / f'c: 1_RAPOR row labelled '(Ongerme Aninda)' in column Q; f'c in J, f'ci in N (J40 / N40 today)
    Dim rFci As Long
    rFci = FindLabel(ChapterSheet(wb, "KDME"), 17, "(" & ChrW(214) & "ngerme An").Row
    fciRatio = Num(ChapterSheet(wb, "KDME").Cells(rFci, 14).Value, "4-7 1_RAPOR f'ci (column N, row '(Ongerme Aninda)')") / _
               Num(ChapterSheet(wb, "KDME").Cells(rFci, 10).Value, "4-7 1_RAPOR f'c (column J, row '(Ongerme Aninda)')")
    If fciRatio <= 0.2 Or fciRatio >= 1 Then Err.Raise vbObjectError + 15, , "4-7 f'ci / f'c = " & fciRatio & " is out of range (0.2-1)."
    ageG = ACI_A_G * fciRatio / (1 - ACI_B_G * fciRatio)   ' t from t / (A + B t) = f'ci / f'c
    aStr = Num(LabelValue(ChapterSheet(wb, "KDME"), 11, "Ahalat", 14), "4-7 1_RAPOR Ahalat") / 1000000#
    fpuS = Num(NmRange(wb, "G_HALAT_CEKME_DAYANIMI").Cells(1).Value, "4-7 G_HALAT_CEKME_DAYANIMI")
    epS = Num(NmRange(wb, "G_HALAT_ELASTISITE").Cells(1).Value, "4-7 G_HALAT_ELASTISITE")
    d = Num(NmRange(wb, "G_HALAT_CAPI").Cells(1).Value, "4-7 G_HALAT_CAPI")
    dStr = IIf(d < 2, d * 0.0254, d / 1000#)          ' inch (0.6) or mm
    If fpi <= 0 Or fpi > fpuS Or aStr <= 0 Then Err.Raise vbObjectError + 15, , "4-7 strand values out of range (fpi " & fpi & ", fpu " & fpuS & ", Ahalat " & aStr & ")."
End Sub

' isY = end-zone girder (fills kbY; the height must equal the span girder)
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

Private Function ReadEnlemeH(wb As Workbook) As Double
    Dim r As Range, i As Long, t As String
    On Error Resume Next
    Set r = wb.Names("G_ENLEME_YUKSEKLIGI").RefersToRange
    On Error GoTo 0
    If Not r Is Nothing Then ReadEnlemeH = Num(r.Cells(1).Value, "G_ENLEME_YUKSEKLIGI"): Exit Function
    Set r = NmRange(wb, "G_ENLEME_KALINLIGI").Cells(1)
    For i = 1 To 5
        t = LCase$(CStr(r.Offset(i, -1).Value))
        If Left$(t, 8) = "enleme y" Then ReadEnlemeH = Num(r.Offset(i, 0).Value, "Enleme yuksekligi (INPUT)"): Exit Function
    Next i
    Err.Raise vbObjectError + 7, , "Enleme yuksekligi not found (row below 'Enleme Kalinligi', or name G_ENLEME_YUKSEKLIGI)."
End Function

Private Sub ReadCombTables(ws As Worksheet)
    Dim hc As Range, hr As Long, c As Long, t As String, cDC As Long, cLL As Long, cWA As Long, cWS As Long, cWL As Long, cTU As Long
    Dim r As Long, lab As String, g As Range
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
    Next c
    If cDC * cLL * cWA * cWS * cWL * cTU = 0 Then Err.Raise vbObjectError + 8, , "Tablo 14 header columns (DC, LL, WA, WS, WL, TU) not all found."
    Set facRows = CreateObject("Scripting.Dictionary")
    facRows.CompareMode = vbBinaryCompare
    For r = hr + 1 To hr + 20
        lab = Replace(Trim$(CStr(ws.Cells(r, 3).Value)), " ", "")
        If lab <> "" Then facRows(lab) = Array(Fac(ws.Cells(r, cDC).Value), Fac(ws.Cells(r, cLL).Value), Fac(ws.Cells(r, cWA).Value), _
                                             Fac(ws.Cells(r, cWS).Value), Fac(ws.Cells(r, cWL).Value), Fac(ws.Cells(r, cTU).Value))
    Next r
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

' ============================================================================
'  MODEL
' ============================================================================
Private Sub BuildModel()
    Dim m As Long, dx As Double, i As Long, j As Long, cand As Object, W As Double, yc As Double
    Dim ysC As Object, ysF As Object, xs As Variant, a As Long, b As Long, x As Double, y As Double

    Set nodeKey = CreateObject("Scripting.Dictionary")
    Set loadItems = CreateObject("Scripting.Dictionary")
    Set loadIds = CreateObject("Scripting.Dictionary")
    Set crossN = CreateObject("Scripting.Dictionary")
    Set girdN = CreateObject("Scripting.Dictionary"): Set deckN = CreateObject("Scripting.Dictionary")
    nCnt = 0: eCnt = 0: crossE = "": girdE = "": deckE = ""
    ReDim nX(1 To 64): ReDim nY(1 To 64): ReDim nZ(1 To 64)
    ReDim eMat(1 To 64): ReDim eSec(1 To 64): ReDim eN1(1 To 64): ReDim eN2(1 To 64)

    Lg = Lc + 2 * sark
    m = Int(Lc / stripDx + 0.5): If m < 2 Then m = 2
    dx = Lc / m
    dxAct = dx
    ' Hambly: every transverse member carries the slab half way to its neighbours. End line (X = 0) ->
    ' half way to the diaphragm line; diaphragm line (X = sark) -> from there to half way to the first strip.
    endStripW = sark / 2: If endStripW < 0.01 Then endStripW = 0.01
    diaW = sark / 2 + dx / 2
    ' visual only: first / last strip drawn wide enough to reach the diaphragm web (no drawn gap); SECF scales
    ' every property back to the true width dx, so the structure is unchanged (owner 2026-11-03)
    visW = 2 * dx - IIf(DiaFlange() > enlB, DiaFlange(), enlB)
    If visW < dx + 0.001 Then visW = 0
    endExtra = sark + Hg / 2                    ' bearing + girder H/2 (owner 2026-11-03)

    ' girder stations
    Set cand = CreateObject("Scripting.Dictionary")
    AddX cand, 0: AddX cand, sark: AddX cand, Lg - sark: AddX cand, Lg
    If endExtra < Lg / 2 Then AddX cand, endExtra: AddX cand, Lg - endExtra
    AddX cand, zoneL: AddX cand, Lg - zoneL                 ' end zone (section over the support)
    For i = 1 To m - 1: AddX cand, sark + i * dx: Next i
    ' 4-7 Tablo 27 live-load stations (from the bearing), both halves: girder nodes there so MAK_LL reads
    ' node results, no interpolation (owner 2026-11-03); AddX skips those already on a strip / other node
    For i = 1 To nLLst
        If llX(i) > 0 And sark + llX(i) < Lg - sark Then AddX cand, sark + llX(i): AddX cand, Lg - sark - llX(i)
    Next i
    st = SortedKeys(cand): nSt = UBound(st)

    ' deck and tracks
    W = kSol + h1 + h2 + kSag
    yc = (nG - 1) * sG / 2
    yL = yc - W / 2: yR = yc + W / 2
    nTr = 1: tr(1) = yL + kSol + h1 / 2
    If h2 > 0 Then nTr = 2: tr(2) = yL + kSol + h1 + h2 / 2
    For i = 1 To nTr
        trG(i) = Int(tr(i) / sG + 0.5)
        If trG(i) < 0 Then trG(i) = 0
        If trG(i) > nG - 1 Then trG(i) = nG - 1
    Next i

    Set ysC = CreateObject("Scripting.Dictionary"): Set ysF = CreateObject("Scripting.Dictionary")
    For j = 0 To nG - 1: AddX ysC, j * sG: AddX ysF, j * sG: Next j
    For i = 1 To nTr
        If tr(i) > 0.001 And tr(i) < (nG - 1) * sG - 0.001 Then AddX ysC, tr(i)
        AddX ysF, tr(i)
    Next i
    If yL < -0.001 Then AddX ysF, yL
    If yR > (nG - 1) * sG + 0.001 Then AddX ysF, yR
    ovL = IIf(yL < -0.001, -yL, 0): ovR = IIf(yR > (nG - 1) * sG + 0.001, yR - (nG - 1) * sG, 0)

    ' girders (material 1; section 7 in the span, 8 in the end zones; edge girders with their overhang slab:
    ' left 9/10, right 11/12 - Hambly: the cantilever slab is part of the edge girder's flange)
    ReDim gEl(0 To nG - 1, 1 To nSt - 1)
    Dim sc As Long, na As Long, nb As Long
    For j = 0 To nG - 1
        For i = 1 To nSt - 1
            sc = IIf((st(i) + st(i + 1)) / 2 < zoneL Or (st(i) + st(i + 1)) / 2 > Lg - zoneL, 8, 7)
            If j = 0 And ovL > 0 Then sc = sc + 2
            If j = nG - 1 And ovR > 0 Then sc = sc + 4
            na = AddNode(st(i), j * sG, 0): nb = AddNode(st(i + 1), j * sG, 0)
            gEl(j, i) = AddElem(1, sc, na, nb)
            girdE = girdE & IIf(girdE = "", "", ",") & gEl(j, i)
            girdN(na) = 1: girdN(nb) = 1
        Next i
    Next j
    ' cross beams (2/2 = diaphragm T with its slab flange), strips (3/3), end lines (3/5),
    ' overhang slab on the diaphragm lines (3/13, same tributary width as the diaphragm flange)
    TransLine sark, ysC, 2, 2
    TransLine Lg - sark, ysC, 2, 2
    Dim yo As Object
    If ovL > 0 Then
        Set yo = CreateObject("Scripting.Dictionary"): AddX yo, yL: AddX yo, 0
        TransLine sark, yo, 3, 13: TransLine Lg - sark, yo, 3, 13
    End If
    If ovR > 0 Then
        Set yo = CreateObject("Scripting.Dictionary"): AddX yo, (nG - 1) * sG: AddX yo, yR
        TransLine sark, yo, 3, 13: TransLine Lg - sark, yo, 3, 13
    End If
    For i = 1 To m - 1: TransLine sark + i * dx, ysF, 3, IIf((i = 1 Or i = m - 1) And visW > 0, 14, 3): Next i
    EndLine 0, ysF, yc
    EndLine Lg, ysF, yc
    ' edge lines (3/4)
    Dim xe As Object: Set xe = CreateObject("Scripting.Dictionary")
    AddX xe, 0: AddX xe, Lg: AddX xe, sark: AddX xe, Lg - sark
    For i = 1 To m - 1: AddX xe, sark + i * dx: Next i
    xs = SortedKeys(xe)
    If yL < -0.001 Then
        For i = 1 To UBound(xs) - 1: DeckElem AddElem(3, 4, AddNode(xs(i), yL, 0), AddNode(xs(i + 1), yL, 0)): Next i
    End If
    If yR > (nG - 1) * sG + 0.001 Then
        For i = 1 To UBound(xs) - 1: DeckElem AddElem(3, 4, AddNode(xs(i), yR, 0), AddNode(xs(i + 1), yR, 0)): Next i
    End If

    ' supports
    Dim zMid As Double, top As Long, mid As Long, bot As Long, e As Variant, kLong As Double, nL As Long
    zMid = -(Hg + tc)
    rigdJson = "": elnkJson = "": consJson = "": nsprJson = ""
    For Each e In Array(sark, Lg - sark)
        x = CDbl(e)
        kLong = IIf(Abs(x - sark) < 0.0001, kFix, kyB)
        For j = 0 To nG - 1
            top = AddNode(x, j * sG, 0)
            If dynMode Then
                ' dynamic model (owner 2026-11-05): no rigid or elastic links; start line X Y, end line Y, rotations free;
                ' vertical = point spring kv (4-7 elastomer axial stiffness) on every girder bearing node
                consJson = consJson & IIf(consJson = "", "", ",") & QS(CStr(top)) & ":{""ITEMS"":[{""ID"":1,""GROUP_NAME"":""BG_MESNET"",""CONSTRAINT"":" & QS(IIf(Abs(x - sark) < 0.0001, "1100000", "0100000")) & "}]}"
                nsprJson = nsprJson & IIf(nsprJson = "", "", ",") & QS(CStr(top)) & ":{""ITEMS"":[{""ID"":1,""TYPE"":""LINEAR"",""FormType"":0,""SDR"":[0,0," & JN(kEB) & ",0,0,0]," & _
                    """F_S"":[false,false,false,false,false,false],""DAMPING"":false,""Cr"":[0,0,0,0,0,0],""EFFAREA"":0,""DK"":[0,0,0]}]}"
                GoTo NextGirder
            End If
            mid = AddNode(x, j * sG, zMid)
            bot = AddNode(x, j * sG, zMid - linkLen)
            girdN(mid) = 1: girdN(bot) = 1
            nL = nL + 1
            rigdJson = rigdJson & IIf(rigdJson = "", "", ",") & QS(CStr(top)) & ":{""ITEMS"":[{""ID"":1,""GROUP_NAME"":""BG_MESNET"",""DOF"":111111,""S_NODE"":[" & mid & "]}]}"
            elnkJson = elnkJson & IIf(elnkJson = "", "", ",") & QS(CStr(nL)) & ":{""NODE"":[" & mid & "," & bot & "],""ANGLE"":0,""LINK"":""GEN""," & _
                """R_S"":[false,false,false,false,false,false],""SDR"":[" & JN(kEB) & "," & JN(kFix) & "," & JN(kLong) & ",0,0,0],""bSHEAR"":false,""DR"":[0.5,0.5],""BNGR_NAME"":""BG_MESNET""}"
            consJson = consJson & IIf(consJson = "", "", ",") & QS(CStr(bot)) & ":{""ITEMS"":[{""ID"":1,""GROUP_NAME"":""BG_MESNET"",""CONSTRAINT"":""1111110""}]}"
            ' temporary roll restraint (owner 2026-11-03, option B): in CS1 each girder stands on one bearing point
            ' per end (link rotations free) and could roll about its axis; RX of the girder node is held in CS1 only
            ' (temporary bracing), released in CS2 when the deck ties the girders together
            consJson = consJson & "," & QS(CStr(top)) & ":{""ITEMS"":[{""ID"":1,""GROUP_NAME"":""BG_GECICI"",""CONSTRAINT"":""0001000""}]}"
NextGirder:
        Next j
    Next e

    BuildPolygon
    BuildLoads
End Sub

' girder-end lines (owner 2026-11-03, option 2): outer end beam (dis enleme, section 15, slab concrete, full
' rectangle disB x disH hanging from the deck top) over its length disL centred on the girders, tying the girder
' ends; its 0.5 x disB centroid offset beyond the girder ends is ignored (Hambly: member on the line where it
' connects). Beyond the beam ends the thin end slab strip (section 5) runs on to the deck edge.
Private Sub EndLine(ByVal x As Double, ysIn As Object, ByVal yMid As Double)
    Dim ys As Object, k As Variant, v As Variant, i As Long, e As Long, a As Long, b As Long, y1 As Double, y2 As Double, ym As Double
    Set ys = CreateObject("Scripting.Dictionary")
    For Each k In ysIn.Keys: AddX ys, CDbl(ysIn(k)): Next k
    y1 = yMid - disL / 2: y2 = yMid + disL / 2
    If y1 > yL + 0.001 Then AddX ys, y1 Else y1 = yL
    If y2 < yR - 0.001 Then AddX ys, y2 Else y2 = yR
    v = SortedKeys(ys)
    For i = 1 To UBound(v) - 1
        a = AddNode(x, v(i), 0): b = AddNode(x, v(i + 1), 0)
        ym = (v(i) + v(i + 1)) / 2
        If ym > y1 And ym < y2 Then e = AddElem(2, 15, a, b) Else e = AddElem(3, 5, a, b)
        crossE = crossE & IIf(crossE = "", "", ",") & e
        crossN(a) = 1: crossN(b) = 1
        DeckElem e
    Next i
End Sub

Private Sub TransLine(ByVal x As Double, ys As Object, ByVal mat As Long, ByVal sec As Long)
    Dim v As Variant, i As Long, e As Long, a As Long, b As Long
    v = SortedKeys(ys)
    For i = 1 To UBound(v) - 1
        a = AddNode(x, v(i), 0): b = AddNode(x, v(i + 1), 0)
        e = AddElem(mat, sec, a, b)
        crossE = crossE & IIf(crossE = "", "", ",") & e
        crossN(a) = 1: crossN(b) = 1
        DeckElem e
    Next i
End Sub

' transverse / edge elements: deck structure group (cast in CS2)
Private Sub DeckElem(ByVal e As Long)
    deckE = deckE & IIf(deckE = "", "", ",") & e
    deckN(eN1(e)) = 1: deckN(eN2(e)) = 1
End Sub

Private Sub BuildPolygon()
    PolyOf kb, polyArea, polyJson, polyPer, polyPts
    PolyOf kbY, polyAreaY, polyJsonY, polyPerY, polyPtsY
End Sub

' KIRIS_INPUT: k(1..6) = b2..b7, k(7..12) = h2..h7 (m); X = width, Y = height from the soffit
Private Sub PolyOf(k() As Double, ByRef area As Double, ByRef json As String, ByRef per As Double, ByRef mp As String)
    Dim px(1 To 16) As Double, py(1 To 16) As Double, n As Long, i As Long, s As String, q As Long
    Dim yA As Double, yB As Double, yC As Double, yD As Double
    yA = Hg - k(7)                       ' under the top flange
    yB = yA - k(8)                       ' end of the first haunch
    yC = yB - k(9)                       ' top of the web
    yD = k(12)                           ' top of the bottom flange
    n = 0
    AddP px, py, n, k(3) / 2, 0
    AddP px, py, n, k(3) / 2, yD
    AddP px, py, n, k(2) / 2, yD + k(11)
    AddP px, py, n, k(2) / 2, yC
    AddP px, py, n, k(2) / 2 + k(4), yB
    AddP px, py, n, k(1) / 2, yA
    AddP px, py, n, k(1) / 2, Hg
    AddP px, py, n, -k(1) / 2, Hg
    AddP px, py, n, -k(1) / 2, yA
    AddP px, py, n, -(k(2) / 2 + k(4)), yB
    AddP px, py, n, -k(2) / 2, yC
    AddP px, py, n, -k(2) / 2, yD + k(11)
    AddP px, py, n, -k(3) / 2, yD
    AddP px, py, n, -k(3) / 2, 0
    area = 0: per = 0: s = "": mp = ""
    For i = 1 To n
        q = IIf(i = n, 1, i + 1)
        area = area + px(i) * py(q) - px(q) * py(i)
        per = per + Sqr((px(q) - px(i)) ^ 2 + (py(q) - py(i)) ^ 2)
        s = s & IIf(s = "", "", ",") & "{""X"":" & JN(px(i)) & ",""Y"":" & JN(py(i)) & "}"
        mp = mp & IIf(mp = "", "", ", ") & JN(px(i)) & ", " & JN(py(i))
    Next i
    area = Abs(area) / 2
    json = """OUTER_POLYGON"":[{""VERTEX"":[" & s & "]}]"
End Sub

Private Sub AddP(px() As Double, py() As Double, n As Long, ByVal x As Double, ByVal y As Double)
    If n > 0 Then
        If Abs(px(n) - x) < 0.000001 And Abs(py(n) - y) < 0.000001 Then Exit Sub
    End If
    n = n + 1: px(n) = x: py(n) = y
End Sub

Private Sub BuildLoads()
    Dim i As Long, nR As Long, rails(1 To 4) As Double, k As Long, j As Long, lcs As Variant, lc As Variant
    Dim eL As Double, eR As Double, mx As Double, ltot As Double, xm As Double
    eL = yL + kSol / 2
    eR = yR - kSag / 2 - (nG - 1) * sG
    For k = 1 To nTr
        nR = nR + 1: rails(nR) = tr(k) - RAIL_HALF_GAUGE
        nR = nR + 1: rails(nR) = tr(k) + RAIL_HALF_GAUGE
    Next k
    For Each lc In Array("DW_1", "DW_2", "WS-Y")
        For i = 1 To nSt - 1
            If lc = "DW_1" Then
                For j = 0 To nG - 1: AddLoad gEl(j, i), CStr(lc), "GZ", -1, "", 0, -qBal: Next j
                AddLoad gEl(0, i), CStr(lc), "GZ", 1, "LY", eL, -qBor * nG / 2
                AddLoad gEl(0, i), CStr(lc), "GZ", 1, "LY", eL, -qKork * nG / 2
                AddLoad gEl(0, i), CStr(lc), "GZ", 1, "LY", eL, -qCep * nG / 2
                AddLoad gEl(nG - 1, i), CStr(lc), "GZ", 1, "LY", eR, -qBor * nG / 2
                AddLoad gEl(nG - 1, i), CStr(lc), "GZ", 1, "LY", eR, -qKork * nG / 2
                AddLoad gEl(nG - 1, i), CStr(lc), "GZ", 1, "LY", eR, -qCep * nG / 2
            End If
            If lc = "DW_2" Then
                For j = 0 To nG - 1: AddLoad gEl(j, i), CStr(lc), "GZ", -1, "", 0, -qKor: Next j
                For k = 1 To nR
                    j = NearestGirder(rails(k))
                    AddLoad gEl(j, i), CStr(lc), "GZ", -1, "", 0, -qRay * nG / nR
                Next k
            End If
            If lc = "WS-Y" Then AddLoad gEl(0, i), CStr(lc), "LY", 0, "LZ", 0, wG
            If lc = "WS-Y" Then AddLoad gEl(trG(1), i), "WL", "LY", 1, "LZ", WL_ECC, wV
        Next i
    Next lc
    ' nosing on the track-1 girder around midspan
    For i = 1 To nSt - 1
        xm = (st(i) + st(i + 1)) / 2
        If Abs(xm - Lg / 2) <= NOSING_LEN / 2 + 0.000001 Then ltot = ltot + (st(i + 1) - st(i))
    Next i
    If ltot <= 0 Then Err.Raise vbObjectError + 9, , "No girder element found for the nosing load."
    For i = 1 To nSt - 1
        xm = (st(i) + st(i + 1)) / 2
        If Abs(xm - Lg / 2) <= NOSING_LEN / 2 + 0.000001 Then AddLoad gEl(trG(1), i), "MOVE", "LY", 1, "LZ", NOSING_ECC, NOSING_Q * alphaLM / ltot
    Next i
    ' centrifugal force: Qtk per track on the girder under each track, whole girder length, same direction as
    ' WL (the grillage is straight, so the two add up), lever from the slab-top node to rail level + 1.8 m
    If qtk > 0 Then
        For k = 1 To nTr
            For i = 1 To nSt - 1
                AddLoad gEl(trG(k), i), "CF", "LY", 1, "LZ", railH + CF_HEIGHT, qtk
            Next i
        Next k
    End If
End Sub

Private Function NearestGirder(ByVal y As Double) As Long
    Dim j As Long: j = Int(y / sG + 0.5)
    If j < 0 Then j = 0
    If j > nG - 1 Then j = nG - 1
    NearestGirder = j
End Function

' eccType: -1 = no eccentricity, 0/1 = ECCEN_TYPE
Private Sub AddLoad(ByVal el As Long, ByVal lc As String, ByVal dirn As String, ByVal eccType As Long, ByVal eccDir As String, ByVal ecc As Double, ByVal q As Double)
    Dim id As Long, s As String
    id = loadIds(el) + 1: loadIds(el) = id
    s = "{""ID"":" & id & ",""LCNAME"":" & QS(lc) & ",""GROUP_NAME"":" & QS(IIf(lc = "DW_1" Or lc = "DW_2", "LG_DW", "")) & ",""CMD"":""BEAM"",""TYPE"":""UNILOAD"",""DIRECTION"":" & QS(dirn) & _
        ",""USE_PROJECTION"":false,""USE_ECCEN"":" & IIf(eccType >= 0, "true", "false")
    If eccType >= 0 Then s = s & ",""ECCEN_TYPE"":" & eccType & ",""ECCEN_DIR"":" & QS(eccDir) & ",""I_END"":" & JN(ecc) & ",""J_END"":" & JN(ecc) & ",""USE_J_END"":false"
    s = s & ",""D"":[0,1,0,0],""P"":[" & JN(q) & "," & JN(q) & ",0,0],""USE_ADDITIONAL"":false,""ADDITIONAL_I_END"":0,""ADDITIONAL_J_END"":0,""USE_ADDITIONAL_J_END"":false}"
    If loadItems.Exists(el) Then loadItems(el) = loadItems(el) & "," & s Else loadItems(el) = s
End Sub

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

' ============================================================================
'  PAYLOADS
' ============================================================================
Private Sub BuildPayloads()
    Dim s As String, i As Long, j As Long, k As Variant, cnt As Long
    Set payloads = New Collection
    Add "db/UNIT", "{""Assign"":{""1"":{""FORCE"":""KN"",""DIST"":""M"",""HEAT"":""J"",""TEMPER"":""C""}}}"
    Add "db/STYP", "{""Assign"":{""1"":{""STYP"":0,""MASS"":1,""GRAV"":" & JN(GRAV_STYP) & ",""TEMP"":0,""bALIGNBEAM"":false,""bALIGNSLAB"":false,""bMASSOFFSET"":false,""bSELFWEIGHT"":true,""SMASS"":1}}}"
    ' PJCF: Civil NX rejects longer than 80 characters and (not always) Turkish letters -> ASCII, 80 max
    Add "db/PJCF", "{""Assign"":{""1"":{""PROJECT"":" & QS(Pj(projName)) & ",""USER"":" & QS(Pj(Application.UserName)) & ",""CLIENT"":" & QS(Pj(clientName)) & _
        ",""TITLE"":" & QS(Pj(bridgeName)) & ",""COMMENT"":" & QS(Pj("Hesap Acikligi - " & spanNo & " = " & Lc & " m; " & SCRIPT_ID & " " & SCRIPT_VERSION & IIf(dynMode, " DYNAMIC", ""))) & "}}}"
    Add "db/MATL", "{""Assign"":{" & _
        ConcMat(1, clsG, eG, denG) & "," & ConcMat(2, clsS, eS, denS) & "," & _
        """3"":{""TYPE"":""CONC"",""NAME"":""Dummy"",""HE_SPEC"":0,""HE_COND"":0,""bMASS_DENS"":false,""DAMP_RAT"":0.05,""PARAM"":[{""P_TYPE"":2,""ELAST"":" & JN(eS * 1000#) & ",""POISN"":0.2,""THERMAL"":0.00001,""DEN"":0}]}," & _
        """4"":{""TYPE"":""STEEL"",""NAME"":""HALAT"",""HE_SPEC"":0,""HE_COND"":0,""bMASS_DENS"":false,""DAMP_RAT"":0.02,""PARAM"":[{""P_TYPE"":2,""ELAST"":" & JN(epS * 1000#) & ",""POISN"":0.3,""THERMAL"":0.000012,""DEN"":78.5}]}}}"
    ' time-dependent concrete (owner 2026-11-03): AASHTO creep/shrinkage (AASHTO 2020 5.4.2.3), strength gain
    ' ACI 209 t / (A + B t) (MIDAS has no AASHTO strength curve), linked to materials 1 (girder) and 2 (slab).
    ' V/S (m): girder = total precast surface (AASHTO 5.4.2.3.2), slab = both faces (t / 2)
    If Not dynMode Then
        Add "db/TDMT", "{""Assign"":{""1"":" & TdmtJson("AASHTO_KIRIS", fckG, ageG, VsG()) & ",""2"":" & TdmtJson("AASHTO_TABLIYE", fckS, AGE_SLAB, tc / 2) & "}}"
        Add "db/TDME", "{""Assign"":{""1"":{""NAME"":""ACI_KIRIS"",""TYPE"":""CODE"",""CODENAME"":""ACI"",""STRENGTH"":" & JN(fckG * 1000#) & ",""A"":" & JN(ACI_A_G) & ",""B"":" & JN(ACI_B_G) & "}," & _
            """2"":{""NAME"":""ACI_TABLIYE"",""TYPE"":""CODE"",""CODENAME"":""ACI"",""STRENGTH"":" & JN(fckS * 1000#) & ",""A"":" & JN(ACI_A_S) & ",""B"":" & JN(ACI_B_S) & "}}}"
        Add "db/TMAT", "{""Assign"":{""1"":{""TDMT_NAME"":""AASHTO_KIRIS"",""TDME_NAME"":""ACI_KIRIS""},""2"":{""TDMT_NAME"":""AASHTO_TABLIYE"",""TDME_NAME"":""ACI_TABLIYE""}}}"
    End If
    ' Hambly: slab members use the slab thickness tc of the composite girder; diaphragm = T (web + slab flange)
    ' diaphragm flange (Hambly 4.4.3): to midway (diaW), but shear lag -> at most 0.3 x girder spacing;
    ' not wider than the web -> plain rectangle (support diaphragm, slab flange ignored as in Hambly 4.5.1)
    Dim fl As Double: fl = diaW: If fl > 0.3 * sG Then fl = 0.3 * sG
    Add "db/SECT", "{""Assign"":{" & IIf(fl > enlB + 0.001, TSect(2, "ENLEME", enlH, fl, enlB, tc), SbSect(2, "ENLEME", enlH, enlB)) & "," & SbSect(3, "DUMMY1", tc, dxAct) & "," & SbSect(4, "DUMMY2", 0.001, 0.001) & "," & _
        SbSect(5, "DUMMY3", tc, endStripW) & "," & PscSect(6) & "," & _
        SbSect(13, "DUMMY4", tc, diaW) & IIf(visW > 0, "," & SbSect(14, "DUMMY1_VIS", tc, visW), "") & "," & SbSect(15, "DIS_ENLEME", DisHgt(), disB) & "}}"
    Add "db/CO_S", SectColorJson()

    s = ""
    For i = 1 To nCnt: s = s & IIf(i = 1, "", ",") & QS(CStr(i)) & ":{""X"":" & JN(nX(i)) & ",""Y"":" & JN(nY(i)) & ",""Z"":" & JN(nZ(i)) & "}": Next i
    Add "db/NODE", "{""Assign"":{" & s & "}}"
    s = ""
    For i = 1 To eCnt
        s = s & IIf(i = 1, "", ",") & QS(CStr(i)) & ":{""TYPE"":""BEAM"",""MATL"":" & eMat(i) & ",""SECT"":" & eSec(i) & _
            ",""NODE"":[" & eN1(i) & "," & eN2(i) & ",0,0,0,0,0,0],""ANGLE"":0,""STYPE"":0}"
    Next i
    Add "db/ELEM", "{""Assign"":{" & s & "}}"
    s = ""
    For Each k In SortedKeysL(crossN): s = s & IIf(s = "", "", ",") & k: Next k
    Add "db/GRUP", "{""Assign"":{""1"":{""NAME"":""CROSS_BEAM"",""P_TYPE"":0,""N_LIST"":[" & s & "],""E_LIST"":[" & crossE & "]}," & _
        """2"":{""NAME"":""SG_KIRIS"",""P_TYPE"":0,""N_LIST"":[" & JoinKeys(girdN) & "],""E_LIST"":[" & girdE & "]}," & _
        """3"":{""NAME"":""SG_TABLIYE"",""P_TYPE"":0,""N_LIST"":[" & JoinKeys(deckN) & "],""E_LIST"":[" & deckE & "]}}}"
    If dynMode Then
        Add "db/BNGR", "{""Assign"":{""1"":{""NAME"":""BG_MESNET"",""AUTOTYPE"":0}}}"
    Else
        Add "db/BNGR", "{""Assign"":{""1"":{""NAME"":""BG_MESNET"",""AUTOTYPE"":0},""2"":{""NAME"":""BG_GECICI"",""AUTOTYPE"":0},""3"":{""NAME"":""BG_SF"",""AUTOTYPE"":0}}}"
    End If
    Add "db/CONS", "{""Assign"":{" & consJson & "}}"
    If dynMode Then
        Add "db/NSPR", "{""Assign"":{" & nsprJson & "}}"
    Else
        Add "db/RIGD", "{""Assign"":{" & rigdJson & "}}"
        Add "db/ELNK", "{""Assign"":{" & elnkJson & "}}"
    End If
    ' Hambly torsion of the slab members (sections 3, 5, 13) and true strip of the drawn section 14: section
    ' stiffness factors. Staged analysis used to ignore ESSF and any ungrouped factor; since Civil NX 2026 v2.3 an
    ' ungrouped ESSF is applied in PostCS (measured 2026-11-07), an ungrouped SECF is still ignored, and a grouped
    ' ESSF still stops the analysis. db/SECF works in a boundary group
    ' activated in a stage (BG_SF, CS2 with the deck), measured 2026-11-05. IPART 1 = plain section.
    s = SlabSecf(3) & "," & SlabSecf(5) & "," & SlabSecf(13)
    If visW > 0 Then s = s & "," & VisSecf()
    If dynMode Then s = s & DYN_TAG   ' girder factors: known after the measurement in SendAll
    Add "db/SECF", "{""Assign"":{" & s & "}}"
    Add "db/STLD", "{""Assign"":{" & StldJson() & "}}"
    ' TU+ / TU-: uniform temperature change on the whole model (system temperature; STYP initial temperature 0)
    Add "db/STMP", "{""Assign"":{""1"":{""TEMPER"":" & JN(tuPos) & ",""LCNAME"":""TU+"",""GROUP_NAME"":""""},""2"":{""TEMPER"":" & JN(tuNeg) & ",""LCNAME"":""TU-"",""GROUP_NAME"":""""}}}"
    Add "db/LDGR", "{""Assign"":{""1"":{""NAME"":""LG_ZATI""},""2"":{""NAME"":""LG_ONGERME""},""3"":{""NAME"":""LG_DW""}}}"
    Add "db/BODF", "{""Assign"":{""1"":{""LCNAME"":""DC"",""GROUP_NAME"":""LG_ZATI"",""FV"":[0,0,-1]}}}"
    s = "": cnt = 0
    For Each k In SortedKeysL(loadItems)
        s = s & IIf(s = "", "", ",") & QS(CStr(k)) & ":{""ITEMS"":[" & loadItems(k) & "]}"
        cnt = cnt + 1
        If cnt = BMLD_CHUNK Then Add "db/BMLD", "{""Assign"":{" & s & "}}": s = "": cnt = 0
    Next k
    If s <> "" Then Add "db/BMLD", "{""Assign"":{" & s & "}}"
    Add "db/LTOM", "{""Assign"":{""1"":{""DIR"":""XYZ"",""bNODAL"":true,""bBEAM"":true,""bFLOOR"":true,""bPRES"":true,""GRAV"":" & JN(GRAV) & _
        ",""vLC"":[{""LCNAME"":""DW_1"",""FACTOR"":1},{""LCNAME"":""DW_2"",""FACTOR"":1}]}}}"
    Add "db/ACTL", "{""Assign"":{""1"":{""ARDC"":true,""ANRC"":true,""CSECF"":false,""TRS"":true,""BMSTRESS"":false,""CRBAR"":false,""CLATS"":false,""ITER"":20,""TOL"":0.001}}}"
    Add "db/EIGV", "{""Assign"":{""1"":{""TYPE"":""LANCZOS"",""iFREQ"":" & N_MODES & ",""bMINMAX"":false,""bSTRUM"":false}}}"
    Add "db/THGC", "{""Assign"":{""1"":{""ENERGYRESULT"":true,""SDVI"":false,""SDVE"":false,""SDST"":false,""SDHY"":false,""SDIS"":false,""bMSSSTATUS"":true," & _
        """GNT"":0,""ILT"":0,""IEPI"":true,""NSTEP"":1,""bROT"":false,""SNIO"":1,""bPCF"":true,""MAXNS"":16,""MAXIT"":10,""bDN"":true,""bFN"":false,""bEN"":false,""DN"":0.001,""bULSM"":true,""ULSM"":3}}}"
    Add "db/MVCD", "{""Assign"":{""1"":{""CODE"":""EUROCODE""}}}"
    s = LaneJson("LANE2", 1)
    If nTr = 2 Then s = LaneJson("LANE1", 2) & "," & s
    Add "db/LLAN", "{""Assign"":{" & RenumberTop(s) & "}}"
    Add "db/MVHL", "{""Assign"":{" & _
        """1"":{""MVLD_CODE"":11,""VEHICLE_LOAD_NAME"":""Load Model 71"",""VEHICLE_LOAD_NUM"":1,""VEHICLE_TYPE_NAME"":""Load Model 71"",""VEH_EUROCODE"":{""SUB_TYPE"":23,""W1"":80,""DD1"":0,""D1"":0.8,""W2"":80,""DD2"":0,""D2"":0.8,""V_LOAD_FACTOR"":" & JN(alphaLM) & ",""LONGI_DIST"":false,""ECCEN_VERT_LOAD"":false}}," & _
        """2"":{""MVLD_CODE"":11,""VEHICLE_LOAD_NAME"":""HSLM A1 ~ HSLM A10"",""VEHICLE_LOAD_NUM"":1,""VEHICLE_TYPE_NAME"":""HSLM A1 ~ HSLM A10"",""VEH_EUROCODE"":{""SUB_TYPE"":23,""SEL_VEHICLE"":""A8"",""V_LOAD_FACTOR"":1,""LONGI_DIST"":false,""ECCEN_VERT_LOAD"":false,""HSLMA_NUM"":12,""HSLMA_LENGTH"":25,""HSLMA_SPACING"":2.5,""HSLMA_FORCE"":190,""PHI_DYN_EFF1"":0,""PHI_DYN_EFF2"":0}}," & _
        """3"":{""MVLD_CODE"":11,""VEHICLE_LOAD_NAME"":""Load Model 71 - a=1"",""VEHICLE_LOAD_NUM"":1,""VEHICLE_TYPE_NAME"":""Load Model 71"",""VEH_EUROCODE"":{""SUB_TYPE"":23,""W1"":80,""DD1"":0,""D1"":0.8,""W2"":80,""DD2"":0,""D2"":0.8,""V_LOAD_FACTOR"":1,""LONGI_DIST"":false,""ECCEN_VERT_LOAD"":false}}}}"
    Add "db/MVLDeu", "{""Assign"":{" & MvCase(1, "LM71 - S", "Load Model 71", False) & "," & MvCase(2, "HSLM-A 8 - S", "HSLM A1 ~ HSLM A10", False) & "," & _
        MvCase(3, "LM71", "Load Model 71", True) & "," & MvCase(4, "HSLM-A 8", "HSLM A1 ~ HSLM A10", True) & "," & MvCase(5, "LM71 - a=1", "Load Model 71 - a=1", False) & "}}"
    Add "db/DYFG", "{""Assign"":{""1"":{""INPUT_TYPE"":0,""LENGTH"":" & JN(Lc) & ",""MAINTAIN_TYPE"":1,""OPT_REDUCE_EFF"":false,""HEIGHT_COVER"":0,""DYN_FACTOR"":" & JN(Phi3(Lc)) & "}}}"
    If Not dynMode Then BuildStages
    AddCombs "db/LCOM-GEN", False
    AddCombs "db/LCOM-CONC", True
End Sub

' ---- construction stages + pretensioned strands (owner 2026-11-02) ----
'  CS1 girders + bearings + self-weight + prestress (girder age ageG), CS2 deck (slab part of
'  the composite sections, diaphragms, strips), CS3 DW_1 + DW_2, CS4 long term
Private Sub BuildStages()
    Dim s As String
    s = """1"":" & Stage("CS1", DUR_CS1, """ACT_ELEM"":[{""GRUP_NAME"":""SG_KIRIS"",""AGE"":" & JN(ageG) & "}],""ACT_BNGR"":[{""BNGR_NAME"":""BG_MESNET"",""POS"":""ORIGINAL""},{""BNGR_NAME"":""BG_GECICI"",""POS"":""ORIGINAL""}]," & _
        """ACT_LOAD"":[{""LOAD_NAME"":""LG_ZATI"",""DAY"":""FIRST""},{""LOAD_NAME"":""LG_ONGERME"",""DAY"":""FIRST""}]") & "," & _
        """2"":" & Stage("CS2", DUR_CS2, """ACT_ELEM"":[{""GRUP_NAME"":""SG_TABLIYE"",""AGE"":" & JN(AGE_SLAB) & "}],""DACT_BNGR"":[""BG_GECICI""],""ACT_BNGR"":[{""BNGR_NAME"":""BG_SF"",""POS"":""ORIGINAL""}]") & "," & _
        """3"":" & Stage("CS3", DUR_CS3, """ACT_LOAD"":[{""LOAD_NAME"":""LG_DW"",""DAY"":""FIRST""}]") & "," & _
        """4"":" & Stage("CS4", DUR_CS4, "")
    Add "db/STAG", "{""Assign"":{" & s & "}}"
    Add "db/CSCS", "{""Assign"":{""1"":" & CscsJson(7) & ",""2"":" & CscsJson(8) & ",""3"":" & CscsJson(9) & ",""4"":" & CscsJson(10) & ",""5"":" & CscsJson(11) & ",""6"":" & CscsJson(12) & "}}"
    Add "db/TDNT", "{""Assign"":{""1"":{""NAME"":""HALAT"",""TYPE"":""INTERNAL"",""MATL"":4,""AREA"":" & JN(aStr) & ",""D_AREA"":" & JN(dStr) & _
        ",""RM"":0,""RV"":" & JN(RELAX_KL) & ",""US"":" & JN(fpuS * 1000#) & ",""YS"":" & JN(0.9 * fpuS * 1000#) & ",""LT"":""PRE""}}}"
    Add "db/TDGR", "{""Assign"":{""1"":{""NAME"":""TG_KIRIS""}}}"
    BuildTendons
    Add "db/TDNA", "{""Assign"":{" & tdnaJson & "}}"
    Add "db/TDCS", "{""Assign"":{" & tdcsJson & "}}"
    Add "db/TDPL", "{""Assign"":{" & tdplJson & "}}"
    Add "db/STCT", "{""Assign"":{""1"":{""bLAST_FINAL"":true,""FINAL_STAGE"":""CS4"",""iINC_NLA"":0,""iNLA_TYPE"":1,""bINC_PDL"":false,""iITER"":30,""TOL"":0.01," & _
        """bINC_TDE"":true,""bCNS"":true,""TYPE"":""BOTH"",""iITER_CR"":5,""TOL_CR"":0.01,""bOUCC"":false,""bITS"":false,""iITS"":2,""bATS"":true," & _
        """iT10"":2,""iT100"":5,""iT1K"":7,""iT5K"":10,""iT10K"":20,""bTTLE_CS"":true,""bRCE"":false,""bVAR"":true,""bAPPLY_ELA"":false,""bTTLE_ES"":true,""iTTLE_ES"":0," & _
        """vEREC"":[{""LTYPECC"":""CS_DW_1"",""EREC"":""DC"",""vLCNAME"":[""DW_1""]},{""LTYPECC"":""CS_DW_2"",""EREC"":""DW"",""vLCNAME"":[""DW_2""]}]," & _
        """CPFC"":""EXTERNAL"",""bCONV"":true,""bTRUSS"":true,""bBEAM"":true,""bCHANGE_CABLE"":true,""bAPPLY_IMF"":true,""bCAMBER"":false,""bSD"":false,""iBSC"":0}}}"
End Sub

Private Function Stage(ByVal nm As String, ByVal dur As Double, ByVal acts As String) As String
    Stage = "{""NAME"":" & QS(nm) & ",""DURATION"":" & JN(dur) & ",""bSV_RSLT"":true,""bSV_STEP"":false,""bLOAD_STEP"":false,""ADD_STEP"":[]" & IIf(acts = "", "", "," & acts) & "}"
End Function

' composite section for the stages: part 1 = girder (element material, CS1), part 2 = slab (material 2, CS2)
Private Function CscsJson(ByVal sec As Long) As String
    Dim hG As Double, vsP As Double
    If sec = 8 Or sec = 10 Or sec = 12 Then
        hG = 2 * polyAreaY / polyPerY: vsP = polyAreaY / polyPerY
    Else
        hG = NotionalG(): vsP = VsG()
    End If
    ' PARTINFO_H = notional size 2Ac/u (CEB codes), PARTINFO_VS = volume/surface (AASHTO / ACI codes)
    CscsJson = "{""SEC"":" & sec & ",""ASTAGE"":""CS1"",""TYPE"":""NORMAL"",""bTAP"":false,""vPARTINFO"":[" & _
        "{""PART"":1,""MTYPE"":""ELEM"",""MAT"":"""",""CSTAGE"":"""",""AGE"":" & JN(ageG) & ",""PARTINFO_H"":" & JN(hG) & PartRest(vsP) & "}," & _
        "{""PART"":2,""MTYPE"":""MATL"",""MAT"":""2"",""CSTAGE"":""CS2"",""AGE"":" & JN(AGE_SLAB) & ",""PARTINFO_H"":" & JN(tc) & PartRest(tc / 2) & "}]}"
End Function

Private Function PartRest(ByVal vs As Double) As String
    PartRest = ",""PARTINFO_VS"":" & JN(vs) & ",""PARTINFO_M"":0,""AREA"":1,""ASY"":1,""ASZ"":1,""IXX"":1,""IYY"":1,""IZZ"":1,""WAREA"":1,""IW"":1"
End Function

' volume / surface of the span girder, Ac / u (m) - total precast surface (AASHTO 5.4.2.3.2)
Private Function VsG() As Double
    VsG = polyArea / polyPer
End Function

' notional size of the span girder, h = 2 Ac / u (m)
Private Function NotionalG() As Double
    NotionalG = 2 * polyArea / polyPer
End Function

' AASHTO creep/shrinkage record (MIDAS TDMT CODE AASHTO: STR f'c, HU, AGE, VOL = V/S in m, bEXPOSE =
' 'exposed to drying before 5 days of curing' (+20 % shrinkage, AASHTO 5.4.2.3.3) - false: cured concrete)
Private Function TdmtJson(ByVal nm As String, ByVal fck As Double, ByVal age As Double, ByVal vs As Double) As String
    TdmtJson = "{""NAME"":" & QS(nm) & ",""CODE"":""AASHTO"",""STR"":" & JN(fck * 1000#) & ",""HU"":" & JN(humid) & ",""AGE"":" & JN(age) & _
        ",""VOL"":" & JN(vs) & ",""bEXPOSE"":false}"
End Function

' transfer length: AUTO2 = the drop-down 'Auto' entry the owner chose (2026-11-03)
' one straight tendon per girder, strand row and debond length: full girder length, debonded
' DeBondBLEN / DeBondELEN from both ends; length = the KILIFLAMA station where the strands count as bonded
Private Sub BuildTendons()
    Dim j As Long, r As Long, d As Long, cnt As Long, L As Double, el As String, i As Long
    tdnaJson = "": tdcsJson = "": tdplJson = "": nTdn = 0: nTdnGroup = 0
    For j = 0 To nG - 1
        el = ""
        For i = 1 To nSt - 1: el = el & IIf(el = "", "", ",") & gEl(j, i): Next i
        For r = 1 To nRow
            For d = 1 To nDb
                If d = 1 Then
                    cnt = rowN(r) - dbN(r, 1): L = 0
                Else
                    cnt = dbN(r, d - 1) - dbN(r, d): L = dbX(d)
                End If
                If cnt > 0 Then AddTendon "K" & (j + 1) & "-S" & r & "-" & Format(L * 100, "0"), el, gEl(j, 1), cnt, L, rowZ(r), IIf(j = 0 And ovL > 0, 3, IIf(j = nG - 1 And ovR > 0, 5, 1))
            Next d
        Next r
    Next j
End Sub

Private Sub AddTendon(ByVal nm As String, ByVal el As String, ByVal e1 As Long, ByVal cnt As Long, ByVal L As Double, ByVal z As Double, ByVal cscsNo As Long)
    Dim k As String, zr As Double
    ' z from the girder node (top of the slab, section offset CT), bottom option OFF: MIDAS drew
    ' bBOTZ = true + soffit height above the deck; checked live 2026-11-03 (owner: 6 cm above soffit OK)
    zr = z - (Hg + tc)
    nTdn = nTdn + 1: nTdnGroup = nTdnGroup + cnt: k = QS(CStr(nTdn))
    tdnaJson = tdnaJson & IIf(nTdn = 1, "", ",") & k & ":{""NAME"":" & QS(nm) & ",""TDN_PROP"":1,""ELEM"":[" & el & "],""BELENG"":0,""ELENG"":0," & _
        """CURVE"":""SPLINE"",""INPUT"":""2D"",""TDN_GRUP"":1,""LENG_OPT"":""AUTO2"",""bTP"":true,""CNT"":" & cnt & _
        ",""DeBondBLEN"":" & JN(L) & ",""DeBondELEN"":" & JN(L) & ",""SHAPE"":""ELEMENT"",""INS_PT"":""END-I"",""INS_ELEM"":" & e1 & _
        ",""AXIS_IJ"":""I-J"",""XAR_ANGLE"":0,""bPJ"":true,""OFF_YZ"":[0,0]," & _
        """PROFY"":[{""PT"":[0,0],""bFIX"":true,""R"":0},{""PT"":[" & JN(Lg) & ",0],""bFIX"":true,""R"":0}]," & _
        """PROFZ"":[{""PT"":[0," & JN(zr) & "],""bFIX"":true,""R"":0,""bBOTZ"":false},{""PT"":[" & JN(Lg) & "," & JN(zr) & "],""bFIX"":true,""R"":0,""bBOTZ"":false}]}"
    tdcsJson = tdcsJson & IIf(nTdn = 1, "", ",") & k & ":{""TDNA"":" & nTdn & ",""CSCS"":" & cscsNo & ",""PART_NUM"":1}"
    tdplJson = tdplJson & IIf(nTdn = 1, "", ",") & k & ":{""ITEMS"":[{""ID"":1,""LCNAME"":""Prestress"",""GROUP_NAME"":""LG_ONGERME"",""TENDON_NAME"":" & QS(nm) & _
        ",""TYPE"":""STRESS"",""ORDER"":""BOTH"",""BEGIN"":" & JN(fpi * 1000#) & ",""END"":" & JN(fpi * 1000#) & ",""GROUTING"":0}]}"
End Sub

Private Function JoinKeys(d As Object) As String
    Dim k As Variant, s As String
    For Each k In SortedKeysL(d): s = s & IIf(s = "", "", ",") & k: Next k
    JoinKeys = s
End Function

' outer end beam depth: deck top to the bearing bottom (girder H + slab tc + bearing / link height)
Private Function DisHgt() As Double
    DisHgt = Hg + tc + linkLen
End Function

' db/CO_S (Section Color): W = wireframe (2D), HF = hidden fill (3D surface), HE = hidden edge (3D outline)
Private Function SectColorJson() As String
    Dim rows() As String, f() As String, c() As String, i As Long, s As String, k As Long, v(1 To 3) As Long, t(1 To 3) As Long
    rows = Split(SECT_COLORS, ";")
    For i = LBound(rows) To UBound(rows)
        f = Split(rows(i), "|"): c = Split(f(1), ",")
        If f(0) <> "14" Or visW > 0 Then
            For k = 1 To 3
                v(k) = CLng(c(k - 1))
                t(k) = CLng(v(k) + (255 - v(k)) * SECT_FILL_TINT): If t(k) > 255 Then t(k) = 255
            Next k
            s = s & IIf(s = "", "", ",") & QS(f(0)) & ":{""W_R"":" & v(1) & ",""W_G"":" & v(2) & ",""W_B"":" & v(3) & _
                ",""HF_R"":" & t(1) & ",""HF_G"":" & t(2) & ",""HF_B"":" & t(3) & _
                ",""HE_R"":" & v(1) & ",""HE_G"":" & v(2) & ",""HE_B"":" & v(3) & ",""bBLEMD"":false,""FACT"":0.5}"
        End If
    Next i
    SectColorJson = "{""Assign"":{" & s & "}}"
End Function

Private Function SbSect(ByVal no As Long, ByVal nm As String, ByVal h As Double, ByVal b As Double) As String
    SbSect = QS(CStr(no)) & ":{""SECTTYPE"":""DBUSER"",""SECT_NAME"":" & QS(nm) & ",""SECT_BEFORE"":{" & SectHead() & ",""SHAPE"":""SB"",""DATATYPE"":2,""SECT_I"":{""vSIZE"":[" & _
        JN(h) & "," & JN(b) & ",0,0,0,0,0,0,0,0]}}}"
End Function

' DB user T: vSIZE [H, B (flange), tw (web), tf (flange)], top of the flange at the node (offset CT)
' effective diaphragm slab flange (Hambly 4.4.3), 0 = no flange (rectangle)
Private Function DiaFlange() As Double
    Dim fl As Double: fl = diaW: If fl > 0.3 * sG Then fl = 0.3 * sG
    If fl > enlB + 0.001 Then DiaFlange = fl
End Function

Private Function TSect(ByVal no As Long, ByVal nm As String, ByVal h As Double, ByVal b As Double, ByVal tw As Double, ByVal tf As Double) As String
    TSect = QS(CStr(no)) & ":{""SECTTYPE"":""DBUSER"",""SECT_NAME"":" & QS(nm) & ",""SECT_BEFORE"":{" & SectHead() & ",""SHAPE"":""T"",""DATATYPE"":2,""SECT_I"":{""vSIZE"":[" & _
        JN(h) & "," & JN(b) & "," & JN(tw) & "," & JN(tf) & ",0,0,0,0,0,0]}}}"
End Function

' Hambly: torsion of a slab strip of width b and depth t is b t^3 / 6 (half of the solid-bar value);
' factor on the rectangle torsion constant that MIDAS computes (RectJ, Analysis Reference Eq. 2)
' Section factors of the visual strip (section 14, drawn visW x tc): every property back to the true strip dxAct x tc.
' Rectangle B x t: A, Asy, Asz, Iyy (vertical bending, b t^3/12) ~ B; Izz (in plane) ~ B^3; torsion = Hambly dx t^3/6
' over the rectangle J of the drawn section (same formula as HamblyJ).
Private Function VisSecf() As String
    Dim r As Double, jx As Double, a As Double, c As Double
    r = dxAct / visW
    a = IIf(visW >= tc, visW, tc): c = IIf(visW >= tc, tc, visW)
    jx = (dxAct * tc ^ 3 / 6#) / RectJ(a, c)
    VisSecf = SecfJson(14, r, jx, r, r ^ 3, r, 1)
End Function

' Hambly torsion factor on a slab member section (only Ixx changes)
Private Function SlabSecf(ByVal sec As Long) As String
    SlabSecf = SecfJson(sec, 1, HamblyJ(SlabW(sec), tc), 1, 1, 1, 1)
End Function

Private Function SecfJson(ByVal sec As Long, ByVal aF As Double, ByVal jF As Double, ByVal iyF As Double, ByVal izF As Double, ByVal wF As Double, ByVal part As Long) As String
    SecfJson = QS(CStr(sec)) & ":{""ITEMS"":[{""ID"":1,""GROUP_NAME"":" & QS(IIf(dynMode, "", "BG_SF")) & ",""AREA_SF"":" & JN(aF) & ",""ASY_SF"":" & JN(aF) & ",""ASZ_SF"":" & JN(aF) & _
        ",""IXX_SF"":" & JN(jF) & ",""IYY_SF"":" & JN(iyF) & ",""IZZ_SF"":" & JN(izF) & ",""WGT_SF"":" & JN(wF) & ",""W_SF"":1,""IPART"":" & part & ",""bDiffIJ"":false}]}"
End Function

' width of a slab member section
Private Function SlabW(ByVal sec As Long) As Double
    Select Case sec
        Case 3: SlabW = dxAct
        Case 5: SlabW = endStripW
        Case 13: SlabW = diaW
    End Select
End Function

Private Function HamblyJ(ByVal b As Double, ByVal t As Double) As Double
    Dim a As Double, c As Double
    a = IIf(b >= t, b, t): c = IIf(b >= t, t, b)
    HamblyJ = (b * t ^ 3 / 6#) / RectJ(a, c)
End Function

' torsion constant of a solid rectangle as MIDAS computes it (Civil NX Analysis Reference, "Torsional
' Resistance (Ixx)" Eq. 2: ixx = a b^3 [16/3 - 3.36 (b/a)(1 - b^4/(12 a^4))] with 2a, 2b = sides);
' written with the full sides a >= c: a c^3/3 (1 - 0.63 (c/a)(1 - (c/a)^4/12))
Private Function RectJ(ByVal a As Double, ByVal c As Double) As Double
    Dim r As Double: r = c / a
    RectJ = a * c ^ 3 / 3# * (1# - 0.63 * r * (1# - r ^ 4 / 12#))
End Function

Private Function SectHead() As String
    SectHead = """OFFSET_PT"":""CT"",""OFFSET_CENTER"":0,""USER_OFFSET_REF"":0,""HORZ_OFFSET_OPT"":0,""USERDEF_OFFSET_YI"":0,""VERT_OFFSET_OPT"":0,""USERDEF_OFFSET_ZI"":0,""USE_SHEAR_DEFORM"":true,""USE_WARPING_EFFECT"":false"
End Function

' 6 = PSC value (span girder, not used by elements), 7 = composite span girder, 8 = composite end-zone girder,
' 9/10 = left edge girder (span / end zone), 11/12 = right edge girder
Private Function PscSect(ByVal no As Long) As String
    Dim core As String, chk As String, k() As Double, pj As String, i As Long, nm As String
    ReDim k(1 To 12)
    Dim isY As Boolean, bc As Double
    isY = (no = 8 Or no = 10 Or no = 12)
    For i = 1 To 12: k(i) = IIf(isY, kbY(i), kb(i)): Next i
    pj = IIf(isY, polyJsonY, polyJson)
    ' slab flange: inner girders = spacing; edge girders = half spacing + overhang (Hambly)
    bc = sG: If no = 9 Or no = 10 Then bc = sG / 2 + ovL
    If no = 11 Or no = 12 Then bc = sG / 2 + ovR
    nm = KirisAdi() & IIf(isY, "-Y_KOMP", "_KOMP") & IIf(no = 9 Or no = 10, "_SOL", IIf(no = 11 Or no = 12, "_SAG", ""))
    core = """SECT_I"":{""vSIZE"":[" & JN(Hg) & "," & JN(k(3)) & "," & JN(k(7) + k(8) + k(9)) & "," & JN(k(11) + k(12)) & "]," & pj & "}"
    chk = """SHEAR_CHK"":true,""SHEAR_CHK_POS"":[[0,0,0],[0,0,0]],""USE_AUTO_QY"":[[true,true,true],[false,false,false]],""WEB_THICK"":[" & _
          IIf(no = 6, JN(kb(2)), "0") & ",0],""USE_WEB_THICK_SHEAR"":[[true,true,true],[false,false,false]]"
    If no = 6 Then
        PscSect = """6"":{""SECTTYPE"":""PSC"",""SECT_NAME"":" & QS(KirisAdi() & "") & ",""CALC_OPT"":true,""SECT_BEFORE"":{" & SectHead() & ",""SHAPE"":""VALU""," & core & "," & chk & "}}"
    Else
        PscSect = QS(CStr(no)) & ":{""SECTTYPE"":""COMPOSITE"",""SECT_NAME"":" & QS(nm) & ",""CALC_OPT"":true,""SECT_BEFORE"":{" & SectHead() & ",""SHAPE"":""PC""," & core & _
            ",""MATL_ELAST"":" & JN(eG / eS) & ",""MATL_DENS"":1,""MATL_POIS_S"":0.2,""MATL_POIS_C"":0.2,""MATL_THERMAL"":1," & chk & ",""USE_MULTI_ELAST"":false}," & _
            """SECT_AFTER"":{""SECT_I"":{""vSIZE"":[" & JN(bc) & ",0],""BUILT_FLAG"":1},""SECT_J"":{""vSIZE"":[" & JN(bc) & "," & JN(tc) & ",0]}}}"
    End If
End Function

' Composite girder sections 7-12 as MCT text (*SECT-PSCVALUE, Composite PC) with bHUMBLY = YES:
' the "Use Hambly Eq." option has no JSON key, but MCT import sets it (tested live 2026-11-03).
' Stiffness lines are zeros - MIDAS recalculates them from the polygon on import (tested).
Private Function SectMct() As String
    Dim s As String, no As Long, isY As Boolean, bc As Double, k() As Double, i As Long, nm As String, pts As String
    ReDim k(1 To 12)
    s = "*UNIT" & vbCrLf & "   KN   , M, KJ, C" & vbCrLf & "*SECT-PSCVALUE" & vbCrLf
    For no = 7 To 12
        isY = (no = 8 Or no = 10 Or no = 12)
        For i = 1 To 12: k(i) = IIf(isY, kbY(i), kb(i)): Next i
        bc = sG: If no = 9 Or no = 10 Then bc = sG / 2 + ovL
        If no = 11 Or no = 12 Then bc = sG / 2 + ovR
        nm = Pj(KirisAdi() & IIf(isY, "-Y_KOMP", "_KOMP") & IIf(no = 9 Or no = 10, "_SOL", IIf(no = 11 Or no = 12, "_SAG", "")))
        pts = IIf(isY, polyPtsY, polyPts)
        s = s & " SECT=" & no & ", COMPOSITE , " & nm & ", CT, 0, 0, 0, 0, 0, 0, YES, NO, YES, PC" & vbCrLf & _
            "       0, 0, 0, 0, 0, 0" & vbCrLf & _
            "       0, 0, 0, 0, 0, 0, 0, 0, 0, 0" & vbCrLf & _
            "       0, 0, 0, 0, 0, 0, 0, 0" & vbCrLf & _
            "       " & JN(Hg) & ", " & JN(k(3)) & ", " & JN(k(7) + k(8) + k(9)) & ", " & JN(k(11) + k(12)) & vbCrLf & _
            "       YES, 0, 0, YES, , YES, , YES, , 0, YES, , YES, , YES, " & vbCrLf & _
            "       " & JN(bc) & ", 1, 0,   " & JN(bc) & ", " & JN(tc) & ", 0,   " & JN(eG / eS) & ", 1, 0.2, 0.2, 1, NO, , " & vbCrLf & _
            "       OPOLY=" & pts & vbCrLf
    Next no
    SectMct = s & "*ENDDATA" & vbCrLf
End Function

Private Function KirisAdi() As String
    KirisAdi = Replace(CStr(ThisWorkbook.Names("G_KIRIS_TIPI").RefersToRange.Cells(spanNo).Value), " ", "")
End Function

Private Function StldJson() As String
    Dim a As Variant, i As Long, p As Variant, s As String
    ' owner 2026-11-03: water, earth, surcharge, snow and ice cases removed (no such loads on this superstructure);
    ' FR, PL, BR, WS-X, WS-Z stay as (still empty) placeholders
    a = Array("DC|D|Self Wieght", "DW_1|DC|Barrier + Pavement", "DW_2|DW|Wearing Surface", "Prestress|PS|", _
        "FR|FR|Friction Force", "WS-X|W|Horizontal Wind Load (Long.)", "WS-Y|W|Horizontal Wind Load (Trans.)", "WS-Z|W|Vertical Wind Load", _
        "WL|WL|Wind Load on Live Load", "TU+|T|Temperature Effect (+)", "TU-|T|Temperature Effect (-)", _
        "PL|L|Pedestrian Live Load", "BR|BRK|Breaking Load", "CF|CF|Centrifugal Force", "MOVE|L|Eq. Loading")
    For i = 0 To UBound(a)
        p = Split(a(i), "|")
        s = s & IIf(i = 0, "", ",") & QS(CStr(i + 1)) & ":{""NO"":" & (i + 1) & ",""NAME"":" & QS(CStr(p(0))) & ",""TYPE"":" & QS(CStr(p(1))) & ",""DESC"":" & QS(CStr(p(2))) & "}"
    Next i
    StldJson = s
End Function

Private Function LaneJson(ByVal nm As String, ByVal t As Long) As String
    Dim i As Long, s As String, j As Long, ecc As Double
    j = trG(t): ecc = j * sG - tr(t)
    For i = 1 To nSt - 1
        s = s & IIf(i = 1, "", ",") & "{""ELEM"":" & gEl(j, i) & ",""ECC"":" & JN(ecc) & ",""FACT"":0,""SPAN_START"":false,""ECCEN_VERT_LOAD"":0,""CENT_F"":0.5}"
    Next i
    LaneJson = """#"":{""COMMON"":{""LL_NAME"":" & QS(nm) & ",""LOAD_DIST"":""CROSS"",""GROUP_NAME"":""CROSS_BEAM"",""SKEW_START"":0,""SKEW_END"":0,""MOVING"":""BOTH""," & _
        """WHEEL_SPACE"":1.5,""WIDTH"":2,""OPT_AUTO_LANE"":false,""ALLOW_WIDTH"":3},""LANE_ITEMS"":[" & s & "]}"
End Function

Private Function RenumberTop(ByVal s As String) As String
    Dim n As Long, p As Long
    Do
        p = InStr(s, """#"":")
        If p = 0 Then Exit Do
        n = n + 1
        s = Left$(s, p - 1) & QS(CStr(n)) & ":" & Mid$(s, p + 4)
    Loop
    RenumberTop = s
End Function

Private Function MvCase(ByVal no As Long, ByVal nm As String, ByVal veh As String, ByVal twoLanes As Boolean) As String
    Dim lanes As String, mn As Long, mx As Long
    If twoLanes And nTr = 2 Then
        lanes = """LANE1"",""LANE2""": mn = 0: mx = 2
    ElseIf twoLanes Then
        lanes = """LANE2""": mn = 0: mx = 1
    Else
        lanes = """LANE2""": mn = 1: mx = 1
    End If
    MvCase = QS(CStr(no)) & ":{""LCNAME"":" & QS(nm) & ",""OPT_AUTO_OPTIMIZE"":false,""TYPE_LOADMODEL"":5,""DESC"":"""",""OPT_COMB"":1,""SCALE_FACTOR1"":0.8,""SCALE_FACTOR2"":0.7,""SCALE_FACTOR3"":0.6," & _
        """OPT_PSI_FACTOR"":true,""MULTI_FACTOR1"":1,""MULTI_FACTOR2"":1,""MULTI_FACTOR3"":0.75,""SUB_LOAD_LIST"":[{""TYPE"":2,""NAME"":" & QS(veh) & ",""SCALE_FACTOR"":1," & _
        """MIN_LOAD_LANE_TYPE"":" & mn & ",""MAX_LOAD_LANE_TYPE"":" & mx & ",""SLN_LIST"":[" & lanes & "]}]}"
End Function

Private Function Phi3(ByVal L As Double) As Double
    Dim p As Double
    p = 2.16 / (Sqr(L) - 0.2) + 0.73
    If p < 1 Then p = 1
    If p > 2 Then p = 2
    Phi3 = p
End Function

' ---- combinations ----
Private Sub AddCombs(ByVal ep As String, ByVal conc As Boolean)
    Dim t1 As String, t2 As String, no As Long, k As Long, sv As Variant, mins As Variant, maxs As Variant, cbs As Variant
    mins = Array("STR-1-Min.", "STR-2-Min.", "STR-3-Min", "STR-4-Min.", "STR-5-Min")
    maxs = Array("STR-1-Maks.", "STR-2-Maks.", "STR-3-Maks.", "STR-4-Maks.", "STR-5-Maks")
    cbs = Array(True, True, False, False, True)
    ' owner 2026-11-03: Servis 4 removed (AASHTO Service IV = prestressed substructure tension), Servis 2 removed
    ' (Service II = steel yielding / slip-critical connections), STR-2 removed (= STR-1 with a smaller LL factor,
    ' same LM71 loading, always enveloped by STR-1)
    For Each sv In Array(1, 3)
        k = sv
        no = no + 1: t1 = t1 & IIf(t1 = "", "", ",") & Comb(no, "Servis " & k, "ACTIVE", True, 0, TableItems("Servis-" & k, 0, k), conc)
    Next sv
    For k = 1 To 5
        If k = 2 Then GoTo NextStr
        no = no + 1: t1 = t1 & "," & Comb(no, CStr(mins(k - 1)), "ACTIVE", CBool(cbs(k - 1)), 0, TableItems("STR-" & k, 2, k), conc)
        no = no + 1: t1 = t1 & "," & Comb(no, CStr(maxs(k - 1)), "ACTIVE", CBool(cbs(k - 1)), 0, TableItems("STR-" & k, 1, k), conc)
NextStr:
    Next k
    ' ICE-1 / SNOW-1 / DEAD removed (owner 2026-11-03)
    Dim env As String, a As String, i As Long
    a = "CB|Servis 1|1;CB|Servis 3|1"
    no = no + 1: t2 = Comb(no, "ENV_SER", "ACTIVE", True, 1, a, conc)
    a = ""
    For k = 0 To 4
        If k <> 1 Then a = a & IIf(a = "", "", ";") & "CB|" & mins(k) & "|1;CB|" & maxs(k) & "|1"
    Next k
    no = no + 1: t2 = t2 & "," & Comb(no, "ENV_STR", "ACTIVE", True, 1, a, conc)
    Add ep, "{""Assign"":{" & t1 & "}}"
    Add ep, "{""Assign"":{" & t2 & "}}"
End Sub

' mode 0 = service (table values), 1 = max, 2 = min; strNo for STR-4 DC max
Private Function TableItems(ByVal lab As String, ByVal mode As Long, ByVal strNo As Long) As String
    Dim f As Variant, fDC As Double, fDW As Double, s As String, v As Variant
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
    ' FR takes the Tablo 14 WA column (as in the reference model); TU- carries load but is not combined
    ' (owner 2026-11-03: only TU+ in the combinations, to keep their number down)
    s = CsItems(fDC, fDW, mode = 0) & It("ST", "FR", f(2)) & _
        It("ST", "WS-X", f(3)) & It("ST", "WS-Y", f(3)) & It("ST", "WS-Z", f(3)) & It("ST", "WL", f(4)) & _
        It("ST", "TU+", f(5)) & It("ST", "PL", f(1)) & It("ST", "BR", f(1)) & It("ST", "CF", f(1)) & _
        It("ST", "MOVE", f(1)) & It("MV", "LM71", f(1))
    TableItems = Mid$(s, 2)
End Function

' dead loads + prestress from the construction-stage results (owner 2026-11-02, AASHTO split):
' service = MIDAS 'CS: Summation' forces = Dead + Erection + Tendon Primary + Tendon / Creep / Shrinkage
'   Secondary (Civil NX help 'Construction Stage Analysis Control': Creep / Shrinkage PRIMARY are the
'   imaginary forces that cause the strain, not member forces - not in the summation; owner 2026-11-03);
' strength = dead x gamma + SECONDARY tendon / creep / shrinkage effects x 1.0 (AASHTO Table 3.4.1-3)
Private Function CsItems(ByVal fDC As Double, ByVal fDW As Double, ByVal service As Boolean) As String
    Dim s As String
    ' dynamic model: no stage results, the static cases carry the dead loads; no prestress in this model
    If dynMode Then CsItems = It("ST", "DC", fDC) & It("ST", "DW_1", fDC) & It("ST", "DW_2", fDW): Exit Function
    s = It("CS", "Dead Load", fDC) & It("CS", "CS_DW_1", fDC) & It("CS", "CS_DW_2", fDW)
    If service Then s = s & It("CS", "Tendon Primary", 1)
    CsItems = s & It("CS", "Tendon Secondary", 1) & It("CS", "Creep Secondary", 1) & It("CS", "Shrinkage Secondary", 1)
End Function

Private Function It(ByVal an As String, ByVal lc As String, ByVal f As Double) As String
    If f < 0 Then Exit Function
    It = ";" & an & "|" & lc & "|" & JN(f)
End Function

Private Function Comb(ByVal no As Long, ByVal nm As String, ByVal act As String, ByVal bCB As Boolean, ByVal iType As Long, ByVal items As String, ByVal conc As Boolean) As String
    Dim p As Variant, i As Long, s As String, q2 As Variant, an As String
    p = Split(items, ";")
    For i = 0 To UBound(p)
        q2 = Split(p(i), "|")
        an = q2(0): If conc And an = "CB" Then an = "CBC"
        s = s & IIf(i = 0, "", ",") & "{""ANAL"":" & QS(an) & ",""LCNAME"":" & QS(CStr(q2(1))) & ",""FACTOR"":" & q2(2) & "}"
    Next i
    Comb = QS(CStr(no)) & ":{""NO"":" & no & ",""NAME"":" & QS(nm) & ",""ACTIVE"":" & QS(IIf(conc, "STRENGTH", act)) & _
        IIf(conc, ",""bES"":false", "") & ",""bCB"":" & IIf(bCB, "true", "false") & ",""iTYPE"":" & iType & ",""DESC"":"""",""vCOMB"":[" & s & "]}"
End Function

' girder / slab concrete. Dynamic: EN04(RC) database (Ecm, 25 kN/m3). Staged: user concrete with the AASHTO Ec
' and the workbook unit weight; Poisson 0.2 and thermal 1e-5 as in the database.
Private Function ConcMat(ByVal no As Long, ByVal cls As String, ByVal eMpa As Double, ByVal den As Double) As String
    Dim prm As String
    If dynMode Then
        prm = "{""P_TYPE"":1,""STANDARD"":""EN04(RC)"",""CODE"":"""",""DB"":" & QS(EnDb(cls)) & "}"
    Else
        prm = "{""P_TYPE"":2,""ELAST"":" & JN(eMpa * 1000#) & ",""POISN"":0.2,""THERMAL"":0.00001,""DEN"":" & JN(den) & "}"
    End If
    ConcMat = QS(CStr(no)) & ":{""TYPE"":""CONC"",""NAME"":" & QS(cls) & ",""HE_SPEC"":0,""HE_COND"":0,""bMASS_DENS"":false,""DAMP_RAT"":0.05,""PARAM"":[" & prm & "]}"
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

' ============================================================================
'  SEND / WRITE
' ============================================================================
Private Sub Add(ByVal ep As String, ByVal body As String)
    payloads.Add Array(ep, body)
End Sub

Private Sub WritePayloads(ByVal dirPath As String)
    Dim i As Long, p As Variant
    MakeDir dirPath
    For i = 1 To payloads.Count
        p = payloads(i)
        WriteText dirPath & "\" & Format(i, "00") & "_" & Replace(Replace(p(0), "db/", ""), "/", "_") & ".json", Replace(CStr(p(1)), DYN_TAG, "")
    Next i
    WriteText dirPath & "\00_SECT_composite_7-12.mct", SectMct()
    WriteText dirPath & "\summary.txt", Summary()
End Sub

Private Function SendAll(Optional ByVal saveModel As Boolean = True) As String
    Dim rep As String, i As Long, p As Variant, resp As String, ok As Boolean, path As String, area As Double, t0 As Single
    t0 = Timer
    If Not GetKey() Then SendAll = "FAIL: no MAPI key.": Exit Function
    Set http = Nothing
    ' writes into the model open in Civil NX (no doc/NEW) and saves that file (doc/SAVE).
    If Not Api("GET", "db/NODE", "", resp) Then SendAll = "FAIL: Civil NX does not answer (db/NODE): " & Left$(resp, 300): Exit Function
    ' two models per bridge (owner 2026-11-05): do not wipe the other model by mistake
    If Not ModelTypeOk(rep) Then SendAll = rep: Exit Function
    rep = ""
    ' Clean pass first (like the culvert build): a PUT into a filled model ADDS beam-load
    ' items and keeps old records, so every table this macro writes is deleted and then
    ' read back empty. Anything left -> stop before the first write.
    Dim cleanRep As String
    ProgReset 2 * (UBound(CleanTables()) + 1) + 1 + payloads.Count + 4 + IIf(dynMode, 3, 0)
    cleanRep = CleanModel()
    If Left$(cleanRep, 7) = "STOPPED" Then
        SendAll = cleanRep & vbCrLf & "Nothing was written."
        Exit Function
    End If
    ' composite girder sections 7-12 by MCT import (Hambly torsion option; no JSON key). NOTE: doc/IMPORTMXT
    ' REPLACES the whole model content (tested 2026-11-03), so it runs here, after the clean pass and
    ' before every JSON table; the open file name is kept. The JSON SECT table sends the other sections.
    Dim mctPath As String
    mctPath = Environ$("TEMP") & "\midas_ustyapi_sect.mct"
    WriteText mctPath, SectMct()
    Prog "composite sections (MCT import)"
    If Not Api("POST", "doc/IMPORTMXT", "{""Argument"":" & QS(mctPath) & "}", resp) Then
        SendAll = cleanRep & vbCrLf & "FAIL: MCT import of the composite sections (" & mctPath & "): " & Left$(resp, 300) & vbCrLf & "Nothing else was written."
        Exit Function
    End If
    If Not Api("GET", "db/SECT", "", resp) Then resp = ""
    If CountOf(resp, """SECTTYPE"":""COMPOSITE""") <> 6 Then
        SendAll = cleanRep & vbCrLf & "FAIL: after the MCT import Civil NX holds " & CountOf(resp, """SECTTYPE"":""COMPOSITE""") & " composite sections (6 expected). Nothing else was written."
        Exit Function
    End If
    cleanRep = cleanRep & vbCrLf & "Composite sections 7-12 imported by MCT (Hambly torsion option on)."
    If dynMode Then
        Dim jRep As String
        If Not MeasureGirderJ(mctPath, jRep) Then
            SendAll = cleanRep & vbCrLf & "FAIL: girder J measurement: " & jRep & vbCrLf & "The open file now holds test cantilevers (doc/ANAL writes the open file): rebuild it."
            Exit Function
        End If
        cleanRep = cleanRep & vbCrLf & jRep
    End If
    For i = 1 To payloads.Count
        p = payloads(i)
        Prog "table " & i & "/" & payloads.Count & "  " & p(0)
        If Not Api("PUT", CStr(p(0)), Replace(CStr(p(1)), DYN_TAG, GirderSecf()), resp) Then
            SendAll = cleanRep & vbCrLf & "FAIL at step " & i & " (" & p(0) & "):" & vbCrLf & Left$(resp, 600) & vbCrLf & vbCrLf & "Model is incomplete; nothing saved."
            Exit Function
        End If
    Next i
    ' read back
    rep = "Read-back:" & vbCrLf
    Prog "read-back nodes"
    If Api("GET", "db/NODE", "", resp) Then rep = rep & "  nodes " & CountOf(resp, """X"":") & " / " & nCnt & vbCrLf
    Prog "read-back elements"
    If Api("GET", "db/ELEM", "", resp) Then rep = rep & "  elements " & CountOf(resp, """TYPE"":") & " / " & eCnt & vbCrLf
    Prog "read-back sections"
    If Api("GET", "db/SECT", "", resp) Then
        area = SectArea(resp, "7")
        rep = rep & "  girder area (section 7) MIDAS " & Format(area, "0.000000") & " / polygon " & Format(polyArea, "0.000000") & _
              IIf(Abs(area - polyArea) < 0.0001, "  OK", "  FAIL") & vbCrLf
        ' (DB user sections 2-5/13: GET db/SECT returns no computed properties, so no area / J check here)
        area = SectArea(resp, "8")
        rep = rep & "  end-zone area (section 8) MIDAS " & Format(area, "0.000000") & " / polygon " & Format(polyAreaY, "0.000000") & _
              IIf(Abs(area - polyAreaY) < 0.0001, "  OK", "  FAIL") & vbCrLf
    End If
    rep = rep & SecfCheck()
    Prog "save"
    If InStr(rep, "FAIL") > 0 Then
        rep = rep & "NOT SAVED: the read-back failed." & vbCrLf
    ElseIf Not saveModel Then
        rep = rep & "Not saved (saveModel = False)." & IIf(dynMode, " Note: the J analysis already wrote the test cantilevers to the open file.", "") & vbCrLf
    ElseIf Api("POST", "doc/SAVE", "{""Argument"":{}}", resp) Then
        rep = rep & "Saved (open model file)." & vbCrLf
    Else
        rep = rep & "FAIL save: " & Left$(resp, 300) & vbCrLf
    End If
    If warnText <> "" Then rep = rep & "WARNINGS:" & vbCrLf & warnText
    SendAll = cleanRep & vbCrLf & "All " & payloads.Count & " tables sent (" & Format(Timer - t0, "0") & " s)." & vbCrLf & rep & vbCrLf & SCRIPT_ID & " " & SCRIPT_VERSION & ", " & bridgeName & ", span " & spanNo & " = " & Lc & " m"
End Function

' The project comment of a model built by this macro says which type it is (ends DYNAMIC or not).
Private Function ModelTypeOk(ByRef msg As String) As Boolean
    Dim pj As String, wasDyn As Boolean, txt As String
    ModelTypeOk = True
    If Not Api("GET", "db/PJCF", "", pj) Then Exit Function
    If InStr(pj, SCRIPT_ID) = 0 Then Exit Function
    wasDyn = (InStr(pj, " DYNAMIC") > 0)
    If wasDyn = dynMode Then Exit Function
    txt = "The model open in Civil NX is a " & IIf(wasDyn, "DYNAMIC", "STAGED") & " model, but you are building the " & IIf(dynMode, "DYNAMIC", "STAGED") & " model."
    If askUser Then
        If UiMsg(txt & vbCrLf & vbCrLf & "WIPE it anyway?", vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) = vbYes Then Exit Function
    End If
    msg = "STOPPED: " & txt & " Nothing was changed. Open the right .mcb, or use an empty file."
    ModelTypeOk = False
End Function

' ---- girder J of the dynamic model ----
' Without construction stages MIDAS ignores the Hambly option of the composite sections and uses the J of
' the merged section, which the API does not return (measured 2026-11-05). So: six cantilevers, one per
' section, tip torque, J = T L / (G theta). Target = bare girder J (db/SECT RXX) + b t^3 / 6 / (Eg / Es)
' (Section Properties - Composite page; same Poisson's ratio). The MCT import then wipes the cantilevers.
Private Function MeasureGirderJ(ByVal mctPath As String, ByRef msg As String) As Boolean
    Const TL As Double = 10#, TE As Double = 30000000#, TN As Double = 0.2, TQ As Double = 100#
    Dim no As Long, a As Long, sN As String, sE As String, sC As String, sL As String, keys As String
    Dim resp As String, bare(7 To 12) As Double, rows As Variant, f As Variant, r As Long, sec As Long, th As Double, n As Long
    Erase jMer: Erase jHam: Erase jFac
    Prog "girder J: test cantilevers"
    If Not Api("GET", "db/SECT", "", resp) Then msg = "db/SECT: " & Left$(resp, 200): Exit Function
    For no = 7 To 12
        bare(no) = SectField(resp, CStr(no), "RXX")
        If bare(no) <= 0 Then msg = "no bare girder J (RXX) for section " & no: Exit Function
        a = 2 * (no - 6) - 1
        sN = sN & IIf(sN = "", "", ",") & QS(CStr(a)) & ":{""X"":0,""Y"":" & 5 * (no - 6) & ",""Z"":0}," & QS(CStr(a + 1)) & ":{""X"":" & JN(TL) & ",""Y"":" & 5 * (no - 6) & ",""Z"":0}"
        sE = sE & IIf(sE = "", "", ",") & QS(CStr(no - 6)) & ":{""TYPE"":""BEAM"",""MATL"":1,""SECT"":" & no & ",""NODE"":[" & a & "," & a + 1 & ",0,0,0,0,0,0],""ANGLE"":0,""STYPE"":0}"
        sC = sC & IIf(sC = "", "", ",") & QS(CStr(a)) & ":{""ITEMS"":[{""ID"":1,""GROUP_NAME"":"""",""CONSTRAINT"":""1111111""}]}"
        sL = sL & IIf(sL = "", "", ",") & QS(CStr(a + 1)) & ":{""ITEMS"":[{""ID"":1,""LCNAME"":""JT"",""GROUP_NAME"":"""",""FX"":0,""FY"":0,""FZ"":0,""MX"":" & JN(TQ) & ",""MY"":0,""MZ"":0}]}"
        keys = keys & IIf(keys = "", "", ",") & a + 1
    Next no
    If Not Put1("db/MATL", "{""Assign"":{""1"":{""TYPE"":""USER"",""NAME"":""JTEST"",""HE_SPEC"":0,""HE_COND"":0,""bMASS_DENS"":false,""DAMP_RAT"":0.05,""PARAM"":[{""P_TYPE"":2,""ELAST"":" & JN(TE) & ",""POISN"":" & JN(TN) & ",""THERMAL"":0.00001,""DEN"":25,""MASS"":0}]}}}", msg) Then Exit Function
    If Not Put1("db/NODE", "{""Assign"":{" & sN & "}}", msg) Then Exit Function
    If Not Put1("db/ELEM", "{""Assign"":{" & sE & "}}", msg) Then Exit Function
    If Not Put1("db/CONS", "{""Assign"":{" & sC & "}}", msg) Then Exit Function
    If Not Put1("db/STLD", "{""Assign"":{""1"":{""NAME"":""JT"",""TYPE"":""D"",""DESC"":""""}}}", msg) Then Exit Function
    If Not Put1("db/CNLD", "{""Assign"":{" & sL & "}}", msg) Then Exit Function
    Prog "girder J: analysis"
    If Not Api("POST", "doc/ANAL", "{}", resp) Then msg = "doc/ANAL: " & Left$(resp, 300): Exit Function
    If Not Api("POST", "post/TABLE", "{""Argument"":{""TABLE_NAME"":""D"",""TABLE_TYPE"":""DISPLACEMENTG"",""UNIT"":{""FORCE"":""KN"",""DIST"":""M""},""STYLES"":{""FORMAT"":""Scientific"",""PLACE"":8},""NODE_ELEMS"":{""KEYS"":[" & keys & "]},""LOAD_CASE_NAMES"":[""JT(ST)""]}}", resp) Then
        msg = "post/TABLE: " & Left$(resp, 300): Exit Function
    End If
    ' rows: ["1","2","JT","DX","DY","DZ","RX","RY","RZ"]
    rows = Split(resp, "[""")
    For r = 1 To UBound(rows)
        f = Split(Replace(Replace(rows(r), """,""", "|"), """", ""), "|")
        If UBound(f) >= 6 Then
            If IsNumeric(f(1)) And f(2) = "JT" Then
                sec = 6 + Val(f(1)) / 2
                th = Abs(Val(f(6)))
                If sec >= 7 And sec <= 12 And th > 0 Then
                    jMer(sec) = TQ * TL / (TE / (2# * (1# + TN)) * th)
                    n = n + 1
                End If
            End If
        End If
    Next r
    If n <> 6 Then msg = "twist read for " & n & " of 6 sections: " & Left$(resp, 300): Exit Function
    msg = "Girder J (merged -> Hambly, factor):"
    For no = 7 To 12
        If jMer(no) <= 0 Then msg = "no twist read for section " & no: Exit Function
        jHam(no) = bare(no) + SlabB(no) * tc ^ 3 / 6# / (eG / eS)
        jFac(no) = jHam(no) / jMer(no)
        If jFac(no) < 0.1 Or jFac(no) > 1.2 Then msg = "section " & no & " factor " & jFac(no) & " is not credible (merged " & jMer(no) & ", Hambly " & jHam(no) & ")": Exit Function
        msg = msg & vbCrLf & "  section " & no & ": " & Format(jMer(no), "0.00000") & " -> " & Format(jHam(no), "0.00000") & ", " & Format(jFac(no), "0.0000")
    Next no
    ' doc/IMPORTMXT replaces the whole model: removes the cantilevers, leaves sections 7-12
    Prog "girder J: remove test cantilevers"
    If Not Api("POST", "doc/IMPORTMXT", "{""Argument"":" & QS(mctPath) & "}", resp) Then msg = "MCT import after the measurement: " & Left$(resp, 300): Exit Function
    If Not Api("GET", "db/NODE", "", resp) Then msg = "db/NODE after the MCT import: " & Left$(resp, 200): Exit Function
    If InStr(resp, """X"":") > 0 Then msg = "test cantilevers still in the model after the MCT import": Exit Function
    If Not Api("GET", "db/SECT", "", resp) Then msg = "db/SECT after the MCT import: " & Left$(resp, 200): Exit Function
    If CountOf(resp, """SECTTYPE"":""COMPOSITE""") <> 6 Then msg = "composite sections missing after the second MCT import": Exit Function
    MeasureGirderJ = True
End Function

Private Function Put1(ByVal ep As String, ByVal body As String, ByRef msg As String) As Boolean
    Dim resp As String
    Put1 = Api("PUT", ep, body, resp)
    If Not Put1 Then msg = ep & ": " & Left$(resp, 300)
End Function

' Read-back of the stiffness factors: this API can accept a write without applying it, and a missing factor
' changes the modes without any message.
Private Function SecfCheck() As String
    Dim resp As String, n As Long, want As Long, no As Long, s As String, v As Double
    want = 3 + IIf(visW > 0, 1, 0) + IIf(dynMode, 6, 0)
    If Not Api("GET", "db/SECF", "", resp) Then SecfCheck = "  FAIL stiffness factors: " & Left$(resp, 200) & vbCrLf: Exit Function
    n = CountOf(resp, """IXX_SF"":")
    s = "  stiffness factors " & n & " / " & want & IIf(n = want, "  OK", "  FAIL") & vbCrLf
    If dynMode Then
        For no = 7 To 12
            v = SectField(resp, CStr(no), "IXX_SF")
            If Abs(v - jFac(no)) > 0.000001 Then s = s & "  FAIL section " & no & " factor " & v & " (sent " & jFac(no) & ")" & vbCrLf
        Next no
    Else
        If Not Api("GET", "db/STAG", "", resp) Then resp = ""
        If InStr(resp, "BG_SF") = 0 Then s = s & "  FAIL boundary group BG_SF is not activated in the stages." & vbCrLf
    End If
    SecfCheck = s
End Function

' IPART 2 = after composite; measured: IPART 1 has no effect on a composite section
Private Function GirderSecf() As String
    Dim no As Long, s As String
    If Not dynMode Then Exit Function
    For no = 7 To 12: s = s & "," & SecfJson(no, 1, jFac(no), 1, 1, 1, 2): Next no
    GirderSecf = s
End Function

' slab width of a composite section (as in SectMct)
Private Function SlabB(ByVal no As Long) As Double
    SlabB = sG
    If no = 9 Or no = 10 Then SlabB = sG / 2 + ovL
    If no = 11 Or no = 12 Then SlabB = sG / 2 + ovR
End Function

' ---- clean pass ----
' Every table the build writes, in REVERSE build order (combinations, stages and loads before
' the cases, groups, elements and nodes they reference; sections/materials last). UNIT, STYP and
' PJCF are single records (key 1) that the build overwrites, so they are not deleted. ESSF is no longer
' written; it stays in the list to clean models from older versions.
Private Function CleanTables() As Variant
    CleanTables = Array("LCOM-CONC", "LCOM-GEN", "STCT", "TDPL", "TDCS", "TDNA", "TDGR", "TDNT", "CSCS", "STAG", _
        "DYFG", "MVLDeu", "MVHL", "LLAN", "MVCD", "THGC", "EIGV", "ACTL", "LTOM", "BMLD", "BODF", "STMP", "LDGR", "STLD", "ESSF", "SECF", _
        "NSPR", "ELNK", "RIGD", "CONS", "BNGR", "GRUP", "ELEM", "NODE", "SECT", "TMAT", "TDME", "TDMT", "MATL")
End Function

' Keyless DELETE of every table (a failed DELETE on an already-empty table is harmless, so all
' are attempted), then a GET of every table. Returns "Clean pass: OK ..." or "STOPPED: ..." with
' the tables that still hold records or could not be read.
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

' ---- progress bar (status bar, like the culvert build: a UserForm would need a .frm/.frx) ----
' ProgReset sets the step count; every Prog call advances one step and shows the step that is
' about to run. Clamped at 100 % (extra clean passes just hold the bar full). The callers'
' existing "Application.StatusBar = False" on every exit path hands the bar back to Excel.
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
    Application.StatusBar = "MIDAS_USTYAPI  [" & String(filled, ChrW(9608)) & String(PROG_BAR_WIDTH - filled, ChrW(9617)) & _
        "] " & pct & "%   (" & progStep & "/" & progTotal & ")  " & stepName
    On Error GoTo 0
    DoEvents
End Sub

' Record keys of a GET answer {"<TABLE>":{"<key>":{...},"<key>":{...}}} as "k1","k2",...
' (keys at depth 2; strings and escapes skipped so braces inside text do not count).
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

' An empty table answers {"message":""}; a filled one {"<TABLE>":{"<key>":{...}}}.
Private Function TableHasRecords(ByVal resp As String, ByVal tbl As String) As Boolean
    TableHasRecords = (InStr(1, resp, """" & tbl & """:{""", vbBinaryCompare) > 0)
End Function

' Button / scripted: only the clean pass, no build (like CleanCulvertModel).
Public Sub MIDAS_USTYAPI_TEMIZLE(Optional ByVal reportFile As String = "", Optional ByVal apiKey As String = "")
    Dim rep As String, resp As String
    On Error GoTo Fail
    If apiKey <> "" Then mKey = apiKey
    If reportFile = "" Then
        If UiMsg("DELETE every table of the model open in Civil NX (nodes, elements, sections, loads, stages, combinations)?", _
                  vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) <> vbYes Then Exit Sub
    End If
    If Not GetKey() Then rep = "FAIL: no MAPI key."
    If rep = "" Then
        Set http = Nothing
        If Not Api("GET", "db/NODE", "", resp) Then
            rep = "FAIL: Civil NX does not answer (db/NODE): " & Left$(resp, 300)
        Else
            ProgReset 2 * (UBound(CleanTables()) + 1)
            rep = CleanModel()
        End If
    End If
    Application.StatusBar = False
    If reportFile = "" Then UiMsg FitBox(rep, "clean", 0), IIf(Left$(rep, 2) = "Cl", vbInformation, vbExclamation), MSG_TITLE Else WriteText reportFile, rep
    Exit Sub
Fail:
    Application.StatusBar = False
    If reportFile = "" Then UiMsg "Stopped: " & Err.Description, vbCritical, MSG_TITLE Else WriteText reportFile, "FAIL: " & Err.Description
End Sub

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

Public Sub MIDAS_USTYAPI_FORGET_KEY()
    mKey = ""
End Sub

Private Function SectArea(ByVal resp As String, ByVal key As String) As Double
    Dim p As Long, q As Long
    p = InStr(1, resp, """" & key & """:{")
    If p = 0 Then Exit Function
    p = InStr(p, resp, """AREA"":")
    If p = 0 Then Exit Function
    q = p + 7
    SectArea = Val(Mid$(resp, q, 30))
End Function

Private Function SectField(ByVal resp As String, ByVal key As String, ByVal fld As String) As Double
    Dim p As Long
    p = InStr(1, resp, """" & key & """:{")
    If p = 0 Then Exit Function
    p = InStr(p, resp, """" & fld & """:")
    If p = 0 Then Exit Function
    SectField = Val(Mid$(resp, p + Len(fld) + 3, 30))
End Function

Private Function CountOf(ByVal s As String, ByVal what As String) As Long
    CountOf = (Len(s) - Len(Replace(s, what, ""))) / Len(what)
End Function

Private Function Summary() As String
    ' compact: a MsgBox shows about 1024 characters (FitBox writes longer texts to a file)
    Dim s As String
    s = SCRIPT_ID & " " & SCRIPT_VERSION & vbCrLf & _
        bridgeName & ", span " & spanNo & " = " & Lc & " m; girder " & Lg & " m, a " & sark & " m, end zone " & zoneL & " m" & vbCrLf & _
        nG & " x " & KirisAdi() & " @ " & sG & " m, H " & Hg & ", A " & Format(polyArea, "0.00000") & " / Y " & Format(polyAreaY, "0.00000") & " m2, slab " & Format(tc, "0.000") & vbCrLf & _
        "Deck " & Format(yR - yL, "0.000") & " m, " & nTr & " track(s) on girders " & trG(1) + 1 & IIf(nTr = 2, "/" & trG(IIf(nTr = 2, 2, 1)) + 1, "") & vbCrLf & _
        "Model: " & nCnt & " nodes, " & eCnt & " elements, " & payloads.Count & " tables" & vbCrLf & _
        "Concrete " & clsG & " / " & clsS & ": " & IIf(dynMode, "Eurocode Ecm ", "AASHTO Ec ") & Format(eG, "0") & " / " & Format(eS, "0") & " MPa, ratio " & Format(eG / eS, "0.0000") & vbCrLf & _
        "Bearings: kE " & Format(kEB, "0") & ", ky " & Format(kyB, "0.0") & ", fixed " & Format(kFix, "0") & ", link " & linkLen & " m" & vbCrLf & _
        "Strips " & Format(dxAct, "0.000") & " m (asked " & stripDx & "), end " & Format(endStripW, "0.000") & ", diaphragm " & IIf(DiaFlange() > 0, "T " & Format(DiaFlange(), "0.000"), "rect.") & ", end beam " & disB & " x " & Format(DisHgt(), "0.000") & " x " & disL & _
        ", edge slab " & Format(sG / 2 + ovL, "0.000") & "/" & Format(sG / 2 + ovR, "0.000") & vbCrLf & _
        "Loads/girder: ballast " & qBal & ", prot. " & qKor & "; per edge: kerb " & qBor * nG / 2 & ", railing " & qKork * nG / 2 & ", panel " & qCep * nG / 2 & vbCrLf & _
        "Wind Wg " & wG & ", Wv " & wV & "; TU " & tuPos & "/" & tuNeg & "; alpha " & alphaLM & ", Phi3 " & Format(Phi3(Lc), "0.0000") & vbCrLf & _
        IIf(qtk > 0, "Centrifugal Qtk " & Format(qtk, "0.000") & " kN/m on " & nTr & " track(s), lever " & Format(railH + CF_HEIGHT, "0.00") & " m", "Centrifugal: none") & vbCrLf & _
        "Strands " & nTdnGroup & " in " & nTdn & " tendons, " & nRow & " rows, fpi " & fpi & " MPa; transfer " & Format(ageG, "0.00") & " d (f'ci " & Format(fciRatio, "0.00") & " f'c); CS " & DUR_CS1 & "/" & DUR_CS2 & "/" & DUR_CS3 & "/" & DUR_CS4 & " d, RH " & humid & " %"
    If dynMode Then
        s = "DYNAMIC MODEL" & vbCrLf & Left$(s, InStr(s, "Strands ") - 1) & "No stages, no tendons; girder J factors are measured during the build" & vbCrLf
        s = Left$(s, InStr(s, "Bearings: ") - 1) & "Supports on the girder nodes: start X Y, end Y, vertical springs kv " & Format(kEB, "0") & " (no links)" & Mid$(s, InStr(InStr(s, "Bearings: "), s, vbCrLf))
    End If
    If warnText <> "" Then s = "WARNINGS:" & vbCrLf & warnText & vbCrLf & s
    Summary = s
End Function

' MsgBox shows about 1024 characters: a longer text is written in full to <workbook>\MIDAS\<tag>.txt
' (or %TEMP%) and the box shows the start plus that path. reserve = characters the caller adds.
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

' ============================================================================
'  HELPERS
' ============================================================================
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

Private Function MaxQtk(ws As Worksheet) As Double
    ' largest number in the rows under the Qtk header of Tablo 17; an empty table shows "-" -> 0
    Dim c As Range, first As String, r As Long, v As Variant
    Set c = ws.UsedRange.Find(What:="Qtk", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
    If c Is Nothing Then Err.Raise vbObjectError + 24, , ws.Name & ": Tablo 17 header 'Qtk' not found."
    first = c.Address
    Do
        For r = 1 To 3
            v = c.Offset(r, 0).Value
            If Not IsError(v) Then
                If Not IsEmpty(v) And IsNumeric(v) Then
                    If CDbl(v) > MaxQtk Then MaxQtk = CDbl(v)
                End If
            End If
        Next r
        Set c = ws.UsedRange.FindNext(c)
    Loop While Not c Is Nothing And c.Address <> first
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
        Loop While Not c Is Nothing And c.Address <> first
    End If
    Err.Raise vbObjectError + 23, , ws.Name & ": header '" & header & "' with a number below not found."
End Function

Private Function SortedKeysL(d As Object) As Variant
    Dim a() As Long, i As Long, j As Long, t As Long, k As Variant
    ReDim a(1 To d.Count)
    For Each k In d.Keys: i = i + 1: a(i) = CLng(k): Next k
    For i = 1 To UBound(a) - 1
        For j = i + 1 To UBound(a)
            If a(j) < a(i) Then t = a(i): a(i) = a(j): a(j) = t
        Next j
    Next i
    SortedKeysL = a
End Function

' JSON number, always with a dot (Str ignores the Turkish decimal comma)
Private Function JN(ByVal x As Double) As String
    Dim s As String
    If Abs(x) < 0.000000000001 Then JN = "0": Exit Function
    s = Trim$(Str$(x))
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    JN = s
End Function

' JSON string, non-ASCII as \uXXXX
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

' project-information text: Turkish letters -> ASCII, other non-ASCII -> "?", at most 80 characters
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

Private Function SafeName(ByVal s As String) As String
    Dim i As Long, c As String, o As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c Like "[A-Za-z0-9_-]" Then o = o & c Else o = o & "_"
    Next i
    SafeName = o
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
