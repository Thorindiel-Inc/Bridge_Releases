Attribute VB_Name = "MIDAS_KOPRU"
Option Explicit
' ============================================================================
'  MIDAS_KOPRU - builds the general whole-bridge model in MIDAS Civil NX through the MAPI
'  REST API (reference models/ALTYAPI.mct). Runs from the 1-3A workbook.
'  Step A (owner 2026-11-05): geometry only - deck stick per span, deck-end slabs, bearings
'  on every girder, pier caps, piers, fixed bases and abutment bearing nodes.
'  No loads, spectra, combinations, stiffness factors or rail layer yet.
' ----------------------------------------------------------------------------
'  Button macro: MIDAS_KOPRU_OLUSTUR
'    1. asks for the 4-7 workbook (read-only): sarkiklik, ky, kv (1_RAPOR labels)
'    2. asks element length, bearing link length, fixed horizontal stiffness kh
'    3. reads 1-3A: span list, G_AYAK_TABLOSU (one row per axis), span table right of
'       G_SABIT_UC (fixed end SOL / SAG, Yatay Gabari, Kiris, Mesnet: left bearing of each
'       span, row 'Son' = last right), G_BASLIK_*, G_KOLON_BAGLANTI, G_AKS_MESNET, girder data.
'       One deck section per Yatay Gabari + Kiris combination. Joint spacing: B25, or the
'       cell of the same Hesap Acikligi in B25:G25 when filled.
'    4. summary; Yes = clean pass on the model open in Civil NX, MCT import of the
'       composite deck section (no JSON form; IMPORTMXT replaces the model, so it runs
'       first), all tables, read-back, doc/SAVE
'  Dry run (no dialogs, no HTTP): MIDAS_KOPRU_PAYLOAD path47, outDir
'  Levels: deck node Z 0 (section offset CT), bearing top -(H + tc), bearing bottom = cap
'  top node level, pier top = cap top - G_KOLON_BAGLANTI, pier base = pier top - height.
'  X: A1 bearing axis = 0; pier bearings G_AKS_MESNET from the pier axis.
'  Plan (owner 2026-11-05): G_GUZERGAH filled = axes on the alignment (km from G_KM_A1), each
'  span a straight chord between its axes, caps on the bisector of the two chords, piers and
'  bearing links turned with them; model X = chord A1 to A2. Empty table = straight. Level deck.
'  G_AKS_XY (axis X / Y from the DXF centerline, DXF_OKU 2026-11-05b) comes first: filled =
'  those points (A1 first, span count + 1 rows, chords checked against the span list).
'  Numbers stay below 10500 so the rail layer (RTI) can be added later.
'  MAPI key: INPUT!T3, else MIDAS_MAPI_KEY, else asked (memory only).
' ============================================================================

Private Const SCRIPT_ID As String = "midas-bridge-build"
Private Const SCRIPT_VERSION As String = "2026-11-05h"
Private Const SCRIPT_CHANGELOG As String = "Plan from G_AKS_XY (axis points of the DXF centerline) before G_GUZERGAH, STYP bMASSOFFSET true (consider mass at centroid); horizontal alignment from G_GUZERGAH: axes on the curve, chord spans, radial caps, piers and links turned; pier elements: at least 5 equal, none over 3 m; self-weight mass in X, Y and Z (was Z only), as ALTYAPI; deck types per span (Yatay Gabari + Kiris), box and solid piers; whole-section J, Asy, Asz of the deck solved in the macro; deck section with calculated properties (Civil NX refuses zero stiffness); clean pass skipped on an empty model; first version, step A: whole-bridge geometry (deck stick, bearings, caps, piers, fixed bases)"

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
Private Const MAX_TYPES As Long = 20
Private Const PI As Double = 3.14159265358979
Private Const ALIGN_DS As Double = 0.25        ' integration step along the alignment (m)
Private Const NUM_LIMIT As Long = 10500         ' rail layer numbers start here
Private Const SOLVE_H As Double = 0.05          ' deck section solver: coarse cell size (m); fine = half

' ---- inputs ----
Private nG As Long, sG As Double, Hg As Double, Lc As Double, deckW As Double
Private tc As Double, eG As Double, eS As Double, clsG As String, clsS As String, clsO As String
Private sark As Double, kyB As Double, kEB As Double, jointL As Double, eAx As Double
Private capS(1 To 5) As Double, capL As Double, capOff As Double, capDrop As Double
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
Private warnText As String, askUser As Boolean

' ---- model ----
Private nodeKey As Object
Private nX() As Double, nY() As Double, nZ() As Double, nCnt As Long
Private eMat() As Long, eSec() As Long, eN1() As Long, eN2() As Long, eCnt As Long
Private grpDeck As String, grpCap As String, grpPier As String
Private rigdJson As String, elnkJson As String, consJson As String, nLink As Long
Private capSlaves As Object, polyArea As Double, deckArea As Double
' plan: axis points, cap direction (unit), span chord direction (unit), chord length, heading (deg)
Private pxA() As Double, pyA() As Double, rxA() As Double, ryA() As Double, cxS() As Double, cyS() As Double, lenS() As Double, phiS() As Double, phiA() As Double
Private pierAng As Object, nAlign As Long, maxOff As Double, nAxXY As Long
Private gCz As Double, gIy As Double, gIz As Double, gPer As Double
Private dkLog As String

Private payloads As Collection
Private mKey As String
Private http As Object
Private progStep As Long, progTotal As Long
Private Const PROG_BAR_WIDTH As Long = 24

Public Sub MIDAS_KOPRU_OLUSTUR()
    RunBuild "", "", False
End Sub

Public Sub MIDAS_KOPRU_PAYLOAD(ByVal path47 As String, ByVal outDir As String)
    RunBuild path47, outDir, True
End Sub

' scripted build, no dialogs (tests)
Public Sub MIDAS_KOPRU_BUILD(ByVal path47 As String, ByVal reportFile As String, ByVal apiKey As String, Optional ByVal saveModel As Boolean = True)
    Dim rep As String
    On Error GoTo Fail
    warnText = ""
    meshDx = DEF_MESH: linkLen = DEF_LINK_LEN: kFix = DEF_K_FIX
    PrepareModel path47
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

Private Sub PrepareModel(ByVal path47 As String)
    Dim wb47 As Workbook, opened47 As Boolean
    Application.StatusBar = "MIDAS_KOPRU: reading inputs..."
    Set wb47 = OpenWb(path47, opened47)
    ReadInputs ThisWorkbook, wb47
    If opened47 Then wb47.Close SaveChanges:=False
    BuildModel
    BuildPayloads
End Sub

Private Sub RunBuild(ByVal path47 As String, ByVal outDir As String, ByVal dry As Boolean)
    Dim msg As String, rep As String, v As Variant
    On Error GoTo Fail
    warnText = ""
    meshDx = DEF_MESH: linkLen = DEF_LINK_LEN: kFix = DEF_K_FIX
    askUser = Not dry
    If path47 = "" Then
        v = Application.GetOpenFilename("Excel (*.xlsm;*.xlsx),*.xlsm;*.xlsx", , "4-7 (KDME) workbook: sarkiklik + bearing stiffness")
        If VarType(v) = vbBoolean Then Exit Sub
        path47 = CStr(v)
    End If
    If Not dry Then
        If Not AskModelValues() Then Exit Sub
    End If
    PrepareModel path47
    If dry Then
        WritePayloads outDir
        Application.StatusBar = False
        Exit Sub
    End If
    msg = FitBox(Summary(), "kopru_summary", 170) & vbCrLf & vbCrLf & _
          "YES = WIPE the open Civil NX model, rebuild it and SAVE it." & vbCrLf & _
          "NO = only write the JSON payloads (MIDAS\payload_kopru)."
    Select Case MsgBox(msg, vbYesNoCancel + vbQuestion, MSG_TITLE)
        Case vbNo
            WritePayloads ThisWorkbook.Path & "\MIDAS\payload_kopru"
            MsgBox "Payloads written to " & ThisWorkbook.Path & "\MIDAS\payload_kopru", vbInformation, MSG_TITLE
            Application.StatusBar = False
            Exit Sub
        Case vbCancel
            Application.StatusBar = False
            Exit Sub
    End Select
    rep = SendAll()
    Application.StatusBar = False
    MsgBox FitBox(rep, "kopru_report", 0), IIf(InStr(rep, "FAIL") > 0, vbExclamation, vbInformation), MSG_TITLE
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
        MsgBox "Stopped: " & d, vbCritical, MSG_TITLE
    End If
End Sub

Private Function AskModelValues() As Boolean
    Dim v As Variant
    v = Application.InputBox("Deck element length (m)." & vbCrLf & "The macro uses the nearest length that divides each span exactly.", MSG_TITLE & " - mesh", DEF_MESH, Type:=1)
    If VarType(v) = vbBoolean Then Exit Function
    If v <= 0 Then MsgBox "Element length must be > 0.", vbExclamation, MSG_TITLE: Exit Function
    meshDx = CDbl(v)
    v = Application.InputBox("Bearing link length (m): girder bottom to cap top.", MSG_TITLE & " - bearing", DEF_LINK_LEN, Type:=1)
    If VarType(v) = vbBoolean Then Exit Function
    If v <= 0 Then MsgBox "Link length must be > 0.", vbExclamation, MSG_TITLE: Exit Function
    linkLen = CDbl(v)
    v = Application.InputBox("Fixed horizontal bearing stiffness kh (kN/m)." & vbCrLf & "Used transversely at every bearing and longitudinally at the fixed end of each span.", MSG_TITLE & " - kh", DEF_K_FIX, Type:=1)
    If VarType(v) = vbBoolean Then Exit Function
    If v <= 0 Then MsgBox "kh must be > 0.", vbExclamation, MSG_TITLE: Exit Function
    kFix = CDbl(v)
    AskModelValues = True
End Function

Private Sub ReadInputs(wb As Workbook, wb47 As Workbook)
    Dim r As Range, gcol As Long, i As Long, v As Variant, n As Long, t As String
    projName = CStr(NmRange(wb, "G_PROJE_ADI").Cells(1).Value)
    clientName = CStr(NmRange(wb, "G_IDARE_ADI").Cells(1).Value)
    bridgeName = CStr(NmRange(wb, "G_KOPRU_ADI").Cells(1).Value)
    tc = Num(NmRange(wb, "G_TABLIYE_KALINLIGI").Cells(1).Offset(0, 1).Value, "Tabliye Kalinligi (INPUT column C)")
    Set r = NmRange(wb, "G_KIRIS_BETONU"): clsG = CStr(r.Cells(1, 1).Value)
    eG = Num(r.Cells(1, 3).Value, "Ecm kiris betonu (INPUT column D)")
    Set r = NmRange(wb, "G_TABLIYE_BETONU"): clsS = CStr(r.Cells(1, 1).Value)
    eS = Num(r.Cells(1, 3).Value, "Ecm tabliye betonu (INPUT column D)")
    clsO = CStr(NmRange(wb, "G_DIGER_BETONLAR").Cells(1, 1).Value)
    eAx = Num(NmRange(wb, "G_AKS_MESNET").Cells(1).Value, "G_AKS_MESNET")
    For i = 1 To 5
        capS(i) = Num(NmRange(wb, "G_BASLIK_KESIT").Cells(i).Value, "G_BASLIK_KESIT row " & i)
    Next i
    capL = Num(NmRange(wb, "G_BASLIK_BOYU").Cells(1).Value, "G_BASLIK_BOYU")
    capOff = NumOr0(NmRange(wb, "G_BASLIK_OFSET").Cells(1).Value)
    capDrop = Num(NmRange(wb, "G_KOLON_BAGLANTI").Cells(1).Value, "G_KOLON_BAGLANTI")

    sark = Num(NmRange(wb47, "G_SARKIKLIK").Cells(1).Value, "4-7 G_SARKIKLIK") / 100#
    kyB = Num(LabelValue(wb47.Worksheets("1_RAPOR"), 3, "Elastomer Yanal Kesme Rijitli", 19), "4-7 Elastomer Yanal Kesme Rijitligi")
    kEB = Num(LabelValue(wb47.Worksheets("1_RAPOR"), 3, "Elastomer Eksenel Rijitli", 19), "4-7 Elastomer Eksenel Rijitligi")

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
    Dim skewV As Variant
    skewV = NmRange(wb, "G_VEREVLIK").Cells(1).Value
    If NumOr0(skewV) <> 0 Then warnText = warnText & "- G_VEREVLIK = " & skewV & ": the model is built square." & vbCrLf
    For i = 1 To nDk
        If (tJoint(i) - tLc(i)) / 2 <= sark + 0.0005 Then warnText = warnText & "- " & tName(i) & ": joint spacing gives no deck-end slab." & vbCrLf
    Next i
End Sub

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

' plan of the axes. Curvature along the alignment: R < 0 = left, heading counter-clockwise in the model
' (checked against the V2 drawing); transition = straight line between the curvatures of its neighbours.
Private Sub BuildPlan()
    Dim t As Range, n As Long, r As Long, s0 As Double, i As Long
    Dim sA() As Double, sB() As Double, kA() As Double, kB() As Double, tp As String
    ReDim pxA(0 To nSp): ReDim pyA(0 To nSp): ReDim rxA(0 To nSp): ReDim ryA(0 To nSp)
    ReDim cxS(1 To nSp): ReDim cyS(1 To nSp): ReDim lenS(1 To nSp): ReDim phiS(1 To nSp): ReDim phiA(0 To nSp)
    Set t = NmRange(ThisWorkbook, "G_GUZERGAH")
    For r = 1 To t.Rows.Count
        If IsEmpty(t.Cells(r, 1).Value) Then Exit For
        n = n + 1
    Next r
    nAlign = n
    nAxXY = ReadAxisXY()
    If nAxXY > 0 Then nAlign = 0
    If nAxXY > 0 Then
        ' axis points read from G_AKS_XY (A1 at the origin)
    ElseIf n = 0 Then
        For i = 1 To nSp: pxA(i) = pxA(i - 1) + spL(i): Next i
    Else
        s0 = Num(NmRange(ThisWorkbook, "G_KM_A1").Cells(1).Value, "G_KM_A1")
        ReDim sA(1 To n): ReDim sB(1 To n): ReDim kA(1 To n): ReDim kB(1 To n)
        For r = 1 To n
            sA(r) = Num(t.Cells(r, 1).Value, "G_GUZERGAH row " & r) - s0
            sB(r) = Num(t.Cells(r, 2).Value, "G_GUZERGAH row " & r) - s0
            tp = UCase$(Trim$(CStr(t.Cells(r, 3).Value)))
            If tp = "KURP" Then
                kA(r) = -1 / Num(t.Cells(r, 4).Value, "G_GUZERGAH R row " & r): kB(r) = kA(r)
            ElseIf tp <> "DOGRU" And tp <> "GECIS" Then
                Err.Raise vbObjectError + 41, , "G_GUZERGAH row " & r & ": Tip must be KURP, GECIS or DOGRU."
            End If
        Next r
        For r = 1 To n
            If UCase$(Trim$(CStr(t.Cells(r, 3).Value))) = "GECIS" Then
                If r = 1 Or r = n Then Err.Raise vbObjectError + 42, , "G_GUZERGAH: a transition at the start or end of the bridge is not supported yet."
                kA(r) = kB(r - 1): kB(r) = kA(r + 1)
            End If
        Next r
        ' walk along the alignment, store the point at each axis
        Dim x As Double, y As Double, th As Double, sNow As Double, ds As Double, ax As Long, sNext As Double, cum As Double, e As Long, km As Double, sM As Double
        ax = 0: cum = 0: e = 1
        Do While ax <= nSp
            sNext = sNow + ALIGN_DS
            If sNext > cum Then sNext = cum
            ds = sNext - sNow
            If ds > 0 Then
                sM = sNow + ds / 2
                Do While e < n And sM > sB(e): e = e + 1: Loop
                km = kA(e) + (kB(e) - kA(e)) * (sM - sA(e)) / (sB(e) - sA(e))
                x = x + ds * Cos(th + km * ds / 2)
                y = y + ds * Sin(th + km * ds / 2)
                th = th + km * ds
                sNow = sNext
            End If
            If sNow >= cum - 0.0000001 Then
                pxA(ax) = x: pyA(ax) = y
                ax = ax + 1
                If ax <= nSp Then cum = cum + spL(ax)
            End If
        Loop
    End If
    If nAxXY > 0 Or nAlign > 0 Then
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
    Dim nx As Double, ny As Double, q As Double
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
            ch = Sqr((pxA(i) - pxA(i - 1)) ^ 2 + (pyA(i) - pyA(i - 1)) ^ 2)
            If Abs(ch - spL(i)) > 0.05 Then Err.Raise vbObjectError + 44, , "G_AKS_XY: span " & i & " chord " & Format(ch, "0.000") & " m, span list " & Format(spL(i), "0.000") & " m. Run DXF OKU again."
        End If
    Next i
    ReadAxisXY = n
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

Private Sub BuildModel()
    Dim i As Long, j As Long, m As Long, k As Long, a As Long, b As Long, xL As Double, xR As Double, ext As Double
    Dim zLs() As Double, y() As Double, dx As Double, zCap() As Double
    Set nodeKey = CreateObject("Scripting.Dictionary")
    Set capSlaves = CreateObject("Scripting.Dictionary")
    nCnt = 0: eCnt = 0: nLink = 0
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

    For i = 1 To nSp
        UseType spType(i)
        xL = IIf(i = 1, 0, eAx)
        xR = lenS(i) - IIf(i = nSp, 0, eAx)
        If Abs(xR - xL - Lc) > 0.01 Then warnText = warnText & "- Span " & i & ": bearing to bearing " & Format(xR - xL, "0.000") & " m, Hesap Acikligi " & Lc & " m." & vbCrLf
        ext = (jointL - Lc) / 2
        m = CLng((xR - xL) / meshDx): If m < 1 Then m = 1
        dx = (xR - xL) / m
        If i > 1 And ext > sark + 0.0005 Then Member grpDeck, 2, SlabSec(spType(i)), SpN(i, xL - ext, 0, 0), SpN(i, xL - sark, 0, 0)
        Member grpDeck, 1, DeckSec(spType(i)), SpN(i, xL - sark, 0, 0), SpN(i, xL, 0, 0)
        For k = 1 To m
            Member grpDeck, 1, DeckSec(spType(i)), SpN(i, xL + (k - 1) * dx, 0, 0), SpN(i, IIf(k = m, xR, xL + k * dx), 0, 0)
        Next k
        Member grpDeck, 1, DeckSec(spType(i)), SpN(i, xR, 0, 0), SpN(i, xR + sark, 0, 0)
        If i < nSp And ext > sark + 0.0005 Then Member grpDeck, 2, SlabSec(spType(i)), SpN(i, xR + sark, 0, 0), SpN(i, xR + ext, 0, 0)
    Next i

    For i = 1 To nSp
        UseType spType(i)
        ReDim y(0 To nG - 1)
        For j = 0 To nG - 1: y(j) = (j - (nG - 1) / 2#) * sG: Next j
        xL = IIf(i = 1, 0, eAx)
        xR = lenS(i) - IIf(i = nSp, 0, eAx)
        BearingLine i, xL, i - 1, (i = 1), fixLeft(i), y, zLs(i) + linkLen, zLs(i), zCap(i - 1)
        BearingLine i, xR, i, (i = nSp), Not fixLeft(i), y, zLs(i) + linkLen, zLs(i), zCap(i)
    Next i

    ' caps and piers
    Dim ys As Object, yk As Variant, top As Long, q As Variant, sp As Variant, nEl As Long
    For i = 1 To nSp - 1
        Set ys = CreateObject("Scripting.Dictionary")
        AddX ys, -capL / 2: AddX ys, capL / 2: AddX ys, 0
        For Each sp In Array(i, i + 1)
            For j = 0 To tNG(spType(sp)) - 1: AddX ys, (j - (tNG(spType(sp)) - 1) / 2#) * tSG(spType(sp)): Next j
        Next sp
        yk = SortedKeys(ys)
        For k = UBound(yk) To 2 Step -1
            Member grpCap, 3, SEC_CAP, CapN(i, yk(k), zCap(i)), CapN(i, yk(k - 1), zCap(i))
        Next k
        top = CapN(i, 0, zCap(i) - capDrop)
        Rigid CapN(i, 0, zCap(i)), CStr(top)
        a = top
        nEl = -Int(-pH(i) / MAX_PIER_EL + 0.000000001)
        If nEl < N_PIER_EL Then nEl = N_PIER_EL
        For k = 1 To nEl
            b = CapN(i, 0, zCap(i) - capDrop - pH(i) * k / nEl)
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
    Dim j As Long, up As Long, lo As Long, sl As String, cn As Long, kLong As Double
    kLong = IIf(fixedEnd, kFix, kyB)
    For j = 0 To nG - 1
        up = SpN(sp, u, y(j), zU)
        lo = SpN(sp, u, y(j), zL)
        sl = sl & IIf(sl = "", "", ",") & up
        nLink = nLink + 1
        elnkJson = elnkJson & IIf(elnkJson = "", "", ",") & QS(CStr(nLink)) & ":{""NODE"":[" & up & "," & lo & "],""ANGLE"":" & JN(-phiS(sp)) & ",""LINK"":""GEN""," & _
            """R_S"":[false,false,false,false,false,false],""SDR"":[" & JN(kEB) & "," & JN(kFix) & "," & JN(kLong) & ",0,0,0],""bSHEAR"":false,""DR"":[0.5,0.5],""BNGR_NAME"":""BG_MESNET""}"
        If abut Then
            Fix1 lo
        Else
            cn = CapN(axisNo, y(j), zC)
            If capSlaves.Exists(CStr(cn)) Then
                capSlaves(CStr(cn)) = capSlaves(CStr(cn)) & "," & lo
            Else
                capSlaves(CStr(cn)) = CStr(lo)
            End If
        End If
    Next j
    Rigid SpN(sp, u, 0, 0), sl
End Sub

Private Sub Rigid(ByVal master As Long, ByVal slaves As String)
    rigdJson = rigdJson & IIf(rigdJson = "", "", ",") & QS(CStr(master)) & ":{""ITEMS"":[{""ID"":1,""GROUP_NAME"":""BG_MESNET"",""DOF"":111111,""S_NODE"":[" & slaves & "]}]}"
End Sub

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

Private Function MatJson(ByVal no As Long, ByVal cls As String) As String
    MatJson = QS(CStr(no)) & ":{""TYPE"":""CONC"",""NAME"":" & QS(cls) & ",""HE_SPEC"":0,""HE_COND"":0,""bMASS_DENS"":false,""DAMP_RAT"":0.05,""PARAM"":[{""P_TYPE"":1,""STANDARD"":""EN04(RC)"",""CODE"":"""",""DB"":" & QS(EnDb(cls)) & "}]}"
End Function

Private Sub BuildPayloads()
    Dim s As String, i As Long, sz As String
    Set payloads = New Collection
    Add "db/UNIT", "{""Assign"":{""1"":{""FORCE"":""KN"",""DIST"":""M"",""HEAT"":""J"",""TEMPER"":""C""}}}"
    Add "db/STYP", "{""Assign"":{""1"":{""STYP"":0,""MASS"":1,""GRAV"":" & JN(GRAV_STYP) & ",""TEMP"":0,""bALIGNBEAM"":false,""bALIGNSLAB"":false,""bMASSOFFSET"":true,""bSELFWEIGHT"":true,""SMASS"":1}}}"
    Add "db/PJCF", "{""Assign"":{""1"":{""PROJECT"":" & QS(Pj(projName)) & ",""USER"":" & QS(Pj(Application.UserName)) & ",""CLIENT"":" & QS(Pj(clientName)) & _
        ",""TITLE"":" & QS(Pj(bridgeName)) & ",""COMMENT"":" & QS(Pj(nSp & " spans; " & SCRIPT_ID & " " & SCRIPT_VERSION)) & "}}}"
    Add "db/MATL", "{""Assign"":{" & MatJson(1, clsG) & "," & MatJson(2, clsS) & "," & MatJson(3, clsO) & "}}"
    s = ""
    For i = 1 To nDk
        s = s & DbSect(SlabSec(i), "DOSEME" & IIf(i = 1, "", "_" & i), SectHead2("CT", 0, 0), "SB", JN(tc) & "," & JN(tW(i)) & ",0,0,0,0,0,0,0,0") & ","
    Next i
    s = s & DbSect(SEC_CAP, "BASLIK", SectHead2("CT", 1, capOff), "UDT", JN(capS(1)) & "," & JN(capS(2)) & "," & JN(capS(3)) & "," & JN(capS(4)) & "," & JN(capS(5)) & ",0,0,0,0,0")
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
    Add "db/SECT", "{""Assign"":{" & s & "}}"
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
    Add "db/GRUP", "{""Assign"":{""1"":{""NAME"":""SG_TABLIYE"",""P_TYPE"":0,""E_LIST"":[" & grpDeck & "]}," & _
        """2"":{""NAME"":""SG_BASLIK"",""P_TYPE"":0,""E_LIST"":[" & grpCap & "]},""3"":{""NAME"":""SG_KOLON"",""P_TYPE"":0,""E_LIST"":[" & grpPier & "]}}}"
    Add "db/BNGR", "{""Assign"":{""1"":{""NAME"":""BG_MESNET"",""AUTOTYPE"":0}}}"
    Add "db/CONS", "{""Assign"":{" & consJson & "}}"
    Add "db/RIGD", "{""Assign"":{" & rigdJson & "}}"
    Add "db/ELNK", "{""Assign"":{" & elnkJson & "}}"
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
        Prog "table " & i & "/" & payloads.Count & "  " & p(0)
        If Not Api("PUT", CStr(p(0)), CStr(p(1)), resp) Then
            SendAll = cleanRep & vbCrLf & "FAIL at step " & i & " (" & p(0) & "):" & vbCrLf & Left$(resp, 600) & vbCrLf & vbCrLf & "Model is incomplete; nothing saved."
            Exit Function
        End If
    Next i
    rep = "Read-back:" & vbCrLf
    Prog "read-back nodes"
    If Api("GET", "db/NODE", "", resp) Then rep = rep & "  nodes " & CountOf(resp, """X"":") & " / " & nCnt & IIf(CountOf(resp, """X"":") = nCnt, "", "  FAIL") & vbCrLf
    Prog "read-back elements"
    If Api("GET", "db/ELEM", "", resp) Then rep = rep & "  elements " & CountOf(resp, """TYPE"":") & " / " & eCnt & IIf(CountOf(resp, """TYPE"":") = eCnt, "", "  FAIL") & vbCrLf
    Prog "read-back sections"
    If Api("GET", "db/SECT", "", resp) Then
        UseType 1
        a = SectArea(resp, CStr(DeckSec(1)))
        rep = rep & "  deck section area read " & Format(a, "0.0000") & " m2 (girders " & Format(nG * polyArea, "0.0000") & ", transformed " & Format(deckArea, "0.0000") & ")" & IIf(a > 0, "", "  FAIL: no stiffness in the composite section") & vbCrLf
    End If
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
        If MsgBox(txt & vbCrLf & vbCrLf & "WIPE it anyway?", vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) = vbYes Then Exit Function
    End If
    msg = "STOPPED: " & txt & " Nothing was changed. Open the right .mcb, or use an empty file."
    ModelOk = False
End Function

Private Function Summary() As String
    Dim s As String, i As Long, f As Long, tot As Double
    For i = 1 To nSp
        tot = tot + spL(i)
        If fixLeft(i) Then f = f + 1
    Next i
    s = SCRIPT_ID & " " & SCRIPT_VERSION & vbCrLf & bridgeName & ": " & nSp & " spans, length " & Format(tot, "0.00") & " m, " & nSp - 1 & " piers" & vbCrLf & _
        dkLog & "Slab " & Format(tc, "0.000") & " m, sarkiklik " & sark & " m; concrete " & clsG & " / " & clsS & " / " & clsO & " (MIDAS EN04 database)" & vbCrLf & _
        "Bearings: kv " & Format(kEB, "0") & ", free ky " & Format(kyB, "0.0") & ", kh " & Format(kFix, "0") & ", link " & linkLen & " m; " & f & " spans fixed left, " & nSp - f & " right" & vbCrLf & _
        "Cap " & capL & " m, pier top " & capDrop & " m below cap top; " & nPS & " pier section(s)" & vbCrLf & _
        "Model: " & nCnt & " nodes, " & eCnt & " elements, " & nLink & " bearings, element " & meshDx & " m" & vbCrLf & _
        IIf(nAxXY > 0, "Plan: " & nAxXY & " axis points from G_AKS_XY (DXF centerline), largest offset from the A1-A2 chord " & Format(maxOff, "0.00") & " m", IIf(nAlign > 0, "Plan: " & nAlign & " alignment elements, largest offset from the A1-A2 chord " & Format(maxOff, "0.00") & " m", "Plan: straight (G_AKS_XY and G_GUZERGAH empty)")) & vbCrLf & _
        "Geometry only: no loads, no spectra, no combinations."
    If warnText <> "" Then s = "WARNINGS:" & vbCrLf & warnText & vbCrLf & s
    Summary = s
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
    CleanTables = Array("LCOM-CONC", "LCOM-GEN", "STCT", "TDPL", "TDCS", "TDNA", "TDGR", "TDNT", "CSCS", "STAG", _
        "DYFG", "MVLDeu", "MVHL", "LLAN", "MVCD", "THGC", "EIGV", "ACTL", "LTOM", "BMLD", "BODF", "STMP", "LDGR", "STLD", "ESSF", "SECF", _
        "NSPR", "ELNK", "RIGD", "CONS", "BNGR", "GRUP", "ELEM", "NODE", "SECT", "TMAT", "TDME", "TDMT", "MATL")
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
    If mKey = "" Then mKey = InputBox("MIDAS MAPI key (kept in memory for this Excel session only):", MSG_TITLE)
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

