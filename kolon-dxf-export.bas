Attribute VB_Name = "KOLON_DXF"
Option Explicit
' ============================================================================
'  KOLON_DXF - generates a hollow rectangular bridge pier section DXF for
'  MIDAS GSD (General Section Designer).
'  Managed module (UPDATE feed, kolon-dxf-export). Runs from the macro window (MAKROLAR,
'  row 'Kolon kesiti DXF (GSD)', inputs in its fields) or by Alt+F8 KOLON_DXF_OLUSTUR
'  (inputs asked one by one). Messages go through PANEL (UiMsg / UiInput / UiLog), so the
'  workbook needs the module PANEL, as every workbook with the MAKROLAR button has.
' ----------------------------------------------------------------------------
'  Rules preserved automatically:
'    - Outer box B x H with corner chamfers, inner void (B - 2t) x (H - 2t),
'      centered at (0, 0), units in meters (AC1021 DXF: LWPOLYLINE + INSERT).
'    - 16 fixed anchor bars: 4 outer corners, 4 inner corners, and 8 outer
'      bars projected opposite the inner corners.
'    - All Row 1 bars align face-to-face across every wall (3 exact zones:
'      corner leg Lc, long middle Lx, short middle Ly; N1 = 4*(2*nc+nx+ny)).
'    - Actual reinforcement ratio stays <= target ro (largest valid count).
'    - Stops when t - 2d', the void width or the void height is below s_min
'      (the fixed anchor bars would sit closer than s_min), when d' <= phi/2,
'      or when the chamfer cuts the outer corner bar.
'    - When target ro requires tighter spacing than s_min, adds doubly
'      symmetric Row 2 bars directly behind Row 1 bars (offset = 2 * phi),
'      filling outer top/bottom faces first, then outer left/right, then
'      inner faces if needed. A Row 2 bar closer than the row offset to any
'      bar already placed is skipped (outer corners where two faces meet).
'  Layers:
'    - BETON-30 : closed outer LWPOLYLINE (with chamfers) + closed inner void
'    - DEMIR-30 : longitudinal rebars = INSERTs of block B (filled donut,
'                 30 mm base diameter scaled by phi / 30)
'  Scripted test entry (no dialogs):
'    KOLON_DXF_BUILD B, H, t, chf, cov, phiMm, sMin, targetRo, outPath
' ============================================================================

Private Const SCRIPT_ID As String = "kolon-dxf-export"
Private Const SCRIPT_VERSION As String = "2026-11-19d"
Private Const SCRIPT_CHANGELOG As String = "New managed module: hollow rectangular pier section DXF (AutoCAD 2007, AC1021) for MIDAS GSD, concrete on BETON-30, bars on DEMIR-30; macro-window row Kolon kesiti DXF (GSD) with input fields ('@TOOL line), messages through the window; stops on a wall or void narrower than s_min, cover <= phi/2 or a chamfer that cuts the corner bar; Row 2 bars never closer than the row offset to another bar"
'@TOOL 87|MIDAS sonu\u00E7lar\u0131|Kolon kesiti DXF (GSD)|KOLON_DXF.KOLON_DXF_OLUSTUR||Bo\u015Fluklu dikd\u00F6rtgen ayak kesitinin MIDAS GSD i\u00E7in DXF \u00E7izimini yazar: d\u0131\u015F hat ve bo\u015Fluk BETON-30, boyuna donat\u0131lar DEMIR-30 katman\u0131nda. Donat\u0131 oran\u0131 hedefi a\u015Fmaz; 1. s\u0131ra s_min ile s\u0131\u011Fmazsa 2. s\u0131ra eklenir. Klas\u00F6r her \u00E7al\u0131\u015Ft\u0131rmada sorulur; \u00E7al\u0131\u015Fma kitab\u0131na bir \u015Fey yazmaz.|b~D\u0131\u015F geni\u015Flik B (m)~n~6.5;h~D\u0131\u015F y\u00FCkseklik H (m)~n~4;t~Perde kal\u0131nl\u0131\u011F\u0131 t (m)~n~1;pah~K\u00F6\u015Fe pah\u0131~c~Var/Yok;c~Pah boyutu c (m)~n~0.05;d~Paspay\u0131 d' (donat\u0131 merkezine, m)~n~0.086;phi~Donat\u0131 \u00E7ap\u0131 (mm)~n~32;smin~1. s\u0131ra en k\u00FC\u00E7\u00FCk aral\u0131k s_min (m)~n~0.115;ro~Hedef donat\u0131 oran\u0131 (%)~n~1.01

Private Const MSG_TITLE As String = "KOLON DXF (MIDAS GSD)"
Private Const LAY_CONC As String = "BETON-30"
Private Const LAY_REBAR As String = "DEMIR-30"

Private Const DEF_B As Double = 6.5
Private Const DEF_H As Double = 4#
Private Const DEF_T As Double = 1#
Private Const DEF_CHAMFER As Double = 0.05
Private Const DEF_COVER As Double = 0.086
Private Const DEF_PHI_MM As Double = 32#
Private Const DEF_SMIN As Double = 0.115
Private Const DEF_RO As Double = 1.01

Private mX() As Double
Private mY() As Double
Private mCount As Long

Public Sub KOLON_DXF_OLUSTUR()
    Dim B As Double, H As Double, t As Double, chf As Double
    Dim cov As Double, phiMm As Double, sMin As Double, targetRo As Double
    Dim nc As Long, nx As Long, ny As Long, n1 As Long, n2 As Long, nTot As Long
    Dim ac As Double, ab As Double, actualRo As Double
    Dim sc As Double, sx As Double, sy As Double
    Dim folderPath As String, fileName As String, outPath As String
    Dim summary As String, ans As VbMsgBoxResult

    On Error GoTo Fail

    If UiHas("b") Then
        ' started from the macro window: its fields (keys of the '@TOOL line)
        B = UiNum("b", DEF_B): H = UiNum("h", DEF_H): t = UiNum("t", DEF_T)
        chf = 0#
        If UiIdx("pah", 1) = 1 Then chf = UiNum("c", DEF_CHAMFER)
        cov = UiNum("d", DEF_COVER): phiMm = UiNum("phi", DEF_PHI_MM)
        sMin = UiNum("smin", DEF_SMIN): targetRo = UiNum("ro", DEF_RO)
    Else
        If Not AskDouble("Dis genislik B (m):", DEF_B, B) Then Exit Sub
        If Not AskDouble("Dis yukseklik H (m):", DEF_H, H) Then Exit Sub
        If Not AskDouble("Perde kalinligi t (m):", DEF_T, t) Then Exit Sub
        If Not AskDouble("Kose pah boyutu c (m, 0 = pahsiz):", DEF_CHAMFER, chf) Then Exit Sub
        If Not AskDouble("Donati merkezine paspayi d' (m):", DEF_COVER, cov) Then Exit Sub
        If Not AskDouble("Boyuna donati capi phi (mm):", DEF_PHI_MM, phiMm) Then Exit Sub
        If Not AskDouble("1. sira min. donati araligi s_min (m):", DEF_SMIN, sMin) Then Exit Sub
        If Not AskDouble("Hedef donati orani ro (%):", DEF_RO, targetRo) Then Exit Sub
    End If

    ValidateInputs B, H, t, chf, cov, phiMm, sMin, targetRo
    SolveLayout B, H, t, chf, cov, phiMm, sMin, targetRo, nc, nx, ny, n1, n2, ac, ab, actualRo, sc, sx, sy
    nTot = n1 + n2

    folderPath = PickFolder()
    If folderPath = "" Then Exit Sub

    fileName = "KOLON_" & Format$(Round(B * 100#, 0), "0") & "x" & Format$(Round(H * 100#, 0), "0") & _
        "_phi" & Format$(Round(phiMm, 0), "0") & "_ro" & DotFixed(actualRo, 2) & ".dxf"
    outPath = folderPath
    If Right$(outPath, 1) <> "\" Then outPath = outPath & "\"
    outPath = outPath & fileName

    summary = "Kesit: " & DotFixed(B, 2) & " x " & DotFixed(H, 2) & " m  (t = " & DotFixed(t, 2) & " m, pah = " & DotFixed(chf * 100#, 1) & " cm)" & vbCrLf & _
        "Net beton alani Ac = " & DotFixed(ac, 4) & " m2" & vbCrLf & _
        "Donati: " & nTot & " phi" & Format$(Round(phiMm, 0), "0") & "  (1. sira: " & n1 & ", 2. sira: " & n2 & ")" & vbCrLf & _
        "Gercek ro = " & DotFixed(actualRo, 3) & "%  (hedef <= " & DotFixed(targetRo, 3) & "%)" & vbCrLf & _
        "1. sira bolumleri: nc=" & nc & " (" & DotFixed(sc * 100#, 2) & " cm), nx=" & nx & " (" & DotFixed(sx * 100#, 2) & " cm), ny=" & ny & " (" & DotFixed(sy * 100#, 2) & " cm)" & vbCrLf & vbCrLf & _
        "Dosya: " & outPath & vbCrLf & vbCrLf & _
        "DXF olusturulsun mu?"

    ans = UiMsg(summary, vbYesNo + vbQuestion, MSG_TITLE)
    If ans <> vbYes Then UiLog "DXF yazilmadi.": Exit Sub

    WriteBoxDxf outPath, B, H, t, chf, phiMm
    UiMsg "DXF olusturuldu (" & nTot & " donati, ro = " & DotFixed(actualRo, 3) & "%):" & vbCrLf & outPath, vbInformation, MSG_TITLE
    Exit Sub

Fail:
    UiMsg "Hata: " & Err.Description, vbExclamation, MSG_TITLE
End Sub

Public Function KOLON_DXF_BUILD(ByVal B As Double, ByVal H As Double, ByVal t As Double, _
    ByVal chf As Double, ByVal cov As Double, ByVal phiMm As Double, _
    ByVal sMin As Double, ByVal targetRo As Double, ByVal outPath As String) As String
    Dim nc As Long, nx As Long, ny As Long, n1 As Long, n2 As Long
    Dim ac As Double, ab As Double, actualRo As Double
    Dim sc As Double, sx As Double, sy As Double

    ValidateInputs B, H, t, chf, cov, phiMm, sMin, targetRo
    SolveLayout B, H, t, chf, cov, phiMm, sMin, targetRo, nc, nx, ny, n1, n2, ac, ab, actualRo, sc, sx, sy
    WriteBoxDxf outPath, B, H, t, chf, phiMm

    KOLON_DXF_BUILD = "OK|N=" & (n1 + n2) & "|N1=" & n1 & "|N2=" & n2 & _
        "|ro=" & DotFixed(actualRo, 4) & "|nc=" & nc & "|nx=" & nx & "|ny=" & ny & _
        "|sc=" & DotFixed(sc * 100#, 2) & "|sx=" & DotFixed(sx * 100#, 2) & "|sy=" & DotFixed(sy * 100#, 2)
End Function

Private Sub ValidateInputs(ByVal B As Double, ByVal H As Double, ByVal t As Double, _
    ByVal chf As Double, ByVal cov As Double, ByVal phiMm As Double, _
    ByVal sMin As Double, ByVal targetRo As Double)
    Dim phiM As Double
    phiM = phiMm / 1000#
    If B <= 0# Or H <= 0# Or t <= 0# Then Err.Raise vbObjectError + 1, , "B, H ve t pozitif olmalidir."
    If B <= 2# * t Or H <= 2# * t Then Err.Raise vbObjectError + 2, , "Ic bosluk pozitif olmalidir (B > 2t ve H > 2t)."
    If chf < 0# Or chf >= t / 2# Then Err.Raise vbObjectError + 3, , "Pah boyutu gecersiz."
    If cov <= 0# Or 2# * cov >= t Then Err.Raise vbObjectError + 4, , "Paspayi perde kalinligina gore cok buyuk (2*d' < t olmali)."
    If phiMm <= 0# Then Err.Raise vbObjectError + 5, , "Donati capi pozitif olmalidir."
    If sMin <= phiMm / 1000# Then Err.Raise vbObjectError + 6, , "Min. aralik donati capindan buyuk olmalidir."
    If targetRo <= 0# Or targetRo > 8# Then Err.Raise vbObjectError + 7, , "Hedef ro %0 ile %8 arasinda olmalidir."
    If cov <= phiM / 2# Then Err.Raise vbObjectError + 8, , "Paspayi d' donati yaricapindan buyuk olmalidir (d' > phi/2)."
    ' Fixed anchor bars: corner leg t - 2d', void width (B - 2t) + 2d', void height (H - 2t) + 2d'
    If t - 2# * cov < sMin - 0.0000001 Then Err.Raise vbObjectError + 12, , _
        "Perde ince: t - 2d' = " & DotFixed((t - 2# * cov) * 100#, 1) & " cm < s_min = " & DotFixed(sMin * 100#, 1) & " cm (sabit ankraj donatilari s_min'den yakin olur)."
    If B - 2# * t + 2# * cov < sMin - 0.0000001 Or H - 2# * t + 2# * cov < sMin - 0.0000001 Then _
        Err.Raise vbObjectError + 13, , "Ic bosluk s_min icin cok dar (ic kose donatilari s_min'den yakin olur)."
    ' Chamfer face distance to the outer corner bar centre = (2d' - c) / sqrt(2)
    If chf > 0# And (2# * cov - chf) / Sqr(2#) < phiM / 2# Then Err.Raise vbObjectError + 14, , _
        "Pah kose donatisini kesiyor: (2d' - c)/kok2 = " & DotFixed((2# * cov - chf) / Sqr(2#) * 1000#, 1) & " mm < phi/2."
End Sub

Private Sub SolveLayout(ByVal B As Double, ByVal H As Double, ByVal t As Double, _
    ByVal chf As Double, ByVal cov As Double, ByVal phiMm As Double, _
    ByVal sMin As Double, ByVal targetRo As Double, _
    ByRef bestNc As Long, ByRef bestNx As Long, ByRef bestNy As Long, _
    ByRef n1 As Long, ByRef n2 As Long, ByRef ac As Double, ByRef ab As Double, _
    ByRef actualRo As Double, ByRef sc As Double, ByRef sx As Double, ByRef sy As Double)

    Dim pi As Double, bIn As Double, hIn As Double
    Dim xOut As Double, yOut As Double, xIn As Double, yIn As Double
    Dim lc As Double, lx As Double, ly As Double
    Dim nMax As Long, maxNc As Long, maxNx As Long, maxNy As Long, n1Cap As Long
    Dim nc As Long, nx As Long, ny As Long, nCand As Long
    Dim dSc As Double, dSx As Double, dSy As Double, spMax As Double, spMin As Double, spread As Double
    Dim bestN1 As Long, bestSpread As Double
    Dim phiM As Double, row2Off As Double, need2 As Long

    pi = 4# * Atn(1#)
    bIn = B - 2# * t
    hIn = H - 2# * t
    ac = B * H - bIn * hIn - 2# * chf * chf
    phiM = phiMm / 1000#
    ab = pi * phiM * phiM / 4#

    nMax = Int(((targetRo / 100#) * ac / ab) + 0.000000001)
    If nMax < 16 Then Err.Raise vbObjectError + 10, , "Hedef ro cok dusuk: en az 16 kose/ankraj donatisi gerekir (min ro = " & DotFixed(16# * ab / ac * 100#, 3) & "%)."

    xOut = B / 2# - cov
    yOut = H / 2# - cov
    xIn = bIn / 2# + cov
    yIn = hIn / 2# + cov

    lc = xOut - xIn
    lx = 2# * xIn
    ly = 2# * yIn

    maxNc = Int(lc / sMin + 0.0000001): If maxNc < 1 Then maxNc = 1
    maxNx = Int(lx / sMin + 0.0000001): If maxNx < 1 Then maxNx = 1
    maxNy = Int(ly / sMin + 0.0000001): If maxNy < 1 Then maxNy = 1
    n1Cap = 4 * (2 * maxNc + maxNx + maxNy)

    bestN1 = -1
    bestSpread = 1E+30

    If nMax <= n1Cap Then
        ' Single-row fits with all spacings >= sMin: maximize N1 <= nMax, then minimize spacing spread
        For nc = 1 To maxNc
            dSc = lc / nc
            For nx = 1 To maxNx
                dSx = lx / nx
                For ny = 1 To maxNy
                    nCand = 4 * (2 * nc + nx + ny)
                    If nCand <= nMax Then
                        dSy = ly / ny
                        spMax = Max3(dSc, dSx, dSy)
                        spMin = Min3(dSc, dSx, dSy)
                        spread = spMax - spMin
                        If nCand > bestN1 Or (nCand = bestN1 And spread < bestSpread - 0.0000001) Then
                            bestN1 = nCand
                            bestSpread = spread
                            bestNc = nc
                            bestNx = nx
                            bestNy = ny
                        End If
                    End If
                Next ny
            Next nx
        Next nc
    Else
        ' Row 1 at tightest uniform spacing near sMin, then add Row 2 bars
        bestNc = maxNc
        bestNx = maxNx
        bestNy = maxNy
        bestN1 = n1Cap
    End If

    If bestN1 < 16 Then Err.Raise vbObjectError + 11, , "Uygun 1. sira donati yerlesimi bulunamadi."

    sc = lc / bestNc
    sx = lx / bestNx
    sy = ly / bestNy

    ' Build Row 1 coordinates
    mCount = 0
    ReDim mX(1 To nMax + 64)
    ReDim mY(1 To nMax + 64)
    BuildRow1 xOut, yOut, xIn, yIn, bestNc, bestNx, bestNy
    n1 = mCount

    ' Build Row 2 if needed
    n2 = 0
    need2 = nMax - n1
    If need2 >= 2 And nMax > n1Cap Then
        row2Off = 2# * phiM
        If row2Off < 0.05 Then row2Off = 0.05
        AddRow2Bars xOut, yOut, xIn, yIn, bestNc, bestNx, bestNy, phiM, row2Off, t, cov, need2
        n2 = mCount - n1
    End If

    actualRo = (CDbl(n1 + n2) * ab / ac) * 100#
End Sub

Private Sub BuildRow1(ByVal xOut As Double, ByVal yOut As Double, ByVal xIn As Double, ByVal yIn As Double, _
    ByVal nc As Long, ByVal nx As Long, ByVal ny As Long)
    Dim i As Long, x As Double, y As Double
    Dim lc As Double, lx As Double, ly As Double

    lc = xOut - xIn
    lx = 2# * xIn
    ly = 2# * yIn

    ' 16 fixed anchor bars:
    ' 4 outer corners
    AddBar -xOut, yOut: AddBar xOut, yOut: AddBar -xOut, -yOut: AddBar xOut, -yOut
    ' 4 inner corners
    AddBar -xIn, yIn: AddBar xIn, yIn: AddBar -xIn, -yIn: AddBar xIn, -yIn
    ' 8 projections of inner corners onto outer faces
    AddBar -xIn, yOut: AddBar xIn, yOut: AddBar -xIn, -yOut: AddBar xIn, -yOut
    AddBar -xOut, yIn: AddBar xOut, yIn: AddBar -xOut, -yIn: AddBar xOut, -yIn

    ' Corner legs: horizontal (top & bottom outer)
    For i = 1 To nc - 1
        x = xIn + lc * (CDbl(i) / CDbl(nc))
        AddBar -x, yOut: AddBar x, yOut: AddBar -x, -yOut: AddBar x, -yOut
    Next i

    ' Corner legs: vertical (left & right outer)
    For i = 1 To nc - 1
        y = yIn + lc * (CDbl(i) / CDbl(nc))
        AddBar -xOut, y: AddBar xOut, y: AddBar -xOut, -y: AddBar xOut, -y
    Next i

    ' Long middle zone: 4 horizontal lines aligned face-to-face
    For i = 1 To nx - 1
        x = -xIn + lx * (CDbl(i) / CDbl(nx))
        AddBar x, yOut: AddBar x, yIn: AddBar x, -yIn: AddBar x, -yOut
    Next i

    ' Short middle zone: 4 vertical lines aligned face-to-face
    For i = 1 To ny - 1
        y = -yIn + ly * (CDbl(i) / CDbl(ny))
        AddBar -xOut, y: AddBar -xIn, y: AddBar xIn, y: AddBar xOut, y
    Next i
End Sub

Private Sub AddRow2Bars(ByVal xOut As Double, ByVal yOut As Double, ByVal xIn As Double, ByVal yIn As Double, _
    ByVal nc As Long, ByVal nx As Long, ByVal ny As Long, ByVal phiM As Double, _
    ByVal row2Off As Double, ByVal t As Double, ByVal cov As Double, ByVal need2 As Long)
    Dim remBars As Long, limX As Double, limY As Double
    Dim xs() As Double, nXs As Long, ys() As Double, nYs As Long
    Dim innerOk As Boolean

    remBars = need2
    limX = xOut - row2Off - 0.5 * phiM
    limY = yOut - row2Off - 0.5 * phiM

    ' Collect all Row 1 X positions on outer top face (excluding corner collision zone)
    CollectOuterTopX xOut, xIn, nc, nx, limX, xs, nXs
    ' Collect all Row 1 Y positions on outer left face (excluding corner collision zone)
    CollectOuterLeftY yOut, yIn, nc, ny, limY, ys, nYs

    ' 1) Outer Top & Bottom faces (Y = +-(yOut - row2Off))
    If remBars >= 2 And nXs > 0 Then
        remBars = PlaceSymmetricHoriz(xs, nXs, yOut - row2Off, remBars, row2Off)
    End If

    ' 2) Outer Left & Right faces (X = +-(xOut - row2Off))
    If remBars >= 2 And nYs > 0 Then
        remBars = PlaceSymmetricVert(ys, nYs, xOut - row2Off, remBars, row2Off)
    End If

    ' 3) Inner faces if still needed and wall thickness allows clearance
    innerOk = (2# * (cov + row2Off) + 2# * phiM < t)
    If remBars >= 2 And innerOk Then
        CollectInnerTopX xIn, nx, xs, nXs
        If remBars >= 2 And nXs > 0 Then
            remBars = PlaceSymmetricHoriz(xs, nXs, yIn + row2Off, remBars, row2Off)
        End If
        CollectInnerLeftY yIn, ny, ys, nYs
        If remBars >= 2 And nYs > 0 Then
            remBars = PlaceSymmetricVert(ys, nYs, xIn + row2Off, remBars, row2Off)
        End If
    End If
End Sub

Private Sub CollectOuterTopX(ByVal xOut As Double, ByVal xIn As Double, ByVal nc As Long, ByVal nx As Long, _
    ByVal limX As Double, ByRef arr() As Double, ByRef cnt As Long)
    Dim i As Long, x As Double, lc As Double, lx As Double
    lc = xOut - xIn: lx = 2# * xIn
    cnt = 0
    ReDim arr(1 To 2 * nc + nx + 4)

    ' Strictly positive X values only (symmetry will mirror to -X; X=0 handled separately)
    ' Middle zone (nx is total spaces; take X >= 0)
    For i = 1 To nx - 1
        x = -xIn + lx * (CDbl(i) / CDbl(nx))
        If x >= -0.000001 And x <= limX + 0.000001 Then
            If x < 0.000001 Then x = 0#
            cnt = cnt + 1: arr(cnt) = x
        End If
    Next i
    ' Inner corner projection X = xIn
    If xIn <= limX + 0.000001 Then cnt = cnt + 1: arr(cnt) = xIn
    ' Corner leg bars
    For i = 1 To nc - 1
        x = xIn + lc * (CDbl(i) / CDbl(nc))
        If x <= limX + 0.000001 Then cnt = cnt + 1: arr(cnt) = x
    Next i
    SortAsc arr, cnt
End Sub

Private Sub CollectOuterLeftY(ByVal yOut As Double, ByVal yIn As Double, ByVal nc As Long, ByVal ny As Long, _
    ByVal limY As Double, ByRef arr() As Double, ByRef cnt As Long)
    Dim i As Long, y As Double, lc As Double, ly As Double
    lc = yOut - yIn: ly = 2# * yIn
    cnt = 0
    ReDim arr(1 To 2 * nc + ny + 4)

    For i = 1 To ny - 1
        y = -yIn + ly * (CDbl(i) / CDbl(ny))
        If y >= -0.000001 And y <= limY + 0.000001 Then
            If y < 0.000001 Then y = 0#
            cnt = cnt + 1: arr(cnt) = y
        End If
    Next i
    If yIn <= limY + 0.000001 Then cnt = cnt + 1: arr(cnt) = yIn
    For i = 1 To nc - 1
        y = yIn + lc * (CDbl(i) / CDbl(nc))
        If y <= limY + 0.000001 Then cnt = cnt + 1: arr(cnt) = y
    Next i
    SortAsc arr, cnt
End Sub

Private Sub CollectInnerTopX(ByVal xIn As Double, ByVal nx As Long, ByRef arr() As Double, ByRef cnt As Long)
    Dim i As Long, x As Double, lx As Double
    lx = 2# * xIn: cnt = 0
    ReDim arr(1 To nx + 2)
    For i = 1 To nx - 1
        x = -xIn + lx * (CDbl(i) / CDbl(nx))
        If x >= -0.000001 Then
            If x < 0.000001 Then x = 0#
            cnt = cnt + 1: arr(cnt) = x
        End If
    Next i
    SortAsc arr, cnt
End Sub

Private Sub CollectInnerLeftY(ByVal yIn As Double, ByVal ny As Long, ByRef arr() As Double, ByRef cnt As Long)
    Dim i As Long, y As Double, ly As Double
    ly = 2# * yIn: cnt = 0
    ReDim arr(1 To ny + 2)
    For i = 1 To ny - 1
        y = -yIn + ly * (CDbl(i) / CDbl(ny))
        If y >= -0.000001 Then
            If y < 0.000001 Then y = 0#
            cnt = cnt + 1: arr(cnt) = y
        End If
    Next i
    SortAsc arr, cnt
End Sub

Private Function PlaceSymmetricHoriz(ByRef posX() As Double, ByVal cnt As Long, ByVal yPos As Double, ByVal remBars As Long, ByVal dMin As Double) As Long
    Dim hasZero As Boolean, pCount As Long, pArr() As Double
    Dim i As Long, maxPairs As Long, takePairs As Long, idx As Long, used() As Boolean

    ' The layout is doubly symmetric, so checking the +X position covers all four
    hasZero = (cnt > 0 And posX(1) = 0#)
    If hasZero Then hasZero = Not TooClose(0#, yPos, dMin)
    pCount = 0
    ReDim pArr(1 To cnt + 1)
    For i = 1 To cnt
        If posX(i) > 0# Then
            If Not TooClose(posX(i), yPos, dMin) Then pCount = pCount + 1: pArr(pCount) = posX(i)
        End If
    Next i

    maxPairs = remBars \ 4
    takePairs = maxPairs
    If takePairs > pCount Then takePairs = pCount

    If takePairs > 0 Then
        ReDim used(1 To pCount)
        For i = 1 To takePairs
            If takePairs = pCount Then
                idx = i
            ElseIf takePairs = 1 Then
                idx = pCount
            Else
                idx = 1 + CLng(Round(CDbl(i - 1) * CDbl(pCount - 1) / CDbl(takePairs - 1), 0))
            End If
            If idx < 1 Then idx = 1
            If idx > pCount Then idx = pCount
            Do While used(idx) And idx < pCount: idx = idx + 1: Loop
            Do While used(idx) And idx > 1: idx = idx - 1: Loop
            used(idx) = True
            AddBar pArr(idx), yPos
            AddBar -pArr(idx), yPos
            AddBar pArr(idx), -yPos
            AddBar -pArr(idx), -yPos
            remBars = remBars - 4
        Next i
    End If

    If hasZero And remBars >= 2 Then
        AddBar 0#, yPos
        AddBar 0#, -yPos
        remBars = remBars - 2
    End If

    PlaceSymmetricHoriz = remBars
End Function

Private Function PlaceSymmetricVert(ByRef posY() As Double, ByVal cnt As Long, ByVal xPos As Double, ByVal remBars As Long, ByVal dMin As Double) As Long
    Dim hasZero As Boolean, pCount As Long, pArr() As Double
    Dim i As Long, maxPairs As Long, takePairs As Long, idx As Long, used() As Boolean

    ' The layout is doubly symmetric, so checking the +Y position covers all four
    hasZero = (cnt > 0 And posY(1) = 0#)
    If hasZero Then hasZero = Not TooClose(xPos, 0#, dMin)
    pCount = 0
    ReDim pArr(1 To cnt + 1)
    For i = 1 To cnt
        If posY(i) > 0# Then
            If Not TooClose(xPos, posY(i), dMin) Then pCount = pCount + 1: pArr(pCount) = posY(i)
        End If
    Next i

    maxPairs = remBars \ 4
    takePairs = maxPairs
    If takePairs > pCount Then takePairs = pCount

    If takePairs > 0 Then
        ReDim used(1 To pCount)
        For i = 1 To takePairs
            If takePairs = pCount Then
                idx = i
            ElseIf takePairs = 1 Then
                idx = pCount
            Else
                idx = 1 + CLng(Round(CDbl(i - 1) * CDbl(pCount - 1) / CDbl(takePairs - 1), 0))
            End If
            If idx < 1 Then idx = 1
            If idx > pCount Then idx = pCount
            Do While used(idx) And idx < pCount: idx = idx + 1: Loop
            Do While used(idx) And idx > 1: idx = idx - 1: Loop
            used(idx) = True
            AddBar xPos, pArr(idx)
            AddBar xPos, -pArr(idx)
            AddBar -xPos, pArr(idx)
            AddBar -xPos, -pArr(idx)
            remBars = remBars - 4
        Next i
    End If

    If hasZero And remBars >= 2 Then
        AddBar xPos, 0#
        AddBar -xPos, 0#
        remBars = remBars - 2
    End If

    PlaceSymmetricVert = remBars
End Function

Private Sub WriteBoxDxf(ByVal outPath As String, ByVal B As Double, ByVal H As Double, _
    ByVal t As Double, ByVal chf As Double, ByVal phiMm As Double)
    Dim f As Integer, hb As Double, hh As Double, ib As Double, ih As Double
    Dim blkScale As Double, i As Long, hNext As Long, handSeed As String

    hb = B / 2#: hh = H / 2#
    ib = hb - t: ih = hh - t
    blkScale = phiMm / 30#
    hNext = &H100&
    handSeed = Hex$(hNext + mCount + 32)

    f = FreeFile
    Open outPath For Output As #f
    On Error GoTo WriteFail

    ' AutoCAD 2007/2008 (AC1021) DXF: HEADER, CLASSES, TABLES, BLOCKS, ENTITIES, OBJECTS
    Print #f, "0": Print #f, "SECTION"
    Print #f, "2": Print #f, "HEADER"
    Print #f, "9": Print #f, "$ACADVER"
    Print #f, "1": Print #f, "AC1021"
    Print #f, "9": Print #f, "$DWGCODEPAGE"
    Print #f, "3": Print #f, "ANSI_1252"
    Print #f, "9": Print #f, "$INSBASE"
    Print #f, "10": Print #f, "0.0"
    Print #f, "20": Print #f, "0.0"
    Print #f, "30": Print #f, "0.0"
    Print #f, "9": Print #f, "$EXTMIN"
    Print #f, "10": Print #f, DxfNum(-hb)
    Print #f, "20": Print #f, DxfNum(-hh)
    Print #f, "30": Print #f, "0.0"
    Print #f, "9": Print #f, "$EXTMAX"
    Print #f, "10": Print #f, DxfNum(hb)
    Print #f, "20": Print #f, DxfNum(hh)
    Print #f, "30": Print #f, "0.0"
    Print #f, "9": Print #f, "$MEASUREMENT"
    Print #f, "70": Print #f, "1"
    Print #f, "9": Print #f, "$INSUNITS"
    Print #f, "70": Print #f, "6"
    Print #f, "9": Print #f, "$HANDSEED"
    Print #f, "5": Print #f, handSeed
    Print #f, "0": Print #f, "ENDSEC"

    Print #f, "0": Print #f, "SECTION"
    Print #f, "2": Print #f, "CLASSES"
    Print #f, "0": Print #f, "ENDSEC"

    Print #f, "0": Print #f, "SECTION"
    Print #f, "2": Print #f, "TABLES"

    WriteEmptyTable f, "VPORT", "8"

    ' LTYPE table
    StartTable f, "LTYPE", "5", 3
    WriteLtype f, "14", "5", "ByBlock", ""
    WriteLtype f, "15", "5", "ByLayer", ""
    WriteLtype f, "16", "5", "Continuous", "Solid line"
    Print #f, "0": Print #f, "ENDTAB"

    ' LAYER table
    StartTable f, "LAYER", "2", 3
    WriteLayer2021 f, "10", "2", "0", 7
    WriteLayer2021 f, "30", "2", LAY_CONC, 7
    WriteLayer2021 f, "31", "2", LAY_REBAR, 1
    Print #f, "0": Print #f, "ENDTAB"

    ' STYLE table
    StartTable f, "STYLE", "3", 1
    Print #f, "0": Print #f, "STYLE"
    Print #f, "5": Print #f, "11"
    Print #f, "330": Print #f, "3"
    Print #f, "100": Print #f, "AcDbSymbolTableRecord"
    Print #f, "100": Print #f, "AcDbTextStyleTableRecord"
    Print #f, "2": Print #f, "Standard"
    Print #f, "70": Print #f, "0"
    Print #f, "40": Print #f, "0.0"
    Print #f, "41": Print #f, "1.0"
    Print #f, "50": Print #f, "0.0"
    Print #f, "71": Print #f, "0"
    Print #f, "42": Print #f, "0.2"
    Print #f, "3": Print #f, "txt"
    Print #f, "4": Print #f, ""
    Print #f, "0": Print #f, "ENDTAB"

    WriteEmptyTable f, "VIEW", "6"
    WriteEmptyTable f, "UCS", "7"

    ' APPID table
    StartTable f, "APPID", "9", 1
    Print #f, "0": Print #f, "APPID"
    Print #f, "5": Print #f, "12"
    Print #f, "330": Print #f, "9"
    Print #f, "100": Print #f, "AcDbSymbolTableRecord"
    Print #f, "100": Print #f, "AcDbRegAppTableRecord"
    Print #f, "2": Print #f, "ACAD"
    Print #f, "70": Print #f, "0"
    Print #f, "0": Print #f, "ENDTAB"

    ' DIMSTYLE table
    Print #f, "0": Print #f, "TABLE"
    Print #f, "2": Print #f, "DIMSTYLE"
    Print #f, "5": Print #f, "A"
    Print #f, "330": Print #f, "0"
    Print #f, "100": Print #f, "AcDbSymbolTable"
    Print #f, "70": Print #f, "0"
    Print #f, "100": Print #f, "AcDbDimStyleTable"
    Print #f, "71": Print #f, "0"
    Print #f, "0": Print #f, "ENDTAB"

    ' BLOCK_RECORD table
    StartTable f, "BLOCK_RECORD", "1", 3
    WriteBlockRecord f, "1F", "*Model_Space"
    WriteBlockRecord f, "1B", "*Paper_Space"
    WriteBlockRecord f, "40", "B"
    Print #f, "0": Print #f, "ENDTAB"

    Print #f, "0": Print #f, "ENDSEC"

    ' BLOCKS section
    Print #f, "0": Print #f, "SECTION"
    Print #f, "2": Print #f, "BLOCKS"
    WriteEmptyBlock f, "20", "21", "1F", "*Model_Space"
    WriteEmptyBlock f, "1C", "1D", "1B", "*Paper_Space"

    ' Rebar donut block B (0.030 m base diameter, scaled by phiMm / 30)
    Print #f, "0": Print #f, "BLOCK"
    Print #f, "5": Print #f, "41"
    Print #f, "330": Print #f, "40"
    Print #f, "100": Print #f, "AcDbEntity"
    Print #f, "8": Print #f, "0"
    Print #f, "100": Print #f, "AcDbBlockBegin"
    Print #f, "2": Print #f, "B"
    Print #f, "70": Print #f, "0"
    Print #f, "10": Print #f, "0.0"
    Print #f, "20": Print #f, "0.0"
    Print #f, "30": Print #f, "0.0"
    Print #f, "3": Print #f, "B"
    Print #f, "1": Print #f, ""
    Print #f, "0": Print #f, "LWPOLYLINE"
    Print #f, "5": Print #f, "42"
    Print #f, "330": Print #f, "40"
    Print #f, "100": Print #f, "AcDbEntity"
    Print #f, "8": Print #f, LAY_REBAR
    Print #f, "100": Print #f, "AcDbPolyline"
    Print #f, "90": Print #f, "2"
    Print #f, "70": Print #f, "1"
    Print #f, "10": Print #f, "-0.0075"
    Print #f, "20": Print #f, "0.0"
    Print #f, "40": Print #f, "0.015"
    Print #f, "41": Print #f, "0.015"
    Print #f, "42": Print #f, "1.0"
    Print #f, "10": Print #f, "0.0075"
    Print #f, "20": Print #f, "0.0"
    Print #f, "40": Print #f, "0.015"
    Print #f, "41": Print #f, "0.015"
    Print #f, "42": Print #f, "1.0"
    Print #f, "0": Print #f, "ENDBLK"
    Print #f, "5": Print #f, "43"
    Print #f, "330": Print #f, "40"
    Print #f, "100": Print #f, "AcDbEntity"
    Print #f, "8": Print #f, "0"
    Print #f, "100": Print #f, "AcDbBlockEnd"
    Print #f, "0": Print #f, "ENDSEC"

    ' ENTITIES section (all owned by *Model_Space handle 1F)
    Print #f, "0": Print #f, "SECTION"
    Print #f, "2": Print #f, "ENTITIES"

    ' Outer concrete boundary (closed LWPOLYLINE)
    If chf > 0# Then
        StartLwPoly f, NextHandle(hNext), LAY_CONC, 8
        AddLwPt f, hb - chf, hh
        AddLwPt f, -hb + chf, hh
        AddLwPt f, -hb, hh - chf
        AddLwPt f, -hb, -hh + chf
        AddLwPt f, -hb + chf, -hh
        AddLwPt f, hb - chf, -hh
        AddLwPt f, hb, -hh + chf
        AddLwPt f, hb, hh - chf
    Else
        StartLwPoly f, NextHandle(hNext), LAY_CONC, 4
        AddLwPt f, hb, hh
        AddLwPt f, -hb, hh
        AddLwPt f, -hb, -hh
        AddLwPt f, hb, -hh
    End If

    ' Inner hollow void (closed LWPOLYLINE)
    StartLwPoly f, NextHandle(hNext), LAY_CONC, 4
    AddLwPt f, ib, -ih
    AddLwPt f, ib, ih
    AddLwPt f, -ib, ih
    AddLwPt f, -ib, -ih

    ' Longitudinal rebars (INSERT of block B on DEMIR-30)
    For i = 1 To mCount
        Print #f, "0": Print #f, "INSERT"
        Print #f, "5": Print #f, NextHandle(hNext)
        Print #f, "330": Print #f, "1F"
        Print #f, "100": Print #f, "AcDbEntity"
        Print #f, "8": Print #f, LAY_REBAR
        Print #f, "100": Print #f, "AcDbBlockReference"
        Print #f, "2": Print #f, "B"
        Print #f, "10": Print #f, DxfNum(mX(i))
        Print #f, "20": Print #f, DxfNum(mY(i))
        Print #f, "30": Print #f, "0.0"
        Print #f, "41": Print #f, DxfNum(blkScale)
        Print #f, "42": Print #f, DxfNum(blkScale)
        Print #f, "43": Print #f, "1.0"
        Print #f, "50": Print #f, "0.0"
    Next i

    Print #f, "0": Print #f, "ENDSEC"

    ' OBJECTS section (root dictionary + ACAD_GROUP + ACAD_PLOTSTYLENAME + Normal)
    Print #f, "0": Print #f, "SECTION"
    Print #f, "2": Print #f, "OBJECTS"
    Print #f, "0": Print #f, "DICTIONARY"
    Print #f, "5": Print #f, "C"
    Print #f, "330": Print #f, "0"
    Print #f, "100": Print #f, "AcDbDictionary"
    Print #f, "281": Print #f, "1"
    Print #f, "3": Print #f, "ACAD_GROUP"
    Print #f, "350": Print #f, "D"
    Print #f, "3": Print #f, "ACAD_PLOTSTYLENAME"
    Print #f, "350": Print #f, "E"

    Print #f, "0": Print #f, "DICTIONARY"
    Print #f, "5": Print #f, "D"
    Print #f, "330": Print #f, "C"
    Print #f, "100": Print #f, "AcDbDictionary"
    Print #f, "281": Print #f, "1"

    Print #f, "0": Print #f, "ACDBDICTIONARYWDFLT"
    Print #f, "5": Print #f, "E"
    Print #f, "330": Print #f, "C"
    Print #f, "100": Print #f, "AcDbDictionary"
    Print #f, "281": Print #f, "1"
    Print #f, "3": Print #f, "Normal"
    Print #f, "350": Print #f, "F"
    Print #f, "100": Print #f, "AcDbDictionaryWithDefault"
    Print #f, "340": Print #f, "F"

    Print #f, "0": Print #f, "ACDBPLACEHOLDER"
    Print #f, "5": Print #f, "F"
    Print #f, "330": Print #f, "E"

    Print #f, "0": Print #f, "ENDSEC"
    Print #f, "0": Print #f, "EOF"
    Close #f
    Exit Sub

WriteFail:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.Description
    Close #f
    Err.Raise errNum, , "DXF yazilamadi: " & errDesc
End Sub

Private Function NextHandle(ByRef hNext As Long) As String
    NextHandle = Hex$(hNext)
    hNext = hNext + 1
End Function

Private Sub StartTable(ByVal f As Integer, ByVal tblName As String, ByVal hnd As String, ByVal cnt As Long)
    Print #f, "0": Print #f, "TABLE"
    Print #f, "2": Print #f, tblName
    Print #f, "5": Print #f, hnd
    Print #f, "330": Print #f, "0"
    Print #f, "100": Print #f, "AcDbSymbolTable"
    Print #f, "70": Print #f, CStr(cnt)
End Sub

Private Sub WriteEmptyTable(ByVal f As Integer, ByVal tblName As String, ByVal hnd As String)
    StartTable f, tblName, hnd, 0
    Print #f, "0": Print #f, "ENDTAB"
End Sub

Private Sub WriteLtype(ByVal f As Integer, ByVal hnd As String, ByVal ownerHnd As String, _
    ByVal ltName As String, ByVal desc As String)
    Print #f, "0": Print #f, "LTYPE"
    Print #f, "5": Print #f, hnd
    Print #f, "330": Print #f, ownerHnd
    Print #f, "100": Print #f, "AcDbSymbolTableRecord"
    Print #f, "100": Print #f, "AcDbLinetypeTableRecord"
    Print #f, "2": Print #f, ltName
    Print #f, "70": Print #f, "0"
    Print #f, "3": Print #f, desc
    Print #f, "72": Print #f, "65"
    Print #f, "73": Print #f, "0"
    Print #f, "40": Print #f, "0.0"
End Sub

Private Sub WriteLayer2021(ByVal f As Integer, ByVal hnd As String, ByVal ownerHnd As String, _
    ByVal layName As String, ByVal colorIdx As Long)
    Print #f, "0": Print #f, "LAYER"
    Print #f, "5": Print #f, hnd
    Print #f, "330": Print #f, ownerHnd
    Print #f, "100": Print #f, "AcDbSymbolTableRecord"
    Print #f, "100": Print #f, "AcDbLayerTableRecord"
    Print #f, "2": Print #f, layName
    Print #f, "70": Print #f, "0"
    Print #f, "62": Print #f, CStr(colorIdx)
    Print #f, "6": Print #f, "Continuous"
    Print #f, "370": Print #f, "-3"
    Print #f, "390": Print #f, "F"
End Sub

Private Sub WriteBlockRecord(ByVal f As Integer, ByVal hnd As String, ByVal blkName As String)
    Print #f, "0": Print #f, "BLOCK_RECORD"
    Print #f, "5": Print #f, hnd
    Print #f, "330": Print #f, "1"
    Print #f, "100": Print #f, "AcDbSymbolTableRecord"
    Print #f, "100": Print #f, "AcDbBlockTableRecord"
    Print #f, "2": Print #f, blkName
    Print #f, "70": Print #f, "0"
    Print #f, "280": Print #f, "1"
    Print #f, "281": Print #f, "0"
End Sub

Private Sub WriteEmptyBlock(ByVal f As Integer, ByVal hStart As String, ByVal hEnd As String, _
    ByVal ownerHnd As String, ByVal blkName As String)
    Print #f, "0": Print #f, "BLOCK"
    Print #f, "5": Print #f, hStart
    Print #f, "330": Print #f, ownerHnd
    Print #f, "100": Print #f, "AcDbEntity"
    Print #f, "8": Print #f, "0"
    Print #f, "100": Print #f, "AcDbBlockBegin"
    Print #f, "2": Print #f, blkName
    Print #f, "70": Print #f, "0"
    Print #f, "10": Print #f, "0.0"
    Print #f, "20": Print #f, "0.0"
    Print #f, "30": Print #f, "0.0"
    Print #f, "3": Print #f, blkName
    Print #f, "1": Print #f, ""
    Print #f, "0": Print #f, "ENDBLK"
    Print #f, "5": Print #f, hEnd
    Print #f, "330": Print #f, ownerHnd
    Print #f, "100": Print #f, "AcDbEntity"
    Print #f, "8": Print #f, "0"
    Print #f, "100": Print #f, "AcDbBlockEnd"
End Sub

Private Sub StartLwPoly(ByVal f As Integer, ByVal hnd As String, ByVal layName As String, ByVal nPts As Long)
    Print #f, "0": Print #f, "LWPOLYLINE"
    Print #f, "5": Print #f, hnd
    Print #f, "330": Print #f, "1F"
    Print #f, "100": Print #f, "AcDbEntity"
    Print #f, "8": Print #f, layName
    Print #f, "100": Print #f, "AcDbPolyline"
    Print #f, "90": Print #f, CStr(nPts)
    Print #f, "70": Print #f, "1"
    Print #f, "43": Print #f, "0.0"
End Sub

Private Sub AddLwPt(ByVal f As Integer, ByVal x As Double, ByVal y As Double)
    Print #f, "10": Print #f, DxfNum(x)
    Print #f, "20": Print #f, DxfNum(y)
End Sub

Private Sub AddBar(ByVal x As Double, ByVal y As Double)
    mCount = mCount + 1
    If mCount > UBound(mX) Then
        ReDim Preserve mX(1 To mCount + 64)
        ReDim Preserve mY(1 To mCount + 64)
    End If
    mX(mCount) = x
    mY(mCount) = y
End Sub

' True when (x, y) lies closer than dMin to a bar already placed (centre to centre)
Private Function TooClose(ByVal x As Double, ByVal y As Double, ByVal dMin As Double) As Boolean
    Dim i As Long, dx As Double, dy As Double
    For i = 1 To mCount
        dx = mX(i) - x: dy = mY(i) - y
        If dx * dx + dy * dy < dMin * dMin - 0.0000000001 Then TooClose = True: Exit Function
    Next i
End Function

Private Sub SortAsc(ByRef arr() As Double, ByVal cnt As Long)
    Dim i As Long, j As Long, tmp As Double
    For i = 1 To cnt - 1
        For j = i + 1 To cnt
            If arr(j) < arr(i) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

' one number through the PANEL question box; False on Cancel (UiInput asks again on a wrong number)
Private Function AskDouble(ByVal prompt As String, ByVal defVal As Double, ByRef outVal As Double) As Boolean
    Dim v As Variant
    v = UiInput(prompt, MSG_TITLE, DxfNum(defVal), True)
    If VarType(v) = vbBoolean Then Exit Function
    outVal = CDbl(v)
    AskDouble = True
End Function

Private Function PickFolder() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "DXF kayit klasorunu secin"
    If fd.Show = -1 Then PickFolder = fd.SelectedItems(1)
End Function

Private Function DxfNum(ByVal x As Double) As String
    Dim s As String
    If Abs(x) < 0.0000000001 Then DxfNum = "0.0": Exit Function
    s = Trim$(Str$(Round(x, 6)))
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    DxfNum = s
End Function

Private Function DotFixed(ByVal x As Double, ByVal dec As Long) As String
    Dim fmt As String
    If dec <= 0 Then
        DotFixed = Trim$(Str$(Round(x, 0)))
    Else
        fmt = "0." & String$(dec, "0")
        DotFixed = Replace(Format$(x, fmt), ",", ".")
    End If
End Function

Private Function Max3(ByVal a As Double, ByVal B As Double, ByVal c As Double) As Double
    Max3 = a
    If B > Max3 Then Max3 = B
    If c > Max3 Then Max3 = c
End Function

Private Function Min3(ByVal a As Double, ByVal B As Double, ByVal c As Double) As Double
    Min3 = a
    If B < Min3 Then Min3 = B
    If c < Min3 Then Min3 = c
End Function
