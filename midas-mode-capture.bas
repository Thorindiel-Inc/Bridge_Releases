Attribute VB_Name = "MIDAS_SEKIL"
Option Explicit
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal h As LongPtr, ByVal c As Long) As LongPtr
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowTextW Lib "user32" (ByVal h As LongPtr, ByVal s As LongPtr, ByVal n As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal h As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal h As LongPtr, ByVal n As Long) As Long
' ============================================================================
'  MIDAS_SEKIL - model and mode-shape pictures of the seismic model into sheet Z_YATAY
'  (chapter 8 workbook)
'
'  Macro window (MAKROLAR) > MIDAS sonuclari > Modal analiz sonuclari (MIDAS_MOD_OKU, field Islem:
'  read + pictures or pictures only), which calls MIDAS_SEKIL_CAPTURE; or MIDAS_SEKIL_AL alone.
'  Reads the model OPEN in Civil NX through the MAPI; it writes nothing to the model,
'  it only changes the view (view/CAPTURE) and reads db/PJCF and db/NODE.
'  Pages of Z_YATAY, found by their caption row (C29, C61 ...):
'     page 1  model, line view with supports (elastic link symbols would cover a long bridge)
'     page 2  model, solid (hidden) view
'     page 3-6  vibration mode shapes, solid view, contour, no legend, no undeformed shape
'  Modes: the 2 largest TRAN-X and the 2 largest TRAN-Y single-mode mass shares of sheet
'  MOD (written by MIDAS_MOD); a mode already taken goes to the next one of that
'  direction; the four are shown in mode-number order. The caption text of a mode page
'  (column AS) becomes "Mod Sekli - n".
'  Each picture replaces the one on its page at the same place and size (the old one is
'  deleted only after the new one is in). Images go to %TEMP%\Bridge_Z_YATAY.
'  Camera: one 3D view, turned to the bridge direction (principal axis of the node plan). Zoom so
'  the bridge fills FILL of the picture: Civil NX 'zoom fit' (ZOOM_LEVEL 100) shows a height of
'  K_FIT x the 3D diagonal of the model, and each zoom step scales by exp(ZOOM_RATE) (both measured
'  2026-11-09, zoom 60-105); the projected size of the nodes for this camera gives the zoom.
'  Each image is cropped (centred) to the shape of its frame, and the zoom is chosen for the kept
'  part. At screen size the image shape is only known after the first capture: if it needs
'  another zoom, that page is captured once more.
'  Resolution (field 'res' of the macro window): 1 = screen size, no WIDTH / HEIGHT sent, the image
'  is the Civil NX view and its window stays as it is; 2-5 = 4000 / 3000 / 2000 / 1000 px wide
'  (height from the frame of the page), but Civil NX resizes its window for each capture (measured
'  2026-11-09), so the macro maximizes it again after each one. On a long deck small pictures get
'  dark: below about 3500 px wide the mesh lines fill the deck, so 2000 and 1000 px turn
'  the deck black. Each picture is fitted into the frame of the old one (cropped to its shape).
'  MAPI key: INPUT!T3 (G_MAPI_KEY), else MIDAS_MAPI_KEY, else asked.
'  Test entry without Civil NX: MIDAS_SEKIL_PLAN (modes, target pictures, request bodies).
' ============================================================================

Private Const SCRIPT_ID As String = "midas-mode-capture"
Private Const SCRIPT_VERSION As String = "2026-11-09b"
Private Const SCRIPT_CHANGELOG As String = "Scripted entry MIDAS_SEKIL_CAPTURE for the combined tool Modal analiz sonuclari (MIDAS_MOD_OKU: read + pictures, pictures only, read only); new: warm-up capture where the view switches between model and results (Civil NX drops the zoom there); zoom from the model so the bridge fills the picture; picture size choice (screen size, 4000, 3000, 2000 or 1000 px wide); model and mode-shape pictures of the seismic model into Z_YATAY (2 largest X + 2 largest Y modes from MOD), captions follow the modes"

Private Const API_BASE As String = "https://moa-engineers.midasit.com:443/civil/"
Private Const MSG_TITLE As String = "MIDAS_SEKIL"
Private Const KEY_SHEET As String = "INPUT"
Private Const KEY_CELL As String = "T3"
Private Const KEY_NAME As String = "G_MAPI_KEY"
Private Const FIG_SHEET As String = "Z_YATAY"
Private Const MOD_SHEET As String = "MOD"
Private Const R_PCT As Long = 109               ' MOD: mode C, TRAN-X mass % D, TRAN-Y mass % F
Private Const N_MAX As Long = 100
Private Const CAP_COL As String = "AS"          ' caption text of each page
Private Const PAGE_ROWS As String = "29,61,93,125,157,189"   ' caption rows, page 1..6
Private Const N_MODEL As Long = 2               ' pages 1-2 model views, the rest mode shapes
Private Const N_X As Long = 2, N_Y As Long = 2
' view (measured 2026-11-09 on the seismic model)
Private Const CAM_H As Double = 30              ' ANGLE.HORIZONTAL for a bridge along +X
Private Const CAM_V As Double = 30              ' ANGLE.VERTICAL (elevation)
Private Const MODE_SCALE As Double = 0.3        ' mode shape scale factor (Civil NX default 1.0; owner: lower)
Private Const SHOW_UNDEFORMED As Boolean = False ' only the object form {"OPT_CHECK":true} is accepted
Private Const N_COLOR As Long = 12
Private Const WAIT_S As Long = 30               ' seconds to wait for each image file
Private Const FILL As Double = 0.92             ' share of the picture the bridge should fill
Private Const K_FIT As Double = 0.627           ' visible height at zoom 100 / 3D diagonal of the model
Private Const ZOOM_RATE As Double = 0.0203      ' ln(scale) per zoom step
Private Const RES_W As String = "0,4000,3000,2000,1000" ' picture width per choice of field 'res' (0 = screen)

Private mKey As String
Private mResW As Long                           ' 0 = screen size
Private mWm As Double, mHm As Double, mD As Double ' projected width, height, 3D diagonal (m)
Private mAspect As Double                       ' picture width / height at screen size, 0 = not known yet
Private mLastMode As String                     ' SET_MODE of the last capture ("" at the start of a run)
Private http As Object

' ============================================================================
Public Sub MIDAS_SEKIL_AL()
    Dim pj As String, msg As String, modes As Variant, ok As Boolean
    On Error GoTo Fail
    If Not LayoutOk(msg) Then UiMsg msg, vbExclamation, MSG_TITLE: Exit Sub
    modes = PickModes(msg)
    If IsEmpty(modes) Then UiMsg msg, vbExclamation, MSG_TITLE: Exit Sub
    If Not GetKey() Then UiMsg "No MAPI key: nothing was captured.", vbExclamation, MSG_TITLE: Exit Sub
    mResW = ResWidth(UiIdx("res", 1))

    UiLog "Civil NX: model check"
    If Not Api("GET", "db/PJCF", "", pj) Then
        If Left$(pj, 10) = "HTTP error" Or pj = "" Then
            UiMsg "Civil NX does not answer (is it open, is the MAPI connected?)." & vbCrLf & Left$(pj, 200), vbExclamation, MSG_TITLE
            Exit Sub
        End If
        pj = ""
    End If
    If InStr(1, pj, " SEISMIC", vbBinaryCompare) = 0 Or InStr(1, pj, "midas-bridge-build", vbBinaryCompare) = 0 Then
        If UiMsg("The model open in Civil NX is not a MIDAS_KOPRU seismic model." & vbCrLf & vbCrLf & _
                 "Take its pictures anyway?", vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) <> vbYes Then Exit Sub
    End If
    If Not ModMatches(pj, msg) Then
        If UiMsg(msg & vbCrLf & vbCrLf & "Take the pictures with these modes anyway?", vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) <> vbYes Then Exit Sub
    End If
    If UiMsg("Civil NX will show and capture 6 views (modes " & JoinModes(modes) & ")." & vbCrLf & _
             IIf(mResW > 0, "Pictures " & mResW & " px wide: the Civil NX window changes size for each capture and is maximized again after it.", "Pictures at the size of the Civil NX view: its window is not changed.") & vbCrLf & vbCrLf & _
             "Continue?", vbOKCancel + vbQuestion, MSG_TITLE) <> vbOK Then Exit Sub

    msg = CaptureAll(modes, ok)
    UiMsg msg, IIf(ok, vbInformation, vbExclamation), MSG_TITLE
    Exit Sub
Fail:
    MaximizeCivil
    UiMsg "STOPPED: " & Err.Description, vbExclamation, MSG_TITLE
End Sub

' scripted capture for the combined tool MIDAS_MOD_OKU (called late bound with Application.Run):
'   pxWidth        picture width in px, 0 = screen size
'   ask            True: questions in the window when the model or the MOD stamp do not fit;
'                  False: such a run stops with a STOPPED line
'   startQuestion  True: the 'Civil NX will show and capture 6 views' question (pictures only)
'   modelChecked   True: the caller has already accepted the open model (read + pictures)
' returns the summary text; it never raises
Public Function MIDAS_SEKIL_CAPTURE(ByVal pxWidth As Long, ByVal ask As Boolean, Optional ByVal startQuestion As Boolean = False, Optional ByVal modelChecked As Boolean = False) As String
    Dim pj As String, msg As String, modes As Variant, ok As Boolean
    On Error GoTo Fail
    If Not LayoutOk(msg) Then MIDAS_SEKIL_CAPTURE = msg: Exit Function
    modes = PickModes(msg)
    If IsEmpty(modes) Then MIDAS_SEKIL_CAPTURE = msg: Exit Function
    If Not GetKey() Then MIDAS_SEKIL_CAPTURE = "STOPPED: no MAPI key: nothing was captured.": Exit Function
    mResW = pxWidth
    UiLog "Civil NX: model check (pictures)"
    If Not Api("GET", "db/PJCF", "", pj) Then
        If Left$(pj, 10) = "HTTP error" Or pj = "" Then MIDAS_SEKIL_CAPTURE = "STOPPED: Civil NX does not answer (is it open, is the MAPI connected?). " & Left$(pj, 200): Exit Function
        pj = ""
    End If
    If Not modelChecked Then
        If InStr(1, pj, " SEISMIC", vbBinaryCompare) = 0 Or InStr(1, pj, "midas-bridge-build", vbBinaryCompare) = 0 Then
            If Not ask Then MIDAS_SEKIL_CAPTURE = "STOPPED: the model open in Civil NX is not a MIDAS_KOPRU seismic model.": Exit Function
            If UiMsg("The model open in Civil NX is not a MIDAS_KOPRU seismic model." & vbCrLf & vbCrLf & _
                     "Take its pictures anyway?", vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) <> vbYes Then MIDAS_SEKIL_CAPTURE = "Stopped: nothing was captured.": Exit Function
        End If
    End If
    If Not ModMatches(pj, msg) Then
        If Not ask Then MIDAS_SEKIL_CAPTURE = "STOPPED: " & msg: Exit Function
        If UiMsg(msg & vbCrLf & vbCrLf & "Take the pictures with these modes anyway?", vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) <> vbYes Then MIDAS_SEKIL_CAPTURE = "Stopped: nothing was captured.": Exit Function
    End If
    If startQuestion Then
        If UiMsg("Civil NX will show and capture 6 views (modes " & JoinModes(modes) & ")." & vbCrLf & _
                 IIf(mResW > 0, "Pictures " & mResW & " px wide: the Civil NX window changes size for each capture and is maximized again after it.", "Pictures at the size of the Civil NX view: its window is not changed.") & vbCrLf & vbCrLf & _
                 "Continue?", vbOKCancel + vbQuestion, MSG_TITLE) <> vbOK Then MIDAS_SEKIL_CAPTURE = "Stopped: nothing was captured.": Exit Function
    End If
    MIDAS_SEKIL_CAPTURE = CaptureAll(modes, ok)
    Exit Function
Fail:
    MaximizeCivil
    MIDAS_SEKIL_CAPTURE = "STOPPED: " & Err.Description
End Function

' test entry without dialogs: live captures on the open model, only when it is a
' MIDAS_KOPRU seismic model and MOD was read from it; returns the summary
Public Function MIDAS_SEKIL_TEST(Optional ByVal resWidth As Long = 0) As String
    Dim pj As String, msg As String, modes As Variant, ok As Boolean
    On Error GoTo Fail
    If Not LayoutOk(msg) Then MIDAS_SEKIL_TEST = msg: Exit Function
    modes = PickModes(msg)
    If IsEmpty(modes) Then MIDAS_SEKIL_TEST = msg: Exit Function
    mKey = ""
    On Error Resume Next
    mKey = Trim$(CStr(ThisWorkbook.Names(KEY_NAME).RefersToRange.Cells(1, 1).Value))
    On Error GoTo Fail
    If mKey = "" Then mKey = Environ$("MIDAS_MAPI_KEY")
    If mKey = "" Then MIDAS_SEKIL_TEST = "STOPPED: no MAPI key": Exit Function
    If Not Api("GET", "db/PJCF", "", pj) Then MIDAS_SEKIL_TEST = "STOPPED: Civil NX: " & Left$(pj, 200): Exit Function
    If InStr(1, pj, " SEISMIC", vbBinaryCompare) = 0 Then MIDAS_SEKIL_TEST = "STOPPED: not a seismic model": Exit Function
    If Not ModMatches(pj, msg) Then MIDAS_SEKIL_TEST = "STOPPED: " & msg: Exit Function
    mResW = resWidth
    MIDAS_SEKIL_TEST = CaptureAll(modes, ok)
    Exit Function
Fail:
    MaximizeCivil
    MIDAS_SEKIL_TEST = "STOPPED: " & Err.Description
End Function

' ---- the six pages; Civil NX is maximized again at the end ----
Private Function CaptureAll(modes As Variant, ByRef ok As Boolean) As String
    Dim rows As Variant, i As Long, res As String, nOk As Long, report As String, folder As String, axisDeg As Double
    UiLog "Civil NX: bridge direction"
    axisDeg = BridgeAxis()
    mAspect = 0
    mLastMode = ""
    folder = Environ$("TEMP") & "\Bridge_Z_YATAY"
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
    rows = Split(PAGE_ROWS, ",")
    For i = 0 To UBound(rows)
        UiProgress 100 * i / (UBound(rows) + 1), "Page " & (i + 1)
        res = DoPage(i + 1, CLng(rows(i)), modes, axisDeg, folder)
        If res = "" Then
            nOk = nOk + 1
            If i >= N_MODEL Then SetCaption CLng(rows(i)), CLng(modes(i + 1 - N_MODEL))
        Else
            report = report & vbCrLf & "  ! page " & (i + 1) & ": " & res
        End If
    Next i
    UiProgress 100, "Done"
    MaximizeCivil
    ThisWorkbook.Worksheets(FIG_SHEET).Calculate
    ok = (nOk = UBound(rows) + 1)
    CaptureAll = IIf(ok, "Done: ", "Finished with problems: ") & nOk & " of " & (UBound(rows) + 1) & " pictures replaced." & vbCrLf & _
          "  Modes: " & JoinModes(modes) & ModeLines(modes) & vbCrLf & _
          "  Bridge direction " & Format$(axisDeg, "0.0") & " deg from +X, picture " & IIf(mResW > 0, mResW & " px wide", "at screen size") & report
End Function

' offline test: modes, target pictures and the request bodies, nothing is sent
Public Function MIDAS_SEKIL_PLAN() As String
    Dim msg As String, modes As Variant, rows As Variant, i As Long, s As String, shp As Shape
    On Error GoTo Fail
    If Not LayoutOk(msg) Then MIDAS_SEKIL_PLAN = msg: Exit Function
    modes = PickModes(msg)
    If IsEmpty(modes) Then MIDAS_SEKIL_PLAN = msg: Exit Function
    s = "Modes: " & JoinModes(modes) & ModeLines(modes)
    rows = Split(PAGE_ROWS, ",")
    For i = 0 To UBound(rows)
        Set shp = PagePicture(CLng(rows(i)))
        s = s & vbCrLf & "Page " & (i + 1) & " (caption row " & rows(i) & "): "
        If shp Is Nothing Then
            s = s & "no picture"
        Else
            s = s & shp.Name & " at " & shp.TopLeftCell.Address(False, False) & ", " & Format$(shp.Width, "0") & " x " & Format$(shp.Height, "0") & " pt" & vbCrLf & _
                "  " & CaptureBody(i + 1, modes, 0.8, "C:\\T\\p" & (i + 1) & ".png", 0, 0) & vbCrLf & "  sizes 4000 / 3000 / 2000 / 1000 px wide: HEIGHT " & ResH(shp, 4000) & " / " & ResH(shp, 3000) & " / " & ResH(shp, 2000) & " / " & ResH(shp, 1000)
        End If
    Next i
    MIDAS_SEKIL_PLAN = s
    Exit Function
Fail:
    MIDAS_SEKIL_PLAN = "STOPPED: " & Err.Description
End Function

' ---- one page: capture, wait for the file, replace the picture; "" = done ----
Private Function DoPage(ByVal page As Long, ByVal capRow As Long, modes As Variant, ByVal axisDeg As Double, ByVal folder As String) As String
    Dim shp As Shape, f As String, body As String, resp As String, w As Long, h As Long, ok As Boolean
    Dim asp As Double, zoom As Long, z2 As Long, pw As Long, ph As Long, pass As Long, pageMode As String
    Set shp = PagePicture(capRow)
    If shp Is Nothing Then DoPage = "FAIL: no picture on the page (rows " & capRow - 27 & "-" & capRow - 1 & "), nothing captured": Exit Function
    f = folder & "\Z_YATAY_" & page & ".png"
    If mResW > 0 Then
        w = mResW: h = ResH(shp, mResW): asp = w / h
    ElseIf mAspect > 0 Then
        asp = mAspect
    Else
        asp = shp.Width / shp.Height
    End If
    zoom = ZoomFor(asp, shp.Width / shp.Height)
    If page > N_MODEL Then UiLog "Civil NX: page " & page & ", mode " & modes(page - N_MODEL) & ", zoom " & zoom Else UiLog "Civil NX: page " & page & ", model, zoom " & zoom
    ' Civil NX drops ZOOM_LEVEL on the capture that switches between model and result view (measured
    ' 2026-11-09: the first post capture after a pre one came out at zoom 100): that capture is sent
    ' once as a warm-up and then again for the picture
    pageMode = IIf(page > N_MODEL, "post", "pre")
    If pageMode <> mLastMode Then
        body = CaptureBody(page, modes, axisDeg, Replace(f, "\", "\\"), w, h, zoom)
        If Not Api("POST", "view/CAPTURE", body, resp) Then DoPage = "FAIL: Civil NX: " & Left$(resp, 200): Exit Function
        ok = WaitForFile(f, WAIT_S)
        If mResW > 0 Then MaximizeCivil
        mLastMode = pageMode
    End If
    For pass = 1 To 2
        If Len(Dir(f)) > 0 Then
            On Error Resume Next
            Kill f
            On Error GoTo 0
            If Len(Dir(f)) > 0 Then DoPage = "FAIL: cannot delete the old image " & f: Exit Function
        End If
        body = CaptureBody(page, modes, axisDeg, Replace(f, "\", "\\"), w, h, zoom)
        If Not Api("POST", "view/CAPTURE", body, resp) Then DoPage = "FAIL: Civil NX: " & Left$(resp, 200): Exit Function
        ok = WaitForFile(f, WAIT_S)
        If mResW > 0 Then MaximizeCivil
        If Not ok Then DoPage = "FAIL: Civil NX accepted the request, but no image appeared": Exit Function
        If mResW > 0 Or mAspect > 0 Then Exit For
        ' screen size: the first image tells the picture shape
        If Not PngSize(f, pw, ph) Then Exit For
        mAspect = pw / ph
        z2 = ZoomFor(mAspect, shp.Width / shp.Height)
        If z2 = zoom Then Exit For
        UiLog "Civil NX: picture " & pw & " x " & ph & " px, page " & page & " again with zoom " & z2
        zoom = z2
    Next pass
    DoPage = ReplacePicture(shp, f)
End Function

' ---- request body of a page ----
Private Function CaptureBody(ByVal page As Long, modes As Variant, ByVal axisDeg As Double, ByVal pathJson As String, ByVal w As Long, ByVal h As Long, Optional ByVal zoom As Long = 100) As String
    Dim b As String, hor As Double
    hor = CAM_H - axisDeg
    hor = hor - 360 * Int(hor / 360)
    b = "{""Argument"":{"
    b = b & """EXPORT_PATH"":""" & pathJson & ""","
    If w > 0 And h > 0 Then b = b & """WIDTH"":" & w & ",""HEIGHT"":" & h & ","
    b = b & """ZOOM_LEVEL"":" & zoom & ","
    b = b & """ANGLE"":{""HORIZONTAL"":" & JNum(hor) & ",""VERTICAL"":" & JNum(CAM_V) & "},"
    If page = 1 Then
        ' elastic link symbols cover a long bridge (measured 2026-11-09), so supports only
        b = b & """SET_MODE"":""pre"",""SET_HIDDEN"":false,""DISPLAY"":{""BOUNDARY"":{""SUPPORT"":true,""ELASTIC_LINK"":false}}"
    ElseIf page <= N_MODEL Then
        ' supports off again here, so the mode-shape pages that follow are clean too
        b = b & """SET_MODE"":""pre"",""SET_HIDDEN"":true,""DISPLAY"":{""BOUNDARY"":{""SUPPORT"":false,""ELASTIC_LINK"":false}}"
    Else
        b = b & """SET_MODE"":""post"",""SET_HIDDEN"":true,""RESULT_GRAPHIC"":{"
        b = b & """CURRENT_MODE"":""VibrationModeShapes"",""LOAD_CASE_COMB"":{""NAME"":""Mode " & modes(page - N_MODEL) & """},"
        b = b & """COMPONENTS"":{""COMP"":""Md-XYZ""},"
        b = b & """TYPE_OF_DISPLAY"":{""MODE_SHAPE"":{""OPT_CHECK"":true,""SCALE_FACTOR"":" & JNum(MODE_SCALE) & "},"
        b = b & """CONTOUR"":{""OPT_CHECK"":true,""NUM_OF_COLOR"":" & N_COLOR & ",""COLOR_TYPE"":""rgb"",""OPTIONS"":{""CONTOUR_FILL"":true,""GRADIENT_FILL"":false}},"
        If SHOW_UNDEFORMED Then b = b & """UNDEFORMED"":{""OPT_CHECK"":true},"
        b = b & """LEGEND"":{""OPT_CHECK"":false}}}"
    End If
    CaptureBody = b & "}}"
End Function

' ---- modes from MOD: 2 largest X, then 2 largest Y not yet taken; sorted by mode number ----
Private Function PickModes(ByRef msg As String) As Variant
    Dim ws As Worksheet, v As Variant, n As Long, i As Long, k As Long, taken As String
    Dim pick() As Long, np As Long, best As Long, dirCol As Variant, d As Long, t As Long, j As Long
    Set ws = ThisWorkbook.Worksheets(MOD_SHEET)
    v = ws.Range(ws.Cells(R_PCT, 3), ws.Cells(R_PCT + N_MAX - 1, 6)).Value
    For i = 1 To N_MAX
        If IsEmpty(v(i, 1)) Or Not IsNumeric(v(i, 1)) Then Exit For
        n = i
    Next i
    If n < N_X + N_Y Then msg = "STOPPED: sheet MOD holds " & n & " modes (run 'Modal analiz sonuclari' first). Nothing was captured.": Exit Function
    ReDim pick(1 To N_X + N_Y)
    taken = ","
    dirCol = Array(2, 4)                         ' X mass % in D, Y mass % in F (columns 2 and 4 of C:F)
    For d = 0 To 1
        For k = 1 To IIf(d = 0, N_X, N_Y)
            best = 0
            For i = 1 To n
                If InStr(taken, "," & CLng(v(i, 1)) & ",") = 0 Then
                    If best = 0 Then
                        best = i
                    ElseIf CDbl(v(i, dirCol(d))) > CDbl(v(best, dirCol(d))) Then
                        best = i
                    End If
                End If
            Next i
            np = np + 1
            pick(np) = CLng(v(best, 1))
            taken = taken & pick(np) & ","
        Next k
    Next d
    For i = 1 To np - 1                          ' mode-number order
        For j = i + 1 To np
            If pick(j) < pick(i) Then t = pick(i): pick(i) = pick(j): pick(j) = t
        Next j
    Next i
    PickModes = pick
End Function

Private Function ModeLines(modes As Variant) As String
    Dim ws As Worksheet, i As Long, r As Long, s As String
    Set ws = ThisWorkbook.Worksheets(MOD_SHEET)
    For i = LBound(modes) To UBound(modes)
        r = R_PCT + modes(i) - 1
        If CLng(ws.Cells(r, 3).Value) = modes(i) Then
            s = s & vbCrLf & "    mode " & modes(i) & ": X " & Format$(ws.Cells(r, 4).Value, "0.00") & " %, Y " & Format$(ws.Cells(r, 6).Value, "0.00") & " %"
        End If
    Next i
    ModeLines = s
End Function

Private Function JoinModes(modes As Variant) As String
    Dim i As Long, s As String
    For i = LBound(modes) To UBound(modes)
        s = s & IIf(s = "", "", ", ") & modes(i)
    Next i
    JoinModes = s
End Function

' ---- MOD stamp (Q1, written by MIDAS_MOD) against the open model ----
Private Function ModMatches(ByVal pj As String, ByRef msg As String) As Boolean
    Dim stamp As String, tag As String
    stamp = CStr(ThisWorkbook.Worksheets(MOD_SHEET).Range("Q1").Value)
    tag = ModelTag(pj)
    If InStr(1, stamp, tag, vbTextCompare) > 0 Then ModMatches = True: Exit Function
    msg = "Sheet MOD was not read from the model open in Civil NX:" & vbCrLf & _
          "  MOD: " & IIf(stamp = "", "(no stamp)", stamp) & vbCrLf & "  Civil NX: " & tag & vbCrLf & _
          "Run 'Modal analiz sonuclari' first, so the pictures and the tables show the same modes."
End Function

' ---- the picture of a page: the largest picture whose top-left cell is above the caption ----
Private Function PagePicture(ByVal capRow As Long) As Shape
    Dim shp As Shape, r As Long, best As Shape
    For Each shp In ThisWorkbook.Worksheets(FIG_SHEET).Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            r = shp.TopLeftCell.Row
            If r >= capRow - 27 And r < capRow Then
                If best Is Nothing Then
                    Set best = shp
                ElseIf shp.Width * shp.Height > best.Width * best.Height Then
                    Set best = shp
                End If
            End If
        End If
    Next shp
    Set PagePicture = best
End Function

Private Function ResH(shp As Shape, ByVal w As Long) As Long
    ResH = CLng(w * shp.Height / shp.Width)
End Function

Private Function ResWidth(ByVal choice As Long) As Long
    Dim v As Variant
    v = Split(RES_W, ",")
    If choice >= 1 And choice <= UBound(v) + 1 Then ResWidth = CLng(v(choice - 1))
End Function

Private Function ReplacePicture(oldShp As Shape, ByVal f As String) As String
    Dim ws As Worksheet, newShp As Shape, nm As String
    Set ws = oldShp.Parent
    nm = oldShp.Name
    On Error Resume Next
    Set newShp = ws.Shapes.AddPicture(f, msoFalse, msoTrue, oldShp.Left, oldShp.Top, -1, -1)
    If Err.Number <> 0 Or newShp Is Nothing Then
        ReplacePicture = "FAIL: captured, but the picture could not be inserted (" & Err.Description & "); the old one was kept"
        Exit Function
    End If
    On Error GoTo 0
    ' crop the image to the shape of the frame (centred), then give it the frame's size
    Dim w0 As Double, h0 As Double, fa As Double, c As Double
    w0 = newShp.Width: h0 = newShp.Height: fa = oldShp.Width / oldShp.Height
    If w0 / h0 > fa Then
        c = (w0 - h0 * fa) / 2
        newShp.PictureFormat.CropLeft = c
        newShp.PictureFormat.CropRight = c
    Else
        c = (h0 - w0 / fa) / 2
        newShp.PictureFormat.CropTop = c
        newShp.PictureFormat.CropBottom = c
    End If
    newShp.LockAspectRatio = msoFalse
    newShp.Width = oldShp.Width
    newShp.Height = oldShp.Height
    newShp.Left = oldShp.Left
    newShp.Top = oldShp.Top
    newShp.LockAspectRatio = msoTrue
    newShp.Placement = oldShp.Placement
    oldShp.Delete
    newShp.Name = nm
End Function

Private Sub SetCaption(ByVal capRow As Long, ByVal mode As Long)
    ThisWorkbook.Worksheets(FIG_SHEET).Range(CAP_COL & capRow).Value = "Mod " & ChrW(350) & "ekli - " & mode
End Sub

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

' ---- bridge direction: principal axis of the node plan, degrees from +X ----
Private Function BridgeAxis() As Double
    Dim resp As String, re As Object, m As Object, mc As Object, n As Long, i As Long
    Dim x() As Double, y() As Double, z() As Double
    Dim sx As Double, sy As Double, sxx As Double, syy As Double, sxy As Double, mx As Double, my As Double
    mWm = 0: mHm = 0: mD = 0
    If Not Api("GET", "db/NODE", "", resp) Then Exit Function      ' 0 = along +X, zoom 100
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = """X""\s*:\s*(-?[0-9.eE+-]+)\s*,\s*""Y""\s*:\s*(-?[0-9.eE+-]+)\s*,\s*""Z""\s*:\s*(-?[0-9.eE+-]+)"
    Set mc = re.Execute(resp)
    n = mc.Count
    If n < 2 Then Exit Function
    ReDim x(1 To n): ReDim y(1 To n): ReDim z(1 To n)
    For Each m In mc
        i = i + 1
        x(i) = Val(m.SubMatches(0)): y(i) = Val(m.SubMatches(1)): z(i) = Val(m.SubMatches(2))
        sx = sx + x(i): sy = sy + y(i): sxx = sxx + x(i) * x(i): syy = syy + y(i) * y(i): sxy = sxy + x(i) * y(i)
    Next m
    mx = sx / n: my = sy / n
    sxx = sxx / n - mx * mx: syy = syy / n - my * my: sxy = sxy / n - mx * my
    BridgeAxis = 0.5 * Atan2(2 * sxy, sxx - syy) * 180 / (4 * Atn(1))
    ViewExtents x, y, z, n, BridgeAxis
End Function

' projected width and height of the nodes for the camera of CaptureBody, and the 3D diagonal
' (camera at plan bearing 270 - HORIZONTAL from +X, elevation VERTICAL; measured on the culvert)
Private Sub ViewExtents(x() As Double, y() As Double, z() As Double, ByVal n As Long, ByVal axisDeg As Double)
    Dim i As Long, hor As Double, bt As Double, el As Double, dg As Double
    Dim rx As Double, ry As Double, ux As Double, uy As Double, uz As Double, pr As Double, pu As Double
    Dim r0 As Double, r1 As Double, u0 As Double, u1 As Double
    Dim x0 As Double, x1 As Double, y0 As Double, y1 As Double, z0 As Double, z1 As Double
    dg = 4 * Atn(1) / 180
    hor = CAM_H - axisDeg
    hor = hor - 360 * Int(hor / 360)
    bt = (270 - hor) * dg: el = CAM_V * dg
    rx = -Sin(bt): ry = Cos(bt)
    ux = -Sin(el) * Cos(bt): uy = -Sin(el) * Sin(bt): uz = Cos(el)
    For i = 1 To n
        pr = x(i) * rx + y(i) * ry
        pu = x(i) * ux + y(i) * uy + z(i) * uz
        If i = 1 Then
            r0 = pr: r1 = pr: u0 = pu: u1 = pu
            x0 = x(i): x1 = x(i): y0 = y(i): y1 = y(i): z0 = z(i): z1 = z(i)
        Else
            If pr < r0 Then r0 = pr
            If pr > r1 Then r1 = pr
            If pu < u0 Then u0 = pu
            If pu > u1 Then u1 = pu
            If x(i) < x0 Then x0 = x(i)
            If x(i) > x1 Then x1 = x(i)
            If y(i) < y0 Then y0 = y(i)
            If y(i) > y1 Then y1 = y(i)
            If z(i) < z0 Then z0 = z(i)
            If z(i) > z1 Then z1 = z(i)
        End If
    Next i
    mWm = r1 - r0: mHm = u1 - u0
    mD = Sqr((x1 - x0) ^ 2 + (y1 - y0) ^ 2 + (z1 - z0) ^ 2)
End Sub

' zoom at which the bridge fills FILL of the part of an image (width / height = aspect) that is
' kept when it is cropped to the frame (width / height = frameAsp); 100 when unknown
Private Function ZoomFor(ByVal aspect As Double, ByVal frameAsp As Double) As Long
    Dim fw As Double, fh As Double, s As Double, zz As Double
    ZoomFor = 100
    If mD <= 0 Or mWm <= 0 Or mHm <= 0 Or aspect <= 0 Or frameAsp <= 0 Then Exit Function
    If aspect >= frameAsp Then                   ' image wider than the frame: the sides are cropped
        fw = mWm / (K_FIT * mD * frameAsp)
        fh = mHm / (K_FIT * mD)
    Else                                         ' image taller: top and bottom are cropped
        fw = mWm / (K_FIT * mD * aspect)
        fh = mHm / (K_FIT * mD) * frameAsp / aspect
    End If
    s = FILL / fw
    If FILL / fh < s Then s = FILL / fh
    zz = 100 + Log(s) / ZOOM_RATE
    If zz < 25 Then zz = 25
    If zz > 200 Then zz = 200
    ZoomFor = CLng(zz)
End Function

' width and height of a PNG file from its IHDR header
Private Function PngSize(ByVal f As String, ByRef w As Long, ByRef h As Long) As Boolean
    Dim fn As Integer, b(0 To 23) As Byte
    On Error GoTo Bad
    fn = FreeFile
    Open f For Binary Access Read As #fn
    Get #fn, 1, b
    Close #fn
    If b(1) <> 80 Or b(2) <> 78 Or b(3) <> 71 Then Exit Function      ' "PNG"
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

' ---- Civil NX window back to maximized ----
Private Sub MaximizeCivil()
    Dim h As LongPtr, buf As String, n As Long
    On Error Resume Next
    h = GetWindow(GetDesktopWindow(), 5)          ' GW_CHILD: first top-level window
    Do While h <> 0
        If IsWindowVisible(h) <> 0 Then
            buf = String$(512, vbNullChar)
            n = GetWindowTextW(h, StrPtr(buf), 512)
            If n > 0 Then
                If InStr(1, Left$(buf, n), "MIDAS CIVIL NX", vbTextCompare) = 1 Then ShowWindow h, 3: Exit Sub
            End If
        End If
        h = GetWindow(h, 2)                       ' GW_HWNDNEXT
    Loop
End Sub

' ---- checks ----
Private Function LayoutOk(ByRef msg As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(MOD_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then msg = "STOPPED: no sheet " & MOD_SHEET & " in this workbook.": Exit Function
    If Replace(UCase$(Trim$(CStr(ws.Cells(R_PCT - 2, 3).Value))), " ", "") <> "MODENO" Then
        msg = "STOPPED: sheet MOD has another layout ('Mode No' expected in C" & R_PCT - 2 & "). Nothing was captured."
        Exit Function
    End If
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FIG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then msg = "STOPPED: no sheet " & FIG_SHEET & " in this workbook.": Exit Function
    LayoutOk = True
End Function

Private Function ModelTag(ByVal pj As String) As String
    Dim p As Long, q As Long, t As String, c As String
    t = JStr(pj, "TITLE"): c = JStr(pj, "COMMENT")
    ModelTag = t & ", " & c
    If t = "" And c = "" Then ModelTag = "model without project data"
End Function

' value of a flat string field of a JSON answer (first occurrence; \uXXXX decoded)
Private Function JStr(ByVal s As String, ByVal key As String) As String
    Dim p As Long, o As String, ch As String
    p = InStr(1, s, """" & key & """", vbBinaryCompare)
    If p = 0 Then Exit Function
    p = InStr(p + Len(key) + 2, s, """")
    If p = 0 Then Exit Function
    p = p + 1
    Do While p <= Len(s)
        ch = Mid$(s, p, 1)
        If ch = """" Then Exit Do
        If ch = "\" Then
            p = p + 1: ch = Mid$(s, p, 1)
            If ch = "u" Then ch = ChrW(CLng("&H" & Mid$(s, p + 1, 4))): p = p + 4
        End If
        o = o & ch
        p = p + 1
    Loop
    JStr = o
End Function

Private Function JNum(ByVal v As Double) As String
    Dim s As String
    s = Trim$(Str$(Round(v, 6)))
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    JNum = s
End Function

' ---- API ----
Private Function Api(ByVal method As String, ByVal ep As String, ByVal body As String, ByRef resp As String) As Boolean
    On Error GoTo Bad
    If http Is Nothing Then Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open method, API_BASE & ep, False
    http.SetTimeouts 30000, 60000, 300000, 300000
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
    cellKey = Trim$(CStr(ThisWorkbook.Names(KEY_NAME).RefersToRange.Cells(1, 1).Value))
    If cellKey = "" Then cellKey = Trim$(CStr(ThisWorkbook.Worksheets(KEY_SHEET).Range(KEY_CELL).Value))
    On Error GoTo 0
    If cellKey <> "" Then mKey = cellKey
    If mKey = "" Then mKey = Environ$("MIDAS_MAPI_KEY")
    If mKey = "" Then mKey = UiInputS("MIDAS MAPI key (kept in memory for this Excel session only):", MSG_TITLE)
    GetKey = (mKey <> "")
End Function
