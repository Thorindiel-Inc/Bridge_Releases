Attribute VB_Name = "MIDAS_MOD"
Option Explicit
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal h As LongPtr, ByVal c As Long) As LongPtr
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowTextW Lib "user32" (ByVal h As LongPtr, ByVal s As LongPtr, ByVal n As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal h As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal h As LongPtr, ByVal n As Long) As Long
' ============================================================================
'  MIDAS_MOD - modal results of the seismic model into sheet MOD, and its model and
'  mode-shape pictures into sheet Z_YATAY (chapter 8 workbook). MIDAS_SEKIL is merged in
'  (2026-11-10); the retired MIDAS_SEKIL module left in a workbook is empty and can be deleted.
'
'  Macro window (MAKROLAR) > MIDAS sonuclari > Modal analiz sonuclari, or MIDAS_MOD_OKU.
'  Window field 'islem': 1 read + pictures (default), 2 pictures only (no read: the modes come
'  from MOD, whose Q1 stamp must name the open model), 3 read only; field 'res': picture width
'  4000 / 3000 / 2000 px or screen size.
'  Reads the model OPEN in Civil NX through the MAPI; it writes nothing to the model. The read:
'  post/TABLE EIGENVALUEMODE and its sub-tables
'     EIGENVALUE ANALYSIS                      -> MOD C6:J105   (mode, rad/s, Hz, s, tolerance)
'     MODAL PARTICIPATION MASSES PRINTOUT (1)  -> MOD C109:O208 (% mass and sum, TRAN-X..ROTN-Z)
'     MODAL PARTICIPATION MASSES PRINTOUT (2)  -> MOD C211:O310 (mass and sum)
'     MODAL PARTICIPATION FACTOR PRINTOUT      -> MOD C314:N413 (D F H J L N)
'     MODAL DIRECTION FACTOR PRINTOUT          -> MOD C417:N516 (D F H J L N)
'  The addresses are the ones of the old pasted MIDAS text, so 1_RAPOR (Tablo 54, 55,
'  T122 / T123) reads them unchanged. Each block is cleared first (values only, the
'  formats stay); a model with fewer modes leaves the rest of a block empty, a model
'  with more than 100 modes writes the first 100 and says so.
'  The pictures only change the view (view/CAPTURE) and read db/PJCF and db/NODE.
'  Pages of Z_YATAY, found by their caption row (C29, C61 ...):
'     page 1  model, line view with supports (elastic link symbols would cover a long bridge)
'     page 2  model, solid (hidden) view
'     page 3-6  vibration mode shapes, solid view, contour, no legend, no undeformed shape
'  Modes: the 2 largest TRAN-X and the 2 largest TRAN-Y single-mode mass shares of sheet
'  MOD; a mode already taken goes to the next one of that direction; the four are shown in
'  mode-number order. The caption text of a mode page (column AS) becomes "Mod Sekli - n".
'  Each picture replaces the one on its page at the same place and size (the old one is
'  deleted only after the new one is in) and gets a thin black border (0.75 pt, solid; owner
'  2026-11-10, as on the RTI model views). Images go to %TEMP%\Bridge_Z_YATAY.
'  Camera: one 3D view, turned to the bridge direction (principal axis of the node plan). Zoom so
'  the bridge fills FILL of the picture: Civil NX 'zoom fit' (ZOOM_LEVEL 100) shows a height of
'  K_FIT x the 3D diagonal of the model, and each zoom step scales by exp(ZOOM_RATE) (both measured
'  2026-11-09, zoom 60-105); the projected size of the nodes for this camera gives the zoom.
'  Each image is cropped (centred) to the shape of its frame, and the zoom is chosen for the kept
'  part. At screen size the image shape is only known after the first capture: if it needs
'  another zoom, that page is captured once more.
'  Picture size: screen size = no WIDTH / HEIGHT sent, the image is the Civil NX view and its
'  window stays as it is; 4000 / 3000 / 2000 px wide (height from the frame of the page), but
'  Civil NX resizes its window for each capture (measured 2026-11-09), so the macro maximizes it
'  again after each one. On a long deck small pictures get dark: below about 3500 px wide the
'  mesh lines fill the deck.
'  Model check: db/PJCF of a MIDAS_KOPRU seismic model ends with SEISMIC; any other
'  model is read only after a yes in the question box.
'  MAPI key: read by MIDAS_CORE (name G_MAPI_KEY, then INPUT!T3, then MIDAS_MAPI_KEY, then asked; memory only).
'  Test entries: MIDAS_MOD_TEST <saved post/TABLE answer .json> (no Civil NX), MIDAS_MOD_PLAN
'  (modes, target pictures, request bodies; nothing is sent), MIDAS_MOD_PLACETEST <page> <image>,
'  MIDAS_MOD_PICTEST [width] (live captures without dialogs).
' ============================================================================

Private Const SCRIPT_ID As String = "midas-modal-read"
Private Const SCRIPT_VERSION As String = "2026-11-12c"
Private Const SCRIPT_CHANGELOG As String = "Its macro-window row is declared in this module ('@TOOL line), so a changed or new tool needs no PANEL release; Cells read and written found by M_ names (MIDAS_CORE McAnchor), created at the old address when missing; Shared helpers (key, HTTP, JSON, files) are in the module MIDAS_CORE, which UPDATE adds; MAPI key order: G_MAPI_KEY, INPUT!T3, MIDAS_MAPI_KEY, prompt; One module with the pictures: MIDAS_SEKIL merged in (same pages, modes, zoom, border and captions), the retired MIDAS_SEKIL module is emptied by the update; One tool with the mode-shape pictures: window field Islem = read + pictures, pictures only, read only; field res = picture width (pictures by MIDAS_SEKIL); new: periods, frequencies, mass participation and factors of the seismic model into sheet MOD"
Private Const SCRIPT_NEEDS As String = "midas-core"   ' UPDATE adds this module when the workbook lacks it
' macro-window row(s) of this module (PANEL reads them): order|group|title|macro|sheet|text|fields[|modules]
'@TOOL 80|MIDAS sonu\u00E7lar\u0131|Modal analiz sonu\u00E7lar\u0131|MIDAS_MOD_OKU||Sismik modelden periyot, frekans ve k\u00FCtle kat\u0131l\u0131mlar\u0131n\u0131 MOD'a okur; en b\u00FCy\u00FCk X ve Y modlar\u0131n\u0131n \u015Fekillerini Z_YATAY'a resimler. Modele bir \u015Fey yazmaz.|islem~\u0130\u015Flem~c~Oku + resimle/Yaln\u0131z resimle/Yaln\u0131z oku;res~\u00C7\u00F6z\u00FCn\u00FCrl\u00FCk (px)~c~4000 px/3000 px/2000 px/Ekran boyutu

Private Const MSG_TITLE As String = "MIDAS_MOD"
Private Const OUT_SHEET As String = "MOD"
Private Const N_MAX As Long = 100
Private Const R_EIG As Long = 6, R_PCT As Long = 109, R_MAS As Long = 211, R_FAC As Long = 314, R_DIR As Long = 417
Private Const STAMP_CELL As String = "Q1"
Private Const RES_W As String = "4000,3000,2000,0"  ' window field res: picture width, 0 = screen size

Private Const FIG_SHEET As String = "Z_YATAY"
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

Private mResW As Long                           ' 0 = screen size
Private mWm As Double, mHm As Double, mD As Double ' projected width, height, 3D diagonal (m)
Private mAspect As Double                       ' picture width / height at screen size, 0 = not known yet
Private mLastMode As String                     ' SET_MODE of the last capture ("" at the start of a run)

' ============================================================================
Public Sub MIDAS_MOD_OKU()
    Dim resp As String, msg As String, pj As String, mode As Long, pxW As Long, pics As String, ok As Boolean
    On Error GoTo Fail
    ' field islem: 1 read + pictures, 2 pictures only, 3 read only (owner 2026-11-09: one tool, like MIDAS_RTI)
    mode = UiIdx("islem", 1)
    If mode < 1 Or mode > 3 Then mode = 1
    pxW = PicWidth(UiIdx("res", 1))
    If mode = 2 Then
        ' pictures only: the picture part asks its own questions (model, MOD stamp, start)
        pics = Capture(pxW, True, True, False)
        UiMsg pics, IIf(Left$(pics, 4) = "Done", vbInformation, vbExclamation), MSG_TITLE
        Exit Sub
    End If
    If Not LayoutOk(msg) Then UiMsg msg, vbExclamation, MSG_TITLE: Exit Sub
    If Not GetKey() Then UiMsg "No MAPI key: nothing was read.", vbExclamation, MSG_TITLE: Exit Sub
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
                 IIf(mode = 1, "Read its modal results into MOD and take its pictures anyway?", "Read its modal results into MOD anyway?"), vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) <> vbYes Then Exit Sub
    End If
    If mode = 1 Then
        ' one start question for the read and the pictures
        If UiMsg("The modal results go into MOD, then Civil NX shows and captures 6 views into Z_YATAY (2 model views, the modes with the 2 largest X and the 2 largest Y mass shares)." & vbCrLf & _
                 IIf(pxW > 0, "Pictures " & pxW & " px wide: the Civil NX window changes size for each capture and is maximized again after it.", "Pictures at the size of the Civil NX view: its window is not changed.") & vbCrLf & vbCrLf & _
                 "Continue?", vbOKCancel + vbQuestion, MSG_TITLE) <> vbOK Then Exit Sub
    End If
    UiLog "Civil NX: eigenvalue table"
    resp = ReadTable()
    msg = WriteAll(resp, ModelTag(pj))
    ok = (Left$(msg, 4) = "Done")
    If mode = 1 Then
        pics = Capture(pxW, True, False, True)
        ok = ok And (Left$(pics, 4) = "Done")
        msg = msg & vbCrLf & vbCrLf & "Pictures (Z_YATAY): " & pics
    End If
    UiMsg msg, IIf(ok, vbInformation, vbExclamation), MSG_TITLE
    Exit Sub
Fail:
    UiMsg "STOPPED: " & Err.Description, vbExclamation, MSG_TITLE
End Sub

' the pictures of Z_YATAY, for MIDAS_MOD_OKU (read + pictures, pictures only):
'   pxWidth        picture width in px, 0 = screen size
'   ask            True: questions in the window when the model or the MOD stamp do not fit;
'                  False: such a run stops with a STOPPED line
'   startQuestion  True: the 'Civil NX will show and capture 6 views' question (pictures only)
'   modelChecked   True: the caller has already accepted the open model (read + pictures)
' returns the summary text; it never raises
Private Function Capture(ByVal pxWidth As Long, ByVal ask As Boolean, Optional ByVal startQuestion As Boolean = False, Optional ByVal modelChecked As Boolean = False) As String
    Dim pj As String, msg As String, modes As Variant, ok As Boolean
    On Error GoTo Fail
    If Not PicLayoutOk(msg) Then Capture = msg: Exit Function
    modes = PickModes(msg)
    If IsEmpty(modes) Then Capture = msg: Exit Function
    If Not GetKey() Then Capture = "STOPPED: no MAPI key: nothing was captured.": Exit Function
    mResW = pxWidth
    UiLog "Civil NX: model check (pictures)"
    If Not Api("GET", "db/PJCF", "", pj) Then
        If Left$(pj, 10) = "HTTP error" Or pj = "" Then Capture = "STOPPED: Civil NX does not answer (is it open, is the MAPI connected?). " & Left$(pj, 200): Exit Function
        pj = ""
    End If
    If Not modelChecked Then
        If InStr(1, pj, " SEISMIC", vbBinaryCompare) = 0 Or InStr(1, pj, "midas-bridge-build", vbBinaryCompare) = 0 Then
            If Not ask Then Capture = "STOPPED: the model open in Civil NX is not a MIDAS_KOPRU seismic model.": Exit Function
            If UiMsg("The model open in Civil NX is not a MIDAS_KOPRU seismic model." & vbCrLf & vbCrLf & _
                     "Take its pictures anyway?", vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) <> vbYes Then Capture = "Stopped: nothing was captured.": Exit Function
        End If
    End If
    If Not ModMatches(pj, msg) Then
        If Not ask Then Capture = "STOPPED: " & msg: Exit Function
        If UiMsg(msg & vbCrLf & vbCrLf & "Take the pictures with these modes anyway?", vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) <> vbYes Then Capture = "Stopped: nothing was captured.": Exit Function
    End If
    If startQuestion Then
        If UiMsg("Civil NX will show and capture 6 views (modes " & JoinModes(modes) & ")." & vbCrLf & _
                 IIf(mResW > 0, "Pictures " & mResW & " px wide: the Civil NX window changes size for each capture and is maximized again after it.", "Pictures at the size of the Civil NX view: its window is not changed.") & vbCrLf & vbCrLf & _
                 "Continue?", vbOKCancel + vbQuestion, MSG_TITLE) <> vbOK Then Capture = "Stopped: nothing was captured.": Exit Function
    End If
    Capture = CaptureAll(modes, ok)
    Exit Function
Fail:
    MaximizeCivil
    Capture = "STOPPED: " & Err.Description
End Function

' picture width of the window field res: 4000 / 3000 / 2000 px, 0 = screen size
Private Function PicWidth(ByVal choice As Long) As Long
    Dim v As Variant
    v = Split(RES_W, ",")
    If choice >= 1 And choice <= UBound(v) + 1 Then PicWidth = CLng(v(choice - 1))
End Function

' offline test: a saved post/TABLE EIGENVALUEMODE answer; returns the summary
Public Function MIDAS_MOD_TEST(ByVal jsonPath As String) As String
    Dim s As String, msg As String
    On Error GoTo Fail
    If Not LayoutOk(msg) Then MIDAS_MOD_TEST = msg: Exit Function
    s = ReadUtf8(jsonPath)
    MIDAS_MOD_TEST = WriteAll(s, "TEST " & Mid$(jsonPath, InStrRev(jsonPath, "\") + 1))
    Exit Function
Fail:
    MIDAS_MOD_TEST = "STOPPED: " & Err.Description
End Function

' ---- the table: first only node 1 (small answer), the whole table when its sub-tables are missing ----
Private Function ReadTable() As String
    Dim body As String, resp As String, base As String
    base = """TABLE_NAME"":""E"",""TABLE_TYPE"":""EIGENVALUEMODE"",""UNIT"":{""FORCE"":""KN"",""DIST"":""M""},""STYLES"":{""FORMAT"":""Fixed"",""PLACE"":6}"
    body = "{""Argument"":{" & base & ",""NODE_ELEMS"":{""KEYS"":[1]}}}"
    If Api("POST", "post/TABLE", body, resp) Then
        If SubStart(resp, "EIGENVALUE ANALYSIS") > 0 Then ReadTable = resp: Exit Function
    End If
    UiLog "Civil NX: whole eigenvalue table (node 1 gave no sub-tables)"
    body = "{""Argument"":{" & base & "}}"
    If Not Api("POST", "post/TABLE", body, resp) Then Err.Raise vbObjectError + 1, , "No eigenvalue results (has the model been analysed?). " & Left$(resp, 300)
    If SubStart(resp, "EIGENVALUE ANALYSIS") = 0 Then Err.Raise vbObjectError + 1, , "The answer has no EIGENVALUE ANALYSIS table. " & Left$(resp, 300)
    ReadTable = resp
End Function

' ---- write the five blocks ----
Private Function WriteAll(ByVal resp As String, ByVal tag As String) As String
    Dim ws As Worksheet, eig As Variant, pct As Variant, mas As Variant, fac As Variant, dirf As Variant
    Dim nE As Long, notes As String, calcWas As Long
    eig = SubData(resp, "EIGENVALUE ANALYSIS", 5)
    If IsEmpty(eig) Then Err.Raise vbObjectError + 2, , "The answer has no EIGENVALUE ANALYSIS table."
    pct = SubData(resp, "MODAL PARTICIPATION MASSES PRINTOUT (1)", 13)
    mas = SubData(resp, "MODAL PARTICIPATION MASSES PRINTOUT (2)", 13)
    fac = SubData(resp, "MODAL PARTICIPATION FACTOR PRINTOUT", 7)
    dirf = SubData(resp, "MODAL DIRECTION FACTOR PRINTOUT", 7)
    If IsEmpty(pct) Then Err.Raise vbObjectError + 2, , "The answer has no MODAL PARTICIPATION MASSES (%) table: MOD was not changed."
    nE = UBound(eig)
    If UBound(pct) <> nE Then Err.Raise vbObjectError + 2, , "Mode counts differ: eigenvalues " & nE & ", mass participation " & UBound(pct) & ". MOD was not changed."

    Set ws = ThisWorkbook.Worksheets(OUT_SHEET)
    calcWas = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    WriteBlock ws, R_EIG, eig, Array(3, 4, 6, 8, 10)
    WriteBlock ws, R_PCT, pct, Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15)
    WriteBlock ws, R_MAS, mas, Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15)
    WriteBlock ws, R_FAC, fac, Array(3, 4, 6, 8, 10, 12, 14)
    WriteBlock ws, R_DIR, dirf, Array(3, 4, 6, 8, 10, 12, 14)
    McAnchor(ws, "M_MOD_DAMGA", STAMP_CELL).Value = "MIDAS: " & tag & " / " & Format$(Now, "yyyy-mm-dd hh:nn")
    Application.Calculation = calcWas
    Application.ScreenUpdating = True
    Application.Calculate

    If nE > N_MAX Then notes = notes & vbCrLf & "  ! " & nE & " modes in the model, MOD holds the first " & N_MAX & "."
    If IsEmpty(mas) Then notes = notes & vbCrLf & "  ! no mass table (kN s2/m): rows " & R_MAS & "+ left empty."
    If IsEmpty(fac) Then notes = notes & vbCrLf & "  ! no participation factor table: rows " & R_FAC & "+ left empty."
    If IsEmpty(dirf) Then notes = notes & vbCrLf & "  ! no direction factor table: rows " & R_DIR & "+ left empty."
    WriteAll = "Done: " & IIf(nE > N_MAX, N_MAX, nE) & " modes written to MOD." & vbCrLf & _
        "  Mode 1: T = " & Format$(eig(1)(4), "0.0000") & " s" & vbCrLf & _
        "  Sum of mass, last mode: X " & Format$(pct(nE)(3), "0.00") & " %, Y " & Format$(pct(nE)(5), "0.00") & " %, Z " & Format$(pct(nE)(7), "0.00") & " %" & vbCrLf & _
        "  1_RAPOR " & TbTe(1).Address(False, False) & " Tb = " & TbTe(1).Text & ", " & TbTe(2).Address(False, False) & " Te = " & TbTe(2).Text & notes
End Function

Private Sub WriteBlock(ws As Worksheet, ByVal r0 As Long, data As Variant, cols As Variant)
    Dim n As Long, i As Long, j As Long, colv() As Variant
    ws.Range(ws.Cells(r0, cols(0)), ws.Cells(r0 + N_MAX - 1, cols(UBound(cols)))).ClearContents
    If IsEmpty(data) Then Exit Sub
    n = UBound(data)
    If n > N_MAX Then n = N_MAX
    ' column by column, so the empty gap columns (E, G, I ...) stay untouched
    For j = 0 To UBound(cols)
        ReDim colv(1 To n, 1 To 1)
        For i = 1 To n
            If j = 0 Then colv(i, 1) = CLng(data(i)(1)) Else colv(i, 1) = data(i)(j + 1)
        Next i
        ws.Range(ws.Cells(r0, cols(j)), ws.Cells(r0 + n - 1, cols(j))).Value = colv
    Next j
End Sub

' ---- sub-table DATA as rows of Doubles (1-based rows, 1-based columns); Empty when missing ----
Private Function SubData(ByVal resp As String, ByVal prefix As String, ByVal nCols As Long) As Variant
    Dim p As Long, q As Long, root As Object
    Dim head As Variant, data As Variant, rows() As Variant, r As Long, j As Long, row() As Double
    p = SubStart(resp, prefix)
    If p = 0 Then Exit Function
    q = MatchBrace(resp, p)
    If q = 0 Then Exit Function
    Set root = ParseJson(Mid$(resp, p, q - p + 1))
    If root Is Nothing Then Exit Function
    If Not root.Exists("HEAD") Or Not root.Exists("DATA") Then Exit Function
    head = root("HEAD"): data = root("DATA")
    If IsEmpty(data) Then Exit Function
    If UBound(head) <> nCols Then Err.Raise vbObjectError + 3, , prefix & ": " & UBound(head) & " columns, expected " & nCols & " (" & Join(head, ", ") & ")."
    If CStr(head(1)) <> "ModeNo" Then Err.Raise vbObjectError + 3, , prefix & ": first column is " & CStr(head(1)) & ", expected ModeNo."
    ReDim rows(1 To UBound(data))
    For r = 1 To UBound(data)
        ReDim row(1 To nCols)
        For j = 1 To nCols
            row(j) = Val(CStr(data(r)(j)))
        Next j
        rows(r) = row
    Next r
    SubData = rows
End Function

' position of the "{" that opens the sub-table whose key starts with prefix; 0 when missing
Private Function SubStart(ByVal resp As String, ByVal prefix As String) As Long
    Dim p As Long, q As Long
    p = InStr(1, resp, """" & prefix, vbBinaryCompare)
    If p = 0 Then Exit Function
    q = InStr(p + 1, resp, """:", vbBinaryCompare)
    If q = 0 Then Exit Function
    q = q + 2
    Do While Mid$(resp, q, 1) = " ": q = q + 1: Loop
    If Mid$(resp, q, 1) = "{" Then SubStart = q
End Function

' matching "}" of the "{" at p (strings with braces are skipped)
Private Function MatchBrace(ByRef s As String, ByVal p As Long) As Long
    Dim depth As Long, i As Long, ch As String, inStr_ As Boolean
    For i = p To Len(s)
        ch = Mid$(s, i, 1)
        If inStr_ Then
            If ch = "\" Then
                i = i + 1
            ElseIf ch = """" Then
                inStr_ = False
            End If
        Else
            Select Case ch
                Case """": inStr_ = True
                Case "{": depth = depth + 1
                Case "}"
                    depth = depth - 1
                    If depth = 0 Then MatchBrace = i: Exit Function
            End Select
        End If
    Next i
End Function

' ---- checks ----
Private Function LayoutOk(ByRef msg As String) As Boolean
    Dim ws As Worksheet, r As Variant, bad As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(OUT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then msg = "STOPPED: no sheet " & OUT_SHEET & " in this workbook.": Exit Function
    For Each r In Array(R_EIG - 2, R_PCT - 2, R_MAS - 2, R_FAC - 2, R_DIR - 2)
        If Replace(UCase$(Trim$(CStr(ws.Cells(r, 3).Value))), " ", "") <> "MODENO" Then bad = bad & " C" & r
    Next r
    If bad <> "" Then
        msg = "STOPPED: sheet MOD has another layout ('Mode No' expected in" & bad & "). Nothing was written."
        Exit Function
    End If
    LayoutOk = True
End Function

Private Function ModelTag(ByVal pj As String) As String
    Dim root As Object, k As Variant, rec As Object
    On Error GoTo Done
    Set root = ParseJson(pj)
    If root Is Nothing Then GoTo Done
    If Not root.Exists("PJCF") Then GoTo Done
    For Each k In root("PJCF").keys
        Set rec = root("PJCF")(k)
        If rec.Exists("COMMENT") Then ModelTag = CStr(rec("COMMENT"))
        If rec.Exists("TITLE") Then ModelTag = CStr(rec("TITLE")) & ", " & ModelTag
        Exit For
    Next k
Done:
    If ModelTag = "" Then ModelTag = "model without project data"
End Function

' Tb (k = 1) / Te (k = 2) of 1_RAPOR: name M_MOD_TBTE (T122:T123 of the original layout)
Private Function TbTe(ByVal k As Long) As Range
    Set TbTe = McAnchor(ThisWorkbook.Worksheets("1_RAPOR"), "M_MOD_TBTE", "T122:T123").Cells(k, 1)
End Function

Private Function CellText(ByVal sh As String, ByVal addr As String) As String
    On Error Resume Next
    CellText = CStr(ThisWorkbook.Worksheets(sh).Range(addr).Text)
End Function

' ---- API ----
Private Function Api(ByVal method As String, ByVal ep As String, ByVal body As String, ByRef resp As String) As Boolean
    Api = McApi(method, ep, body, resp)
    If McStatus() = 404 Then McForgetKey
End Function

Private Function GetKey() As Boolean
    GetKey = McGetKey(ThisWorkbook, True, MSG_TITLE)
End Function

' ============================================================================
' ---- pictures (Z_YATAY): from MIDAS_SEKIL 2026-11-10a ----

' test entry without dialogs: live captures on the open model, only when it is a
' MIDAS_KOPRU seismic model and MOD was read from it; returns the summary
Public Function MIDAS_MOD_PICTEST(Optional ByVal resWidth As Long = 0) As String
    Dim pj As String, msg As String, modes As Variant, ok As Boolean
    On Error GoTo Fail
    If Not PicLayoutOk(msg) Then MIDAS_MOD_PICTEST = msg: Exit Function
    modes = PickModes(msg)
    If IsEmpty(modes) Then MIDAS_MOD_PICTEST = msg: Exit Function
    McForgetKey
    If Not McGetKey(ThisWorkbook, False, MSG_TITLE) Then MIDAS_MOD_PICTEST = "STOPPED: no MAPI key": Exit Function
    If Not Api("GET", "db/PJCF", "", pj) Then MIDAS_MOD_PICTEST = "STOPPED: Civil NX: " & Left$(pj, 200): Exit Function
    If InStr(1, pj, " SEISMIC", vbBinaryCompare) = 0 Then MIDAS_MOD_PICTEST = "STOPPED: not a seismic model": Exit Function
    If Not ModMatches(pj, msg) Then MIDAS_MOD_PICTEST = "STOPPED: " & msg: Exit Function
    mResW = resWidth
    MIDAS_MOD_PICTEST = CaptureAll(modes, ok)
    Exit Function
Fail:
    MaximizeCivil
    MIDAS_MOD_PICTEST = "STOPPED: " & Err.Description
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

' test: puts an image file as the picture of page n (1-6) through the same path as a capture
' (crop, size, border, name); "placed" when done
Public Function MIDAS_MOD_PLACETEST(ByVal page As Long, ByVal img As String) As String
    Dim shp As Shape, res As String
    Set shp = PagePicture(CLng(Split(PAGE_ROWS, ",")(page - 1)))
    If shp Is Nothing Then MIDAS_MOD_PLACETEST = "no picture on page " & page: Exit Function
    res = ReplacePicture(shp, img)
    MIDAS_MOD_PLACETEST = IIf(res = "", "placed", res)
End Function

' offline test: modes, target pictures and the request bodies, nothing is sent
Public Function MIDAS_MOD_PLAN() As String
    Dim msg As String, modes As Variant, rows As Variant, i As Long, s As String, shp As Shape
    On Error GoTo Fail
    If Not PicLayoutOk(msg) Then MIDAS_MOD_PLAN = msg: Exit Function
    modes = PickModes(msg)
    If IsEmpty(modes) Then MIDAS_MOD_PLAN = msg: Exit Function
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
    MIDAS_MOD_PLAN = s
    Exit Function
Fail:
    MIDAS_MOD_PLAN = "STOPPED: " & Err.Description
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
    Set ws = ThisWorkbook.Worksheets(OUT_SHEET)
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
    Set ws = ThisWorkbook.Worksheets(OUT_SHEET)
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

' ---- MOD stamp (Q1, written by the read) against the open model ----
Private Function ModMatches(ByVal pj As String, ByRef msg As String) As Boolean
    Dim stamp As String, tag As String
    stamp = CStr(McAnchor(ThisWorkbook.Worksheets(OUT_SHEET), "M_MOD_DAMGA", STAMP_CELL).Value)
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
    With newShp.Line
        .Visible = msoTrue: .ForeColor.RGB = RGB(0, 0, 0): .Weight = 0.75: .DashStyle = msoLineSolid
    End With
    oldShp.Delete
    newShp.Name = nm
End Function

Private Sub SetCaption(ByVal capRow As Long, ByVal mode As Long)
    ThisWorkbook.Worksheets(FIG_SHEET).Range(CAP_COL & capRow).Value = "Mod " & ChrW(350) & "ekli - " & mode
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

' sheets for the pictures: MOD with its header, Z_YATAY
Private Function PicLayoutOk(ByRef msg As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(OUT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then msg = "STOPPED: no sheet " & OUT_SHEET & " in this workbook.": Exit Function
    If Replace(UCase$(Trim$(CStr(ws.Cells(R_PCT - 2, 3).Value))), " ", "") <> "MODENO" Then
        msg = "STOPPED: sheet MOD has another layout ('Mode No' expected in C" & R_PCT - 2 & "). Nothing was captured."
        Exit Function
    End If
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FIG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then msg = "STOPPED: no sheet " & FIG_SHEET & " in this workbook.": Exit Function
    PicLayoutOk = True
End Function
