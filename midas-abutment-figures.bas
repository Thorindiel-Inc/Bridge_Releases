Attribute VB_Name = "MIDAS_KA_SEKIL"
Option Explicit
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal h As LongPtr, ByVal c As Long) As LongPtr
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowTextW Lib "user32" (ByVal h As LongPtr, ByVal s As LongPtr, ByVal n As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal h As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal h As LongPtr, ByVal n As Long) As Long
' ============================================================================
'  MIDAS_KA_SEKIL - pictures of the abutment model open in Civil NX into the figure pages of the
'  abutment workbook (8'6 KA1 / KA2): model view, load views and plate moment / shear contours on
'  1_RAPOR, pile moment diagrams on 2A_KAZIK. Owner 2026-11-19 (plan: docs/global-input/abutment-figures-plan.md).
'
'  Macro window (MAKROLAR) > MIDAS sonuclari > Kenar ayak sekilleri, or MIDAS_KA_SEKIL_AL.
'  Window field 'islem': 1 all (default), 2 model + loads, 3 results; field 'res': screen size
'  (default: no WIDTH / HEIGHT sent, the Civil NX window stays as it is) or 4000 / 3000 / 2000 px for
'  a full-width picture (Civil NX resizes its window for each capture; it is maximized again).
'  It only changes the view (POST view/CAPTURE) and reads db/PJCF, STLD, LCOM-GEN, GRUP, ELEM, NODE:
'  nothing is written to the model. The last capture shows the whole model again, loads off.
'
'  Slots: every picture row of a figure page is a cell range with a sheet-level name M_KS_<figure>_<n>
'  (created at the old address when missing; the owner can move or resize it in the Name Manager).
'  A slot row holds 1 or 2 windows side by side, a window 1 or 2 panels (the two wing walls or ear
'  walls are captured one by one). Each picture is cropped (centred) to its cell, sized to it and gets
'  a thin black border; then the pictures that sat in that cell, and the hand-pasted pictures of the
'  slot row, are deleted. A picture that cannot be made leaves its cell as it is and is reported.
'  Figures (rows of BuildList):
'    MODEL  solid view of the whole model
'    EH EQ EV PEQ   pressure loads of one load case each, values on, walls + footing shown
'    UST BR         seat loads (nodal) of one load case each, backwall + top of the stem shown
'    MPY MPD MRY MRD MTE MTB   moments, slot rows Dayanim (ENV_STR) / Deprem (ENV_DEP), windows
'                   IC | DIS (walls) or UST | ALT (footing)
'    MA MK          backwall / ear walls: windows Yatay | Dusey, largest absolute value
'    VPR VT         shear: windows Dayanim | Deprem; stem vertical, wing walls horizontal, footing long / short
'    KZ             pile moment diagrams (2A_KAZIK), Boyuna Mz / Enine My (BeamDiagrams, owner 2026-11-20); the case is the text in the
'                   brackets of the label above the slot ('Kritik Komb.(...)'), which must be the name
'                   of one load case or combination of the model (owner decides the case)
'  Load pictures need a load case of that name: a model built with 'Yuk durumlari + kombinasyonlar'.
'  On a factored-case build they are skipped (a display by load group would overlay the copies of the
'  group in every factored case; measured 2026-11-19: GROUP_SELECTION wins over CASE_SELECTION).
'  Result pictures need the envelope and the analysis results; a Statik model gives the ENV_STR
'  pictures, a Deprem model the ENV_DEP ones, the others stay (as MIDAS_KAZIK).
'  Directions: plate local x of a part is read from the model (node 1 -> node 2); the component for
'  'yatay' / 'dusey' (walls) or 'enine' / 'boyuna' (footing: across = global X, along = global Y) is
'  Mxx or Myy by that, so a wall whose local x is vertical (the right wing wall of today's build) gets
'  the other component. Faces: local z of every wall points to the fill and of the footing up
'  (MIDAS_KENARAYAK); a negative moment is tension on the +z face, so IC / UST = envelope Min and
'  DIS / ALT = envelope Max (FACE_ZP / FACE_ZM; agrees with the reference pictures, not probed live).
'  Camera: plan bearing 270 - HORIZONTAL from +X, elevation VERTICAL (measured). Walls face-on from
'  the side that puts the fill behind the camera (stem, backwall) or from +X (wing and ear walls),
'  footing from above. Zoom so the shown part fills FILL of the kept part of the image: zoom 100
'  fits the projected box of the shown part into the window x FIT_MARGIN, each step scales by
'  exp(ZOOM_RATE) (measured 2026-11-20).
'  Measured 2026-11-20 on an analysed model: PlateForces/Moments with CB + MINMAX, Local, Avg.Nodal, contour,
'  values and legend is accepted; TYPE_OF_DISPLAY.UNDEFORMED as a bare boolean is refused ('second query is
'  wrong', no file), so it is not sent (default: no undeformed shape). OPTIONS LOCAL_UCS Local is needed: with the
'  default (UCS) the vertical walls show nothing (legend 'MAX None'), only the footing draws. Face sign measured on
'  a scratch build (EHA alone, stem Myy -653 kNm/m at the base, fill side in tension): IC = Min, as set. Not probed
'  yet: the pile components (My / Mz for boyuna / enine) and the pile values.
'  Value font: the size set in Civil NX (Display Option) is used; the API has no font setting, so the start
'  question reminds the owner to set it first (as with the Culvert capture macros, owner 2026-11-20).
'  MAPI key: read by MIDAS_CORE (G_MAPI_KEY, INPUT!T3, MIDAS_MAPI_KEY, asked; memory only).
'  Test entries: MIDAS_KA_SEKIL_PLAN [tableDir] [reportFile] (slots, pictures and request bodies;
'  tables from the open model or from a folder of saved db tables / a MIDAS_KENARAYAK_PAYLOAD folder;
'  nothing is captured), MIDAS_KA_SEKIL_PLACETEST tableDir, image [, figure] [, reportFile] (puts an
'  image file into every cell through the same path as a capture), MIDAS_KA_SEKIL_RUN reportFile
'  [, what] [, width] (live captures without questions), MIDAS_KA_SEKIL_RUNFIG reportFile, figure [, width] (one figure).
' ============================================================================

Private Const SCRIPT_ID As String = "midas-abutment-figures"
Private Const SCRIPT_VERSION As String = "2026-11-20a"
Private Const SCRIPT_CHANGELOG As String = "New: model, load, plate moment / shear and pile moment-diagram pictures of the abutment model open in Civil NX into the figure pages of 1_RAPOR and 2A_KAZIK; zoom fits the projected box; ear-wall page hidden through its gate when the model has no ear walls"
Private Const SCRIPT_NEEDS As String = "midas-core"   ' UPDATE adds this module when the workbook lacks it
' macro-window row(s) of this module (PANEL reads them): order|group|title|macro|sheet|text|fields[|modules]
'@TOOL 86|MIDAS sonu\u00E7lar\u0131|Kenar ayak \u015Fekilleri|MIDAS_KA_SEKIL_AL||Civil NX'te a\u00E7\u0131k kenar ayak modelinin model, y\u00FCkleme ve sonu\u00E7 (moment, kesme) g\u00F6r\u00FCnt\u00FClerini al\u0131p 1_RAPOR \u015Fekil sayfalar\u0131na, kaz\u0131k diyagramlar\u0131n\u0131 2A_KAZIK'a yerle\u015Ftirir. Modele bir \u015Fey yazmaz.|islem~\u0130\u015Flem~c~Hepsi/Model + y\u00FCkler/Sonu\u00E7lar;res~\u00C7\u00F6z\u00FCn\u00FCrl\u00FCk (px)~c~Ekran boyutu/4000 px/3000 px/2000 px

Private Const MSG_TITLE As String = "MIDAS_KA_SEKIL"
Private Const SH_REP As String = "1_RAPOR"
Private Const SH_PILE As String = "2A_KAZIK"
Private Const RES_W As String = "0,4000,3000,2000"  ' window field res: width of a full-width picture, 0 = screen size
Private Const NAME_PRE As String = "M_KS_"
Private Const PIC_PRE As String = "KS_"
Private Const ENV_S As String = "ENV_STR"           ' Dayanim
Private Const ENV_D As String = "ENV_DEP"           ' Deprem
Private Const FACE_ZP As String = "Min"             ' tension on the +z face (IC, UST): negative moment
Private Const FACE_ZM As String = "Max"             ' tension on the -z face (DIS, ALT)
Private Const ABS_MAX As String = "All"
Private Const PILE_COMP_B As String = "Mz"          ' boyuna: Mz diagram lies along the bridge (measured 2026-11-20)
Private Const PILE_COMP_E As String = "My"          ' enine: My diagram lies across
Private Const PILE_CAM_B As Double = 285            ' camera from the side (+X), turned 15 deg
Private Const PILE_CAM_E As Double = 15             ' camera along the bridge, turned 15 deg
Private Const PILE_CAM_V As Double = 10
Private Const PILE_SCALE As Double = 2              ' diagram scale
Private Const SEAT_BAND As Double = 1.5             ' m of the stem below the seat shown with the seat loads
Private Const CAM_MODEL_H As Double = 30, CAM_MODEL_V As Double = 25
Private Const CAM_LOAD_H As Double = 20, CAM_LOAD_V As Double = 20
Private Const N_COLOR As Long = 12
Private Const GAP_PT As Double = 3                  ' space between the pictures of a slot row
Private Const WAIT_S As Long = 30
Private Const FILL_RES As Double = 0.92
Private Const FILL_LOAD As Double = 0.8             ' room for the load arrows and values
Private Const FIT_MARGIN As Double = 1.15           ' zoom 100 = projected box x 1.15 (measured 2026-11-20)
Private Const ZOOM_RATE As Double = 0.0203
Private Const LOADS_OFF As String = """LOAD"":{""NODAL_LOAD"":false,""PRESSURE_LOAD"":false,""NODAL_BODY_FORCE"":false,""BEAM_LOAD"":false}"

Private Type TPic
    fig As String          ' figure key
    sheet As String
    slotNo As Long
    slotAddr As String     ' default address of the slot row
    win As Long
    nWin As Long
    pan As Long
    nPan As Long
    what As String         ' text for the log and the report
    skip As String         ' why it is not captured ("" = capture)
    post As Boolean        ' result view
    hidden As Boolean
    hor As Double
    vert As Double
    eList As String        ' ACTIVE elements, "" = all
    nKey As String         ' node set for the zoom
    fillShare As Double
    extra As String        ' DISPLAY / RESULT_GRAPHIC part of the request
    rgSet As String        ' result display set through view/RESULTGRAPHIC before the capture ("" = none)
    done As Boolean
End Type

Private mPic() As TPic, mN As Long
Private mCase As Object, mComb As Object, mGrp As Object, mElN As Object, mXYZ As Object, mSets As Object, mPanels As Object
Private mScripted As Boolean, mLog As String
Private mResW As Long, mAspect As Double, mLastMode As String, mNoResult As String, mOnly As String
Private mWm As Double, mHm As Double, mD As Double
' current slot of BuildList
Private cFig As String, cSheet As String, cSlot As Long, cAddr As String, cWin As Long, cNWin As Long

' ============================================================================
Public Sub MIDAS_KA_SEKIL_AL()
    Dim what As Long, s As String
    what = UiIdx("islem", 1)
    If what < 1 Or what > 3 Then what = 1
    s = RunAll(what, PicWidth(UiIdx("res", 1)), True, "", False)
    UiMsg s, IIf(Left$(s, 4) = "Done", vbInformation, vbExclamation), MSG_TITLE
End Sub

' scripted: live captures of the open model without questions; the summary goes into reportFile
Public Sub MIDAS_KA_SEKIL_RUN(ByVal reportFile As String, Optional ByVal what As Long = 1, Optional ByVal resWidth As Long = 0)
    If what < 1 Or what > 3 Then what = 1
    WriteUtf8 reportFile, RunAll(what, resWidth, False, "", False)
End Sub

' scripted: live captures of one figure only (e.g. "KZ", "MPY"); no ear-wall page step
Public Sub MIDAS_KA_SEKIL_RUNFIG(ByVal reportFile As String, ByVal figKey As String, Optional ByVal resWidth As Long = 0)
    mOnly = UCase$(Trim$(figKey))
    WriteUtf8 reportFile, RunAll(1, resWidth, False, "", False)
    mOnly = ""
End Sub

' the plan: slots, pictures and request bodies; nothing is captured. tableDir "" = the open model
Public Function MIDAS_KA_SEKIL_PLAN(Optional ByVal tableDir As String = "", Optional ByVal reportFile As String = "") As String
    MIDAS_KA_SEKIL_PLAN = RunAll(1, 0, False, tableDir, True)
    If Len(reportFile) > 0 Then WriteUtf8 reportFile, MIDAS_KA_SEKIL_PLAN
End Function

' test: image file img into every cell of figure figKey ("*" = all) through the path of a capture
Public Function MIDAS_KA_SEKIL_PLACETEST(ByVal tableDir As String, ByVal img As String, Optional ByVal figKey As String = "*", Optional ByVal reportFile As String = "") As String
    Dim msg As String, i As Long, res As String, n As Long, s As String
    On Error GoTo Fail
    mScripted = True: mLog = ""
    msg = LoadFacts(tableDir)
    If msg <> "" Then s = msg: GoTo Quit
    BuildList 1
    For i = 1 To mN
        If mPic(i).skip = "" And (figKey = "*" Or UCase$(figKey) = mPic(i).fig) Then
            res = PlacePic(mPic(i), img)
            If res = "" Then n = n + 1 Else s = s & vbCrLf & "  ! " & PicName(mPic(i)) & ": " & res
        End If
    Next i
    s = "placed " & n & s
Quit:
    MIDAS_KA_SEKIL_PLACETEST = s
    If Len(reportFile) > 0 Then WriteUtf8 reportFile, s
    Exit Function
Fail:
    s = "STOPPED: " & Err.Description
    Resume Quit
End Function

' ============================================================================
Private Function RunAll(ByVal what As Long, ByVal pxW As Long, ByVal ask As Boolean, ByVal tableDir As String, ByVal planOnly As Boolean) As String
    Dim pj As String, tag As String, msg As String, i As Long, nCap As Long, nSkip As Long, nOk As Long, s As String
    Dim live As Boolean, folder As String, res As String, rep As String, stopped As Boolean, resp As String
    On Error GoTo Fail
    mScripted = Not ask: mLog = ""
    mResW = pxW: mAspect = 0: mLastMode = "": mNoResult = ""
    live = (Len(tableDir) = 0)
    If SheetOf(SH_REP) Is Nothing Then RunAll = "STOPPED: this workbook has no sheet " & SH_REP & ".": Exit Function
    If live Then
        If Not McGetKey(ThisWorkbook, ask, MSG_TITLE) Then RunAll = "STOPPED: no MAPI key: nothing was captured.": Exit Function
        Logp "Civil NX: model"
        If Not Api("GET", "db/PJCF", "", pj) Then
            If Left$(pj, 10) = "HTTP error" Or pj = "" Then RunAll = "STOPPED: Civil NX does not answer (is it open, is the MAPI connected?). " & Left$(pj, 200): Exit Function
            pj = ""
        End If
        tag = ModelTag(pj)
        If InStr(1, pj, "ABUTMENT", vbTextCompare) = 0 And Not planOnly Then
            If Not ask Then RunAll = "STOPPED: the model open in Civil NX (" & tag & ") is not a MIDAS_KENARAYAK abutment model.": Exit Function
            If UiMsg("The model open in Civil NX (" & tag & ") is not a MIDAS_KENARAYAK abutment model." & vbCrLf & vbCrLf & _
                     "Take its pictures anyway?", vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) <> vbYes Then RunAll = "Stopped: nothing was captured.": Exit Function
        End If
    Else
        tag = "tables from " & tableDir
    End If
    Logp "Civil NX: load cases, combinations, groups, elements, nodes"
    msg = LoadFacts(tableDir)
    If msg <> "" Then RunAll = msg: Exit Function
    BuildList what
    For i = 1 To mN
        If mOnly = "" Or mPic(i).fig = mOnly Then
            If mPic(i).skip = "" Then nCap = nCap + 1 Else nSkip = nSkip + 1
        End If
    Next i
    If planOnly Then RunAll = PlanText(tag): Exit Function
    If Not live Then RunAll = "STOPPED: a table folder only gives the plan (MIDAS_KA_SEKIL_PLAN).": Exit Function
    If nCap = 0 Then RunAll = "STOPPED: no picture can be made from this model (" & tag & ")." & SkipLines(): Exit Function
    If ask Then
        s = "Model: " & tag & vbCrLf & "Civil NX will show and capture " & nCap & " views" & IIf(nSkip > 0, " (" & nSkip & " cannot be made from this model and stay as they are)", "") & "." & vbCrLf & _
            IIf(mResW > 0, "Pictures up to " & mResW & " px wide: the Civil NX window changes size for each capture and is maximized again after it.", "Pictures at the size of the Civil NX view: its window is not changed.") & vbCrLf & _
            "The pictures on the figure pages of " & SH_REP & " / " & SH_PILE & " are replaced." & SkipLines() & vbCrLf & vbCrLf & _
            "Set the font size of the values in Civil NX first (Display Option): the pictures use the size set there." & vbCrLf & vbCrLf & "Continue?"
        If UiMsg(s, vbOKCancel + vbQuestion, MSG_TITLE) <> vbOK Then RunAll = "Stopped: nothing was captured.": Exit Function
    End If
    folder = Environ$("TEMP") & "\Bridge_KA_SEKIL"
    MakeDir folder
    For i = 1 To mN
        If mPic(i).skip = "" And (mOnly = "" Or mPic(i).fig = mOnly) Then
            If UiStopped() Then stopped = True: Exit For
            Prog 100 * i / (mN + 1), mPic(i).what
            res = DoPic(mPic(i), folder)
            If res = "" Then
                mPic(i).done = True: nOk = nOk + 1
            Else
                rep = rep & vbCrLf & "  ! " & PicName(mPic(i)) & " (" & mPic(i).what & "): " & res
            End If
        End If
    Next i
    ' ear-wall page (owner 2026-11-20): hidden when the model has no ear walls, shown again when it has
    If what <> 2 And mOnly = "" Then rep = rep & vbCrLf & "  " & EarPage(mGrp.Exists("KULAK"))
    ' the whole model again, loads off
    Call Api("POST", "view/CAPTURE", "{""Argument"":{""EXPORT_PATH"":""" & PathJ(folder & "\_view.png") & """,""SET_MODE"":""pre"",""SET_HIDDEN"":false," & _
             """ANGLE"":{""HORIZONTAL"":" & JNum(CAM_MODEL_H) & ",""VERTICAL"":" & JNum(CAM_MODEL_V) & "},""ACTIVE"":{""ACTIVE_MODE"":""All""},""DISPLAY"":{" & LOADS_OFF & "}}}", resp)
    If mResW > 0 Then MaximizeCivil
    Prog 100, "Done"
    RunAll = IIf(nOk = nCap And Not stopped, "Done: ", "Finished with problems: ") & nOk & " of " & nCap & " pictures placed. Model: " & tag & vbCrLf & _
             "  Picture " & IIf(mResW > 0, "up to " & mResW & " px wide", "at screen size") & IIf(stopped, vbCrLf & "  Stopped by the user.", "") & rep & SkipLines()
    Exit Function
Fail:
    If mResW > 0 Then MaximizeCivil
    RunAll = "STOPPED: " & Err.Description
End Function

Private Function PicWidth(ByVal choice As Long) As Long
    Dim v As Variant
    v = Split(RES_W, ",")
    If choice >= 1 And choice <= UBound(v) + 1 Then PicWidth = CLng(v(choice - 1))
End Function

Private Function SkipLines() As String
    Dim i As Long, s As String, last As String, t As String
    For i = 1 To mN
        If mPic(i).skip <> "" And Not mPic(i).done Then
            t = mPic(i).fig & ": " & mPic(i).skip
            If t <> last Then s = s & vbCrLf & "  - " & t
            last = t
        End If
    Next i
    If s <> "" Then SkipLines = vbCrLf & "Not made (the cells stay as they are):" & s
End Function

' ============================================================================
'  the pictures of every figure page
' ============================================================================
Private Sub BuildList(ByVal what As Long)
    mN = 0
    ReDim mPic(1 To 120)
    Set mPanels = CreateObject("Scripting.Dictionary")
    If what <> 3 Then
        Slot "MODEL", SH_REP, 1, "D", 203, "W", 226: AddModel
        Slot "EH", SH_REP, 1, "B", 238, "X", 253: AddLoad "EHA", "P"
        Slot "EH", SH_REP, 2, "B", 255, "X", 270: AddLoad "EH" & "0", "P"
        Slot "EQ", SH_REP, 1, "B", 283, "X", 298: AddLoad "EQ-X", "P"
        Slot "EQ", SH_REP, 2, "B", 300, "X", 315: AddLoad "EQ-Y", "P"
        Slot "EV", SH_REP, 1, "B", 330, "X", 345: AddLoad "EV", "P"
        Slot "EV", SH_REP, 2, "B", 347, "X", 362: AddLoad "LS", "P"
        Slot "PEQ", SH_REP, 1, "B", 375, "X", 390: AddLoad "PEQ-X", "P"
        Slot "PEQ", SH_REP, 2, "B", 393, "X", 408: AddLoad "PEQ-Y", "P"
        ' seat loads; WL and CF added (owner 2026-11-19)
        Slot "UST", SH_REP, 1, "B", 422, "X", 427: AddLoad "DC", "N"
        Slot "UST", SH_REP, 2, "B", 429, "X", 434: AddLoad "LL", "N"
        Slot "UST", SH_REP, 3, "B", 436, "X", 441: AddLoad "TU", "N"
        Slot "UST", SH_REP, 4, "B", 443, "X", 448: AddLoad "WS", "N"
        Slot "UST", SH_REP, 5, "B", 450, "X", 455: AddLoad "WL", "N"
        Slot "BR", SH_REP, 1, "B", 467, "X", 471: AddLoad "BR/ACC", "N"
        Slot "BR", SH_REP, 2, "B", 473, "X", 477: AddLoad "PEQ-X", "N"
        Slot "BR", SH_REP, 3, "B", 479, "X", 483: AddLoad "PEQ-Y", "N"
        Slot "BR", SH_REP, 4, "B", 485, "X", 489: AddLoad "CF", "N"
    End If
    If what <> 2 Then
        FaceFig "MPY", 563, 580, "PERDE", "H"
        FaceFig "MPD", 607, 626, "PERDE", "V"
        FaceFig "MRY", 652, 671, "RICAT", "H"
        FaceFig "MRD", 698, 718, "RICAT", "V"
        FaceFig "MTE", 744, 763, "TEMEL", "E"
        FaceFig "MTB", 791, 809, "TEMEL", "B"
        DirFig "MA", 837, 855, "ALIN"
        DirFig "MK", 882, 902, "KULAK"
        ShearRow "VPR", 1, 928, "PERDE", "V"
        ShearRow "VPR", 2, 946, "RICAT", "H"
        ShearRow "VT", 1, 974, "TEMEL", "L"
        ShearRow "VT", 2, 992, "TEMEL", "K"
        If Not SheetOf(SH_PILE) Is Nothing Then
            Slot "KZ", SH_PILE, 1, "B", 285, "X", 300: AddPile "B"
            Slot "KZ", SH_PILE, 2, "B", 304, "X", 319: AddPile "E"
        End If
    End If
End Sub

Private Sub Slot(ByVal fig As String, ByVal sh As String, ByVal n As Long, ByVal c1 As String, ByVal r1 As Long, ByVal c2 As String, ByVal r2 As Long)
    cFig = fig: cSheet = sh: cSlot = n: cAddr = c1 & r1 & ":" & c2 & r2: cWin = 1: cNWin = 1
End Sub

' moments of one direction: slot rows Dayanim / Deprem, windows +z face | -z face
Private Sub FaceFig(ByVal fig As String, ByVal r1 As Long, ByVal r2 As Long, ByVal part As String, ByVal d As String)
    Slot fig, SH_REP, 1, "B", r1, "X", r1 + 13: cNWin = 2
    cWin = 1: AddRes part, "M", d, FACE_ZP, ENV_S
    cWin = 2: AddRes part, "M", d, FACE_ZM, ENV_S
    Slot fig, SH_REP, 2, "B", r2, "X", r2 + 13: cNWin = 2
    cWin = 1: AddRes part, "M", d, FACE_ZP, ENV_D
    cWin = 2: AddRes part, "M", d, FACE_ZM, ENV_D
End Sub

' moments of both directions: slot rows Dayanim / Deprem, windows yatay | dusey, largest absolute value
Private Sub DirFig(ByVal fig As String, ByVal r1 As Long, ByVal r2 As Long, ByVal part As String)
    Slot fig, SH_REP, 1, "B", r1, "X", r1 + 13: cNWin = 2
    cWin = 1: AddRes part, "M", "H", ABS_MAX, ENV_S
    cWin = 2: AddRes part, "M", "V", ABS_MAX, ENV_S
    Slot fig, SH_REP, 2, "B", r2, "X", r2 + 13: cNWin = 2
    cWin = 1: AddRes part, "M", "H", ABS_MAX, ENV_D
    cWin = 2: AddRes part, "M", "V", ABS_MAX, ENV_D
End Sub

' shear of one part and direction: windows Dayanim | Deprem
Private Sub ShearRow(ByVal fig As String, ByVal n As Long, ByVal r1 As Long, ByVal part As String, ByVal d As String)
    Slot fig, SH_REP, n, "B", r1, "X", r1 + 13: cNWin = 2
    cWin = 1: AddRes part, "V", d, ABS_MAX, ENV_S
    cWin = 2: AddRes part, "V", d, ABS_MAX, ENV_D
End Sub

Private Function NewPic(ByVal what As String) As Long
    mN = mN + 1
    If mN > UBound(mPic) Then ReDim Preserve mPic(1 To mN + 60)
    With mPic(mN)
        .fig = cFig: .sheet = cSheet: .slotNo = cSlot: .slotAddr = cAddr
        .win = cWin: .nWin = cNWin: .pan = 1: .nPan = 1
        .what = what: .fillShare = FILL_RES
    End With
    NewPic = mN
End Function

Private Sub AddModel()
    Dim k As Long
    k = NewPic("model, solid view")
    With mPic(k)
        .hidden = True: .hor = CAM_MODEL_H: .vert = CAM_MODEL_V: .nKey = "ALL"
        .extra = """DISPLAY"":{""NODE"":{""NODE"":false}," & LOADS_OFF & "}"
    End With
End Sub

' loads of one load case: kind P = pressure loads on walls + footing, N = nodal (seat) loads on the seat band
Private Sub AddLoad(ByVal lc As String, ByVal kind As String)
    Dim k As Long, setKey As String
    k = NewPic("loads of " & lc & IIf(kind = "P", " (pressure)", " (seat)"))
    setKey = IIf(kind = "P", "WALLS", "SEAT")
    With mPic(k)
        .hor = CAM_LOAD_H: .vert = CAM_LOAD_V: .nKey = setKey: .fillShare = FILL_LOAD
        If Not mCase.Exists(UCase$(lc)) Then
            .skip = "no load case " & lc & " in the model (load pictures need a build with 'Yuk durumlari + kombinasyonlar')"
        ElseIf Not mSets.Exists(setKey) Then
            .skip = "no wall groups (PERDE, ALIN ...) in the model"
        Else
            .eList = CStr(mSets(setKey & "#E"))
            .extra = """DISPLAY"":{""NODE"":{""NODE"":false},""LOAD"":{""CASE_SELECTION"":{""TYPE"":""ST"",""NAME"":" & QS(CStr(mCase(UCase$(lc)))) & "}," & _
                     """LOAD_VALUE"":{""FORMAT"":""Fixed"",""PLACE"":0},""NODAL_LOAD"":" & IIf(kind = "N", "true", "false") & ",""PRESSURE_LOAD"":" & IIf(kind = "P", "true", "false") & _
                     ",""NODAL_BODY_FORCE"":false,""BEAM_LOAD"":false}}"
        End If
    End With
End Sub

' plate result of one part: q M / V, direction d (walls H / V, footing E / B / L / K), envelope side mm, envelope env
Private Sub AddRes(ByVal part As String, ByVal q As String, ByVal d As String, ByVal mm As String, ByVal env As String)
    Dim pans As Collection, i As Long, k As Long, pn As Variant, comp As String, why As String
    Set pans = PanelsOf(part)
    If Not mComb.Exists(UCase$(env)) Then why = "no " & env & " in this model"
    If pans Is Nothing Then why = "no group " & part & " in the model" & IIf(part = "KULAK", " (no ear walls)", "")
    If pans Is Nothing Or why <> "" Then
        k = NewPic(part & " " & q & " " & d & " " & env & " " & mm)
        mPic(k).skip = why
        Exit Sub
    End If
    For i = 1 To pans.Count
        pn = pans(i)          ' Array(eList, nKey, locX, hor, vert, label, longIsX)
        comp = CompOf(q, d, part, CStr(pn(2)), CBool(pn(6)))
        k = NewPic(part & IIf(pans.Count > 1, " " & pn(5), "") & " " & comp & " " & env & " " & mm)
        With mPic(k)
            .pan = i: .nPan = pans.Count
            .post = True: .hor = CDbl(pn(3)): .vert = CDbl(pn(4)): .eList = CStr(pn(0)): .nKey = CStr(pn(1))
            .extra = """DISPLAY"":{""NODE"":{""NODE"":false}," & LOADS_OFF & "},""RESULT_GRAPHIC"":{""CURRENT_MODE"":""PlateForces/Moments""," & _
                     """LOAD_CASE_COMB"":{""TYPE"":""CB"",""MINMAX"":""" & mm & """,""NAME"":" & QS(env) & ",""STEP_INDEX"":1}," & _
                     """OPTIONS"":{""LOCAL_UCS"":{""TYPE"":""Local""},""AVERAGE_NODAL"":{""TYPE"":""Avg.Nodal""}},""COMPONENTS"":{""COMP"":""" & comp & """}," & _
                     """TYPE_OF_DISPLAY"":{" & ContourJ() & ",""VALUES"":{""OPT_CHECK"":true,""VALUE_EXP"":false,""DECIMAL_PT"":0,""SET_ORIENT"":0},""LEGEND"":{""OPT_CHECK"":false}}}"
        End With
    Next i
End Sub

Private Function ContourJ() As String
    ContourJ = """CONTOUR"":{""OPT_CHECK"":true,""NUM_OF_COLOR"":" & N_COLOR & ",""COLOR_TYPE"":""rgb"",""OPTIONS"":{""CONTOUR_FILL"":true,""GRADIENT_FILL"":false}}"
End Function

' the component along direction d: Mxx / Vxx when d is the plate's local x, else Myy / Vyy
Private Function CompOf(ByVal q As String, ByVal d As String, ByVal part As String, ByVal locX As String, ByVal longIsX As Boolean) As String
    Dim wantX As Boolean, isLocX As Boolean
    If part = "TEMEL" Then
        Select Case d
            Case "E": wantX = True               ' across the bridge = global X
            Case "B": wantX = False
            Case "L": wantX = longIsX
            Case Else: wantX = Not longIsX
        End Select
        isLocX = (wantX = (locX = "X"))
    Else
        isLocX = ((d = "H") = (locX = "H"))
    End If
    CompOf = q & IIf(isLocX, "xx", "yy")
End Function

' pile moment diagram (owner 2026-11-20: BeamDiagrams, not the solid BeamForces/Moments contour, whose values only
' print from a setting in the Civil NX window). dirKey B = boyuna: Mz, drawn along the bridge, seen from the side;
' E = enine: My, drawn across, seen along the bridge; the camera is turned 15 deg so the pile rows do not cover each
' other (measured 2026-11-20: the component is taken, values print, a diagram seen in its own plane shows as a line)
Private Sub AddPile(ByVal dirKey As String)
    Dim k As Long, nm As String, why As String, tp As String, isEnv As Boolean, lab As String
    k = NewPic("pile moment diagram, " & IIf(dirKey = "B", "boyuna", "enine"))
    With mPic(k)
        .post = True: .hidden = False: .vert = PILE_CAM_V: .hor = IIf(dirKey = "B", PILE_CAM_B, PILE_CAM_E): .nKey = "KAZIK"
        If Not mSets.Exists("KAZIK") Then .skip = "no group KAZIK in the model (not a piled abutment)": Exit Sub
        .eList = CStr(mSets("KAZIK#E"))
        lab = PileLabel(mPic(k))
        nm = PileCase(lab, tp, isEnv, why)
        If nm = "" Then .skip = why: Exit Sub
        .what = .what & ", " & nm
        .extra = """DISPLAY"":{""NODE"":{""NODE"":false}," & LOADS_OFF & "},""RESULT_GRAPHIC"":{""CURRENT_MODE"":""BeamDiagrams""," & _
                 """LOAD_CASE_COMB"":{""TYPE"":""" & tp & """," & IIf(isEnv, """MINMAX"":""All"",", "") & """NAME"":" & QS(nm) & "}," & _
                 """COMPONENTS"":{""PART"":""Total"",""COMP"":""" & IIf(dirKey = "B", PILE_COMP_B, PILE_COMP_E) & """}," & _
                 """DISPLAY_OPTIONS"":{""FIDELITY"":""5Points"",""FILL"":""Solid"",""SCALE"":" & JNum(PILE_SCALE) & "}," & _
                 """TYPE_OF_DISPLAY"":{" & ContourJ() & ",""VALUES"":{""OPT_CHECK"":true,""VALUE_EXP"":false,""DECIMAL_PT"":0,""SET_ORIENT"":0},""LEGEND"":{""OPT_CHECK"":false}}," & _
                 """OUTPUT_SECT_LOCATION"":{""OPT_BY_MEMBER"":false,""OPT_MAX_MINMAX_ALL"":""AbsMax""}}"
    End With
End Sub

' text of the label above a pile slot: the first cell with a bracket in the 3 rows above the slot,
' slot columns ('Boyuna Yonde - Kritik Komb.(...) :' sits 1 or 2 rows above)
Private Function PileLabel(p As TPic) As String
    Dim r As Range, i As Long, c As Range, t As String
    On Error Resume Next
    Set r = SlotRange(p)
    If r Is Nothing Then Exit Function
    For i = 1 To 3
        If r.Row - i < 1 Then Exit For
        For Each c In r.Rows(1).Offset(-i, 0).Cells
            t = CStr(c.Value)
            If InStr(t, "(") > 0 Then PileLabel = t: Exit Function
        Next c
    Next i
End Function

' the case named in the brackets of the label: one load case or combination of the model
Private Function PileCase(ByVal lab As String, ByRef tp As String, ByRef isEnv As Boolean, ByRef why As String) As String
    Dim a As Long, b As Long, t As String, k As Variant, hits As String, n As Long, one As String
    a = InStr(lab, "("): b = InStrRev(lab, ")")
    If a = 0 Or b <= a + 1 Then why = "no case in brackets in the label above the slot ('" & lab & "')": Exit Function
    t = UCase$(Trim$(Mid$(lab, a + 1, b - a - 1)))
    If mComb.Exists(t) Then
        PileCase = CStr(mComb(t)(0)): tp = "CB": isEnv = CBool(mComb(t)(1)): Exit Function
    End If
    If mCase.Exists(t) Then PileCase = CStr(mCase(t)): tp = "ST": Exit Function
    For Each k In mComb.keys
        If Left$(CStr(k), Len(t)) = t Then n = n + 1: one = CStr(k): hits = hits & IIf(hits = "", "", ", ") & mComb(k)(0)
    Next k
    For Each k In mCase.keys
        If Left$(CStr(k), Len(t)) = t Then n = n + 1: one = CStr(k): hits = hits & IIf(hits = "", "", ", ") & mCase(k)
    Next k
    If n = 1 Then
        If mComb.Exists(one) Then PileCase = CStr(mComb(one)(0)): tp = "CB": isEnv = CBool(mComb(one)(1)) Else PileCase = CStr(mCase(one)): tp = "ST"
        Exit Function
    End If
    If n = 0 Then
        why = "the label names '" & t & "', which is no load case or combination of the model"
    Else
        why = "the label names '" & t & "', which fits " & n & " cases of the model (" & hits & "): the owner picks one"
    End If
End Function

' ============================================================================
'  ear-wall page (owner 2026-11-20: the macro hides it when the model has no ear walls)
'  The page (sheet-level name M_KS_MK_SAYFA, default B876:X920 = one print page) is closed by its
'  gate cell (M_KS_MK_KAPI, default Y917, a free cell outside the merged caption; 1 = open, 0 = closed). The workbook formulas read the gate:
'  figure number AF917 = AF871 + gate, caption code Z917 blank when closed (RAPOR_LISTE skips it),
'  page number W919 = W873 + gate, so the numbers after the page close up. Closing also hides the
'  rows and the drawing objects anchored on them (pictures are 'move but don't size' and would slide
'  onto the next page), takes the page out of the print area (RAPOR_LISTE counts print pages by its
'  pieces) and rebuilds the lists. A workbook whose page number does not read the gate is left as it is.
' ============================================================================
Private Function EarPage(ByVal openIt As Boolean) As String
    Dim ws As Worksheet, pg As Range, gate As Range, c As Range, gAddr As String, okF As Boolean, shp As Shape, r As Long, r1 As Long, r2 As Long
    Dim was As Long, chg As Boolean, s As String, lr As Variant
    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(SH_REP)
    Set pg = McAnchor(ws, NAME_PRE & "MK_SAYFA", "B876:X920", True)
    Set gate = McAnchor(ws, NAME_PRE & "MK_KAPI", "Y917", True)
    r1 = pg.Row: r2 = pg.Row + pg.Rows.Count - 1
    If gate.MergeCells Then EarPage = "Ear-wall page left as it is: the gate cell " & gate.Address(False, False) & " is inside merged cells (a value cannot be written there).": Exit Function
    gAddr = gate.Address(False, False)
    For Each c In ws.Range(ws.Cells(r1, 23), ws.Cells(r2, 23)).Cells
        If c.HasFormula Then
            If InStr(1, Replace(c.Formula, "$", ""), gAddr, vbTextCompare) > 0 Then okF = True
        End If
    Next c
    If Not okF Then
        EarPage = "Ear-wall page (rows " & r1 & "-" & r2 & ") left " & IIf(ws.Rows(r1).Hidden, "hidden", "shown") & ": its page number (W) does not read the gate " & gAddr & " (gate formulas not installed)."
        Exit Function
    End If
    was = CLng(NumOr0(gate.Value))
    If was <> IIf(openIt, 1, 0) Then gate.Value = IIf(openIt, 1, 0): chg = True
    If ws.Rows(r1).Hidden = openIt Then ws.Range(ws.Rows(r1), ws.Rows(r2)).Hidden = Not openIt: chg = True
    For Each shp In ws.Shapes
        r = shp.TopLeftCell.Row
        If r >= r1 And r <= r2 Then
            If (shp.Visible = msoTrue) <> openIt Then shp.Visible = IIf(openIt, msoTrue, msoFalse): chg = True
        End If
    Next shp
    If PrintPiece(ws, pg, openIt) Then chg = True
    If chg Then
        Application.Calculate
        On Error Resume Next
        lr = Application.Run("'" & Replace(ThisWorkbook.Name, "'", "''") & "'!RAPOR_LISTE.LISTELERI_AYARLA_SESSIZ")
        If Err.Number <> 0 Then
            s = " Lists not rebuilt (RAPOR_LISTE.LISTELERI_AYARLA_SESSIZ): " & Err.Description
            Err.Clear
        Else
            s = " Lists rebuilt." & IIf(Len(CStr(lr)) > 0, " " & CStr(lr), "")
        End If
        On Error GoTo Fail
    End If
    EarPage = "Ear-wall page (rows " & r1 & "-" & r2 & ") " & IIf(openIt, "shown", "hidden") & IIf(chg, "", " (no change)") & _
              IIf(openIt, ": the model has ear walls.", ": the model has no ear walls; figure and page numbers after it close up.") & s
    Exit Function
Fail:
    EarPage = "Ear-wall page: FAIL " & Err.Description
End Function

' the page's piece in the print area of ws: out when closed, back in its row order when open; True = changed.
' The pieces are read from PageSetup.PrintArea as text (as RAPOR_LISTE does): RefersToRange of a print area
' of many pieces returns nothing in Turkish Excel (measured 2026-11-20). Written back through PageSetup when
' the address fits in 255 characters, else by giving the sheet's print-area name the union range (a formula
' text over 255 characters is refused with error 1004, a range object is taken; measured 2026-11-20).
Private Function PrintPiece(ws As Worksheet, pg As Range, ByVal openIt As Boolean) As Boolean
    Dim own As Name, nm As Name, parts As Variant, a As String, rg As Range, tops() As Long, adrs() As String, n As Long, i As Long, j As Long, t As Long, ta As String
    Dim inIt As Boolean, addr As String, pa As Range
    parts = Split(Replace(ws.PageSetup.PrintArea, ";", ","), ",")
    If UBound(parts) < 0 Then Exit Function
    ReDim tops(1 To UBound(parts) + 2): ReDim adrs(1 To UBound(parts) + 2)
    For i = LBound(parts) To UBound(parts)
        a = Trim$(parts(i))
        If InStr(a, "!") > 0 Then a = Mid$(a, InStrRev(a, "!") + 1)
        If a <> "" Then
            Set rg = ws.Range(a)
            If rg.Row + rg.Rows.Count - 1 >= pg.Row And rg.Row <= pg.Row + pg.Rows.Count - 1 Then
                inIt = True
            Else
                n = n + 1: tops(n) = rg.Row: adrs(n) = rg.Address
            End If
        End If
    Next i
    If inIt = openIt Then Exit Function              ' already as wanted
    If openIt Then n = n + 1: tops(n) = pg.Row: adrs(n) = pg.Address
    If n = 0 Then Exit Function
    For i = 2 To n                                   ' row order
        t = tops(i): ta = adrs(i): j = i - 1
        Do While j >= 1
            If tops(j) <= t Then Exit Do
            tops(j + 1) = tops(j): adrs(j + 1) = adrs(j): j = j - 1
        Loop
        tops(j + 1) = t: adrs(j + 1) = ta
    Next i
    For i = 1 To n
        addr = addr & IIf(i = 1, "", ",") & adrs(i)
        If pa Is Nothing Then Set pa = ws.Range(adrs(i)) Else Set pa = Union(pa, ws.Range(adrs(i)))
    Next i
    If Len(addr) <= 255 Then
        ws.PageSetup.PrintArea = addr
    Else
        For Each nm In ThisWorkbook.Names
            If Right$(nm.Name, 11) = "!Print_Area" Then
                If SheetNameOf(nm) = ws.Name Then
                    If own Is Nothing Or nm.NameLocal <> nm.Name Then Set own = nm
                End If
            End If
        Next nm
        If own Is Nothing Then Err.Raise vbObjectError + 50, , ws.Name & ": print-area name not found."
        own.RefersTo = pa
    End If
    PrintPiece = True
End Function

Private Function SheetNameOf(nm As Name) As String
    On Error Resume Next
    If TypeOf nm.Parent Is Worksheet Then SheetNameOf = nm.Parent.Name
End Function

' test: the ear-wall page alone (openIt True = show, False = hide); summary into reportFile
Public Function MIDAS_KA_SEKIL_PAGETEST(ByVal openIt As Boolean, Optional ByVal reportFile As String = "") As String
    MIDAS_KA_SEKIL_PAGETEST = EarPage(openIt)
    If Len(reportFile) > 0 Then WriteUtf8 reportFile, MIDAS_KA_SEKIL_PAGETEST
End Function

' ============================================================================
'  model tables
' ============================================================================
' "" = done; else a STOPPED text. dirPath "" = the open model, else a folder of saved tables
Private Function LoadFacts(ByVal dirPath As String) As String
    Dim t As Object, k As Variant, nm As String, el As Variant, e As Variant, rec As Object, xyz As Variant, isEnv As Boolean
    Set mCase = CreateObject("Scripting.Dictionary")
    Set mComb = CreateObject("Scripting.Dictionary")
    Set mGrp = CreateObject("Scripting.Dictionary")
    Set mElN = CreateObject("Scripting.Dictionary")
    Set mXYZ = CreateObject("Scripting.Dictionary")
    Set mSets = CreateObject("Scripting.Dictionary")
    Set t = Tbl("STLD", dirPath)
    If Not t Is Nothing Then
        For Each k In t.keys
            nm = CStr(t(k)("NAME")): mCase(UCase$(nm)) = nm
        Next k
    End If
    Set t = Tbl("LCOM-GEN", dirPath)
    If Not t Is Nothing Then
        For Each k In t.keys
            nm = CStr(t(k)("NAME")): isEnv = False
            If t(k).Exists("iTYPE") Then isEnv = (CLng(NumOr0(t(k)("iTYPE"))) = 1)
            mComb(UCase$(nm)) = Array(nm, isEnv)
        Next k
    End If
    Set t = Tbl("GRUP", dirPath)
    If t Is Nothing Then LoadFacts = "STOPPED: the model has no structure groups (db/GRUP): not a MIDAS_KENARAYAK abutment model.": Exit Function
    For Each k In t.keys
        If t(k).Exists("E_LIST") Then
            el = t(k)("E_LIST")
            If Not IsEmpty(el) Then mGrp(UCase$(CStr(t(k)("NAME")))) = el
        End If
    Next k
    Set t = Tbl("NODE", dirPath)
    If t Is Nothing Then LoadFacts = "STOPPED: the nodes could not be read (db/NODE).": Exit Function
    For Each k In t.keys
        Set rec = t(k)
        mXYZ(CStr(k)) = Array(CDbl(rec("X")), CDbl(rec("Y")), CDbl(rec("Z")))
    Next k
    Set t = Tbl("ELEM", dirPath)
    If t Is Nothing Then LoadFacts = "STOPPED: the elements could not be read (db/ELEM).": Exit Function
    For Each k In t.keys
        mElN(CStr(k)) = t(k)("NODE")
    Next k
    MakeSets
End Function

' one db table as a dictionary; from a folder: every file <TABLE>.json or *_<TABLE>.json ({"TABLE":{..}} or {"Assign":{..}})
Private Function Tbl(ByVal t As String, ByVal dirPath As String) As Object
    Dim resp As String, root As Object, k As Variant, f As String, out As Object, part As Object, kk As Variant, base As String
    If Len(dirPath) = 0 Then
        If Not Api("GET", "db/" & t, "", resp) Then Exit Function
        Set root = ParseJson(resp)
        If root Is Nothing Then Exit Function
        For Each k In root.keys
            If UCase$(CStr(k)) = UCase$(t) Then
                If IsObject(root(k)) Then Set Tbl = root(k)
                Exit Function
            End If
        Next k
        Exit Function
    End If
    f = Dir$(dirPath & "\*.json")
    Do While Len(f) > 0
        base = UCase$(Left$(f, Len(f) - 5))
        If base = UCase$(t) Or Right$(base, Len(t) + 1) = "_" & UCase$(t) Then
            Set root = ParseJson(ReadUtf8(dirPath & "\" & f))
            Set part = Nothing
            If Not root Is Nothing Then
                For Each k In root.keys
                    If UCase$(CStr(k)) = UCase$(t) Or UCase$(CStr(k)) = "ASSIGN" Then
                        If IsObject(root(k)) Then Set part = root(k)
                    End If
                Next k
            End If
            If Not part Is Nothing Then
                If out Is Nothing Then Set out = CreateObject("Scripting.Dictionary")
                For Each kk In part.keys
                    If IsObject(part(kk)) Then Set out(kk) = part(kk) Else out(kk) = part(kk)
                Next kk
            End If
        End If
        f = Dir$()
    Loop
    Set Tbl = out
End Function

' node sets for the zoom and element lists: ALL, WALLS (walls + footing), SEAT, KAZIK
Private Sub MakeSets()
    Dim d As Object, k As Variant, g As Variant, e As Variant, eL As String, zSeat As Double, have As Boolean, z As Double, zMax As Double, nds As Variant, j As Long, zc As Double, n As Long
    Set d = CreateObject("Scripting.Dictionary")
    For Each k In mXYZ.keys: d(k) = 1: Next k
    Set mSets("ALL") = d
    ' walls + footing
    Set d = CreateObject("Scripting.Dictionary"): eL = ""
    For Each g In Array("TEMEL", "PERDE", "RICAT", "ALIN", "KULAK")
        If mGrp.Exists(g) Then
            For Each e In mGrp(g)
                eL = eL & IIf(eL = "", "", ",") & CLng(e)
                AddNodes d, CLng(e)
            Next e
        End If
    Next g
    If d.Count > 0 Then Set mSets("WALLS") = d: mSets("WALLS#E") = eL
    ' seat band: the backwall and the stem elements within SEAT_BAND below it
    Set d = CreateObject("Scripting.Dictionary"): eL = ""
    If mGrp.Exists("ALIN") Then
        For Each e In mGrp("ALIN")
            nds = ElemNodes(CLng(e))
            For j = 1 To 4
                If NodeOk(nds, j) Then
                    z = mXYZ(CStr(CLng(nds(j))))(2)
                    If Not have Or z < zSeat Then zSeat = z: have = True
                End If
            Next j
            eL = eL & IIf(eL = "", "", ",") & CLng(e)
            AddNodes d, CLng(e)
        Next e
    End If
    If mGrp.Exists("PERDE") Then
        If Not have Then
            For Each e In mGrp("PERDE")
                nds = ElemNodes(CLng(e))
                For j = 1 To 4
                    If NodeOk(nds, j) Then
                        z = mXYZ(CStr(CLng(nds(j))))(2)
                        If Not have Or z > zMax Then zMax = z: have = True
                    End If
                Next j
            Next e
            zSeat = zMax
        End If
        For Each e In mGrp("PERDE")
            nds = ElemNodes(CLng(e)): zc = 0: n = 0
            For j = 1 To 4
                If NodeOk(nds, j) Then zc = zc + mXYZ(CStr(CLng(nds(j))))(2): n = n + 1
            Next j
            If n > 0 Then
                If zc / n >= zSeat - SEAT_BAND Then
                    eL = eL & IIf(eL = "", "", ",") & CLng(e)
                    AddNodes d, CLng(e)
                End If
            End If
        Next e
    End If
    If d.Count > 0 Then Set mSets("SEAT") = d: mSets("SEAT#E") = eL
    ' piles
    Set d = CreateObject("Scripting.Dictionary"): eL = ""
    If mGrp.Exists("KAZIK") Then
        For Each e In mGrp("KAZIK")
            eL = eL & IIf(eL = "", "", ",") & CLng(e)
            AddNodes d, CLng(e)
        Next e
    End If
    If d.Count > 0 Then Set mSets("KAZIK") = d: mSets("KAZIK#E") = eL
End Sub

Private Function ElemNodes(ByVal e As Long) As Variant
    If mElN.Exists(CStr(e)) Then ElemNodes = mElN(CStr(e))
End Function

Private Function NodeOk(nds As Variant, ByVal j As Long) As Boolean
    If IsEmpty(nds) Then Exit Function
    If j > UBound(nds) Then Exit Function
    If CLng(NumOr0(nds(j))) <= 0 Then Exit Function
    NodeOk = mXYZ.Exists(CStr(CLng(nds(j))))
End Function

Private Sub AddNodes(d As Object, ByVal e As Long)
    Dim nds As Variant, j As Long
    nds = ElemNodes(e)
    For j = 1 To 4
        If NodeOk(nds, j) Then d(CStr(CLng(nds(j)))) = 1
    Next j
End Sub

' the panels of a part: Collection of Array(eList, nKey, locX, hor, vert, label, longIsX); Nothing = no such group.
' Wing and ear walls: one panel per side (small x = sol, large x = sag); the others one panel.
Private Function PanelsOf(ByVal part As String) As Collection
    Dim out As Collection, e As Variant, nds As Variant, j As Long, n As Long, xc As Double, x0 As Double, x1 As Double, first As Boolean
    Dim side As Long, nSide As Long, eL As String, d As Object, key As String, cnt As Long, xMid As Double
    If mPanels.Exists(part) Then
        If IsObject(mPanels(part)) Then Set PanelsOf = mPanels(part)
        Exit Function
    End If
    mPanels(part) = Empty
    If Not mGrp.Exists(part) Then Exit Function
    first = True
    For Each e In mGrp(part)
        If ElemX(CLng(e), xc) Then
            If first Then x0 = xc: x1 = xc: first = False
            If xc < x0 Then x0 = xc
            If xc > x1 Then x1 = xc
        End If
    Next e
    If first Then Exit Function
    nSide = 1
    If (part = "RICAT" Or part = "KULAK") And x1 - x0 > 0.5 Then nSide = 2
    xMid = (x0 + x1) / 2#
    Set out = New Collection
    For side = 1 To nSide
        Set d = CreateObject("Scripting.Dictionary"): eL = "": cnt = 0
        For Each e In mGrp(part)
            If ElemX(CLng(e), xc) Then
                If nSide = 1 Or (side = 1 And xc <= xMid) Or (side = 2 And xc > xMid) Then
                    eL = eL & IIf(eL = "", "", ",") & CLng(e)
                    AddNodes d, CLng(e): cnt = cnt + 1
                End If
            End If
        Next e
        If cnt > 0 Then
            key = part & "#" & side
            Set mSets(key) = d
            out.Add PanelInfo(part, eL, key, IIf(nSide = 2, IIf(side = 1, "sol", "sag"), ""))
        End If
    Next side
    If out.Count = 0 Then Exit Function
    Set mPanels(part) = out
    Set PanelsOf = out
End Function

' centroid x of a plate element; False for a line element or a missing one
Private Function ElemX(ByVal e As Long, ByRef xc As Double) As Boolean
    Dim nds As Variant, j As Long, n As Long, s As Double
    nds = ElemNodes(e)
    If Not NodeOk(nds, 3) Then Exit Function
    For j = 1 To 4
        If NodeOk(nds, j) Then s = s + mXYZ(CStr(CLng(nds(j))))(0): n = n + 1
    Next j
    If n = 0 Then Exit Function
    xc = s / n: ElemX = True
End Function

' local x class, camera and footing long direction of one panel
Private Function PanelInfo(ByVal part As String, ByVal eL As String, ByVal key As String, ByVal label As String) As Variant
    Dim v As Variant, i As Long, nds As Variant, a As Variant, b As Variant, dx As Double, dy As Double, dz As Double
    Dim nA As Long, nB As Long, locX As String, hor As Double, vert As Double, ang As Double, k As Variant, p As Variant
    Dim sx As Double, sy As Double, sxx As Double, syy As Double, sxy As Double, n As Long, mx As Double, my As Double
    Dim x0 As Double, x1 As Double, y0 As Double, y1 As Double, longIsX As Boolean
    v = Split(eL, ",")
    For i = 0 To UBound(v)
        nds = ElemNodes(CLng(v(i)))
        If NodeOk(nds, 1) And NodeOk(nds, 2) Then
            a = mXYZ(CStr(CLng(nds(1)))): b = mXYZ(CStr(CLng(nds(2))))
            dx = Abs(b(0) - a(0)): dy = Abs(b(1) - a(1)): dz = Abs(b(2) - a(2))
            If part = "TEMEL" Then
                If dx >= dy Then nA = nA + 1 Else nB = nB + 1          ' A: local x along global X
            Else
                If dz > dx And dz > dy Then nB = nB + 1 Else nA = nA + 1   ' A: local x horizontal
            End If
        End If
    Next i
    If part = "TEMEL" Then locX = IIf(nA >= nB, "X", "Y") Else locX = IIf(nA >= nB, "H", "V")
    If nA > 0 And nB > 0 Then mLog = mLog & vbCrLf & "  ! " & part & IIf(label = "", "", " " & label) & ": the plates have mixed local x directions (" & nA & " / " & nB & "): the picture shows one component for all, check it."
    ' plan direction and extents of the panel
    For Each k In mSets(key).keys
        p = mXYZ(CStr(k))
        n = n + 1
        sx = sx + p(0): sy = sy + p(1): sxx = sxx + p(0) * p(0): syy = syy + p(1) * p(1): sxy = sxy + p(0) * p(1)
        If n = 1 Then x0 = p(0): x1 = p(0): y0 = p(1): y1 = p(1)
        If p(0) < x0 Then x0 = p(0)
        If p(0) > x1 Then x1 = p(0)
        If p(1) < y0 Then y0 = p(1)
        If p(1) > y1 Then y1 = p(1)
    Next k
    longIsX = ((x1 - x0) >= (y1 - y0))
    If part = "TEMEL" Then
        hor = 0: vert = 90
    Else
        mx = sx / n: my = sy / n
        sxx = sxx / n - mx * mx: syy = syy / n - my * my: sxy = sxy / n - mx * my
        ang = 0.5 * Atan2(2 * sxy, sxx - syy) * 180 / (4 * Atn(1))      ' wall direction in plan, (-90, 90]
        If ang <= -89.999 Then ang = ang + 180
        hor = 360 - ang
        hor = hor - 360 * Int(hor / 360)
        vert = 0
    End If
    PanelInfo = Array(eL, key, locX, hor, vert, label, longIsX)
End Function

' ============================================================================
'  one picture: capture, wait for the file, place it; "" = done
' ============================================================================
Private Function DoPic(p As TPic, ByVal folder As String) As String
    Dim r As Range, lf As Double, tp As Double, wd As Double, ht As Double, f As String, asp As Double, frameAsp As Double
    Dim w As Long, h As Long, zoom As Long, z2 As Long, pw As Long, ph As Long, pass As Long, body As String, resp As String, ok As Boolean, md As String
    If p.post And mNoResult <> "" Then DoPic = "skipped: " & mNoResult: Exit Function
    Set r = SlotRange(p)
    CellRect p, r, lf, tp, wd, ht
    frameAsp = wd / ht
    f = folder & "\" & PicName(p) & ".png"
    If mResW > 0 Then
        w = CLng(mResW * wd / r.Width)
        If w < 600 Then w = 600
        h = CLng(w / frameAsp)
        If h < 300 Then h = 300
        asp = w / h
    ElseIf mAspect > 0 Then
        asp = mAspect
    Else
        asp = frameAsp
    End If
    Extents p.nKey, p.hor, p.vert
    zoom = ZoomFor(asp, frameAsp, p.fillShare)
    Logp "Civil NX: " & PicName(p) & ", " & p.what & ", zoom " & zoom
    ' Civil NX drops ZOOM_LEVEL on the capture that switches between model and result view
    ' (measured 2026-11-09): that capture is sent once as a warm-up
    md = IIf(p.post, "post", "pre")
    If md <> mLastMode Then
        If Not KillFile(f) Then DoPic = "FAIL: cannot delete the old image " & f: Exit Function
        If Len(p.rgSet) > 0 Then
            If Not Api("POST", "view/RESULTGRAPHIC", "{""Argument"":" & p.rgSet & "}", resp) Then DoPic = Rejected(p, resp): Exit Function
        End If
        body = BodyOf(p, PathJ(f), w, h, zoom)
        If Not Api("POST", "view/CAPTURE", body, resp) Then DoPic = Rejected(p, resp): Exit Function
        ok = WaitForFile(f, WAIT_S)
        If mResW > 0 Then MaximizeCivil
        mLastMode = md
    End If
    For pass = 1 To 2
        If Not KillFile(f) Then DoPic = "FAIL: cannot delete the old image " & f: Exit Function
        If Len(p.rgSet) > 0 Then
            If Not Api("POST", "view/RESULTGRAPHIC", "{""Argument"":" & p.rgSet & "}", resp) Then DoPic = Rejected(p, resp): Exit Function
        End If
        body = BodyOf(p, PathJ(f), w, h, zoom)
        If Not Api("POST", "view/CAPTURE", body, resp) Then DoPic = Rejected(p, resp): Exit Function
        ok = WaitForFile(f, WAIT_S)
        If mResW > 0 Then MaximizeCivil
        If Not ok Then DoPic = "FAIL: Civil NX accepted the request, but no image appeared": Exit Function
        If mResW > 0 Or mAspect > 0 Then Exit For
        If Not PngSize(f, pw, ph) Then Exit For
        mAspect = pw / ph                          ' screen size: the first image tells the picture shape
        z2 = ZoomFor(mAspect, frameAsp, p.fillShare)
        If z2 = zoom Then Exit For
        zoom = z2
    Next pass
    DoPic = PlacePic(p, f)
End Function

' a rejected capture; the first rejected result view ends the result pictures of this run
Private Function Rejected(p As TPic, ByVal resp As String) As String
    Rejected = "FAIL: Civil NX: " & Left$(resp, 200)
    If p.post And mNoResult = "" And InStr(1, resp, "HTTP error", vbTextCompare) = 0 Then
        mNoResult = "Civil NX rejected the first result view (" & Left$(Replace(Replace(resp, vbCr, " "), vbLf, " "), 120) & "): is the model analysed?"
    End If
End Function

Private Function BodyOf(p As TPic, ByVal pathJson As String, ByVal w As Long, ByVal h As Long, ByVal zoom As Long) As String
    Dim b As String
    b = "{""Argument"":{""EXPORT_PATH"":""" & pathJson & ""","
    If w > 0 And h > 0 Then b = b & """WIDTH"":" & w & ",""HEIGHT"":" & h & ","
    b = b & """ZOOM_LEVEL"":" & zoom & ",""ANGLE"":{""HORIZONTAL"":" & JNum(p.hor) & ",""VERTICAL"":" & JNum(p.vert) & "},"
    b = b & """SET_MODE"":""" & IIf(p.post, "post", "pre") & """,""SET_HIDDEN"":" & IIf(p.hidden, "true", "false") & ","
    If Len(p.eList) = 0 Then
        b = b & """ACTIVE"":{""ACTIVE_MODE"":""All""}"
    Else
        b = b & """ACTIVE"":{""ACTIVE_MODE"":""Active"",""E_LIST"":[" & p.eList & "]}"
    End If
    If Len(p.extra) > 0 Then b = b & "," & p.extra
    BodyOf = b & "}}"
End Function

Private Function PathJ(ByVal s As String) As String
    PathJ = Replace(s, "\", "\\")
End Function

Private Function KillFile(ByVal f As String) As Boolean
    If Len(Dir$(f)) > 0 Then
        On Error Resume Next
        Kill f
        On Error GoTo 0
    End If
    KillFile = (Len(Dir$(f)) = 0)
End Function

Private Function PicName(p As TPic) As String
    PicName = PIC_PRE & p.fig & "_" & p.slotNo & "_" & p.win & "_" & p.pan
End Function

' ---- slots ----
Private Function SlotRange(p As TPic) As Range
    Set SlotRange = McAnchor(ThisWorkbook.Worksheets(p.sheet), NAME_PRE & p.fig & "_" & p.slotNo, p.slotAddr, True)
End Function

' the cell of a picture inside its slot row (points)
Private Sub CellRect(p As TPic, r As Range, ByRef lf As Double, ByRef tp As Double, ByRef wd As Double, ByRef ht As Double)
    Dim wWin As Double
    wWin = (r.Width - GAP_PT * (p.nWin - 1)) / p.nWin
    wd = (wWin - GAP_PT * (p.nPan - 1)) / p.nPan
    lf = r.Left + (p.win - 1) * (wWin + GAP_PT) + (p.pan - 1) * (wd + GAP_PT)
    tp = r.Top: ht = r.Height
End Sub

' image file f into the cell of p: cropped (centred) to the cell, sized, thin black border; then the
' pictures that sat in the cell and the hand-pasted pictures of the slot row are deleted
Private Function PlacePic(p As TPic, ByVal f As String) As String
    Dim ws As Worksheet, r As Range, lf As Double, tp As Double, wd As Double, ht As Double, newShp As Shape, shp As Shape
    Dim w0 As Double, h0 As Double, fa As Double, c As Double, cx As Double, cy As Double, old As Collection, o As Variant, inCell As Boolean, inSlot As Boolean
    Set ws = ThisWorkbook.Worksheets(p.sheet)
    Set r = SlotRange(p)
    CellRect p, r, lf, tp, wd, ht
    If wd < 5 Or ht < 5 Then PlacePic = "FAIL: the slot " & NAME_PRE & p.fig & "_" & p.slotNo & " is too small": Exit Function
    ' old pictures: one of this module whose centre is in the cell, or a hand-pasted one that lies
    ' at least 20 % inside the slot row (a pasted picture can straddle two slots of the new layout)
    Set old = New Collection
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            cx = shp.Left + shp.Width / 2#: cy = shp.Top + shp.Height / 2#
            inCell = (cx >= lf And cx <= lf + wd And cy >= tp And cy <= tp + ht)
            inSlot = False
            If Left$(shp.Name, Len(PIC_PRE)) <> PIC_PRE And shp.Width * shp.Height > 0 Then
                inSlot = (Overlap(shp.Left, shp.Width, r.Left, r.Width) * Overlap(shp.Top, shp.Height, r.Top, r.Height) >= 0.2 * shp.Width * shp.Height)
            End If
            If inCell Or inSlot Then old.Add shp
        End If
    Next shp
    On Error Resume Next
    Set newShp = ws.Shapes.AddPicture(f, msoFalse, msoTrue, lf, tp, -1, -1)
    If Err.Number <> 0 Or newShp Is Nothing Then
        PlacePic = "FAIL: the picture could not be inserted (" & Err.Description & "); the cell was left as it is"
        Exit Function
    End If
    On Error GoTo 0
    w0 = newShp.Width: h0 = newShp.Height: fa = wd / ht
    If w0 / h0 > fa Then
        c = (w0 - h0 * fa) / 2#
        newShp.PictureFormat.CropLeft = c
        newShp.PictureFormat.CropRight = c
    Else
        c = (h0 - w0 / fa) / 2#
        newShp.PictureFormat.CropTop = c
        newShp.PictureFormat.CropBottom = c
    End If
    newShp.LockAspectRatio = msoFalse
    newShp.Width = wd
    newShp.Height = ht
    newShp.Left = lf
    newShp.Top = tp
    newShp.LockAspectRatio = msoTrue
    newShp.Placement = xlMove
    With newShp.Line
        .Visible = msoTrue: .ForeColor.RGB = RGB(0, 0, 0): .Weight = 0.75: .DashStyle = msoLineSolid
    End With
    For Each o In old
        o.Delete
    Next o
    newShp.Name = PicName(p)
End Function

' length of the overlap of [a, a + la] and [b, b + lb]
Private Function Overlap(ByVal a As Double, ByVal la As Double, ByVal b As Double, ByVal lb As Double) As Double
    Dim lo As Double, hi As Double
    lo = IIf(a > b, a, b)
    hi = IIf(a + la < b + lb, a + la, b + lb)
    If hi > lo Then Overlap = hi - lo
End Function

' ---- zoom ----
' projected width and height of a node set for the camera, and its 3D diagonal
Private Sub Extents(ByVal key As String, ByVal hor As Double, ByVal vert As Double)
    Dim dg As Double, bt As Double, el As Double, rx As Double, ry As Double, ux As Double, uy As Double, uz As Double
    Dim k As Variant, p As Variant, pr As Double, pu As Double, first As Boolean
    Dim r0 As Double, r1 As Double, u0 As Double, u1 As Double, x0 As Double, x1 As Double, y0 As Double, y1 As Double, z0 As Double, z1 As Double
    mWm = 0: mHm = 0: mD = 0
    If Not mSets.Exists(key) Then Exit Sub
    dg = 4 * Atn(1) / 180
    bt = (270 - hor) * dg: el = vert * dg
    rx = -Sin(bt): ry = Cos(bt)
    ux = -Sin(el) * Cos(bt): uy = -Sin(el) * Sin(bt): uz = Cos(el)
    first = True
    For Each k In mSets(key).keys
        p = mXYZ(CStr(k))
        pr = p(0) * rx + p(1) * ry
        pu = p(0) * ux + p(1) * uy + p(2) * uz
        If first Then
            r0 = pr: r1 = pr: u0 = pu: u1 = pu
            x0 = p(0): x1 = p(0): y0 = p(1): y1 = p(1): z0 = p(2): z1 = p(2)
            first = False
        Else
            If pr < r0 Then r0 = pr
            If pr > r1 Then r1 = pr
            If pu < u0 Then u0 = pu
            If pu > u1 Then u1 = pu
            If p(0) < x0 Then x0 = p(0)
            If p(0) > x1 Then x1 = p(0)
            If p(1) < y0 Then y0 = p(1)
            If p(1) > y1 Then y1 = p(1)
            If p(2) < z0 Then z0 = p(2)
            If p(2) > z1 Then z1 = p(2)
        End If
    Next k
    mWm = r1 - r0: mHm = u1 - u0
    mD = Sqr((x1 - x0) ^ 2 + (y1 - y0) ^ 2 + (z1 - z0) ^ 2)
End Sub

' zoom at which the part fills share of the part of an image (width / height = aspect) that is kept when
' it is cropped to the cell (width / height = frameAsp); 100 when unknown. Civil NX zoom 100 fits the projected
' box of the shown part into the window with a margin: visible height = FIT_MARGIN x max(box height,
' box width / aspect) (measured 2026-11-20 on 64 abutment pictures: margin 1.13-1.16, face-on, from above and iso);
' each zoom step scales by exp(ZOOM_RATE)
Private Function ZoomFor(ByVal aspect As Double, ByVal frameAsp As Double, ByVal share As Double) As Long
    Dim hv As Double, kw As Double, kh As Double, f As Double, zz As Double
    ZoomFor = 100
    If aspect <= 0 Or frameAsp <= 0 Or share <= 0 Then Exit Function
    If mWm <= 0 And mHm <= 0 Then Exit Function
    hv = mHm
    If mWm / aspect > hv Then hv = mWm / aspect
    hv = hv * FIT_MARGIN                         ' visible height of the whole image at zoom 100
    If aspect >= frameAsp Then                   ' image wider than the cell: the sides are cropped
        kh = hv: kw = hv * frameAsp
    Else                                         ' image taller: top and bottom are cropped
        kw = hv * aspect: kh = hv * aspect / frameAsp
    End If
    f = mWm / kw
    If mHm / kh > f Then f = mHm / kh            ' share of the kept part the box fills at zoom 100
    If f <= 0 Then Exit Function
    zz = 100 + Log(share / f) / ZOOM_RATE
    If zz < 25 Then zz = 25
    If zz > 200 Then zz = 200
    ZoomFor = CLng(zz)
End Function

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

' ---- the plan text ----
Private Function PlanText(ByVal tag As String) As String
    Dim i As Long, s As String, r As Range, lf As Double, tp As Double, wd As Double, ht As Double, zoom As Long, nCap As Long
    s = "PLAN. Model: " & tag & vbCrLf & "Load cases: " & mCase.Count & ", combinations: " & mComb.Count & ", groups: " & Join(mGrp.keys, ", ") & vbCrLf
    For i = 1 To mN
        With mPic(i)
            Set r = SlotRange(mPic(i))
            CellRect mPic(i), r, lf, tp, wd, ht
            s = s & PicName(mPic(i)) & " | " & .sheet & "!" & r.Address(False, False) & " window " & .win & "/" & .nWin & " panel " & .pan & "/" & .nPan & _
                " | cell " & Format$(wd, "0") & " x " & Format$(ht, "0") & " pt | " & .what
            If .skip <> "" Then
                s = s & " | NOT MADE: " & .skip & vbCrLf
            Else
                nCap = nCap + 1
                Extents .nKey, .hor, .vert
                zoom = ZoomFor(wd / ht, wd / ht, .fillShare)
                s = s & " | " & Format$(mWm, "0.00") & " x " & Format$(mHm, "0.00") & " m, zoom " & zoom & vbCrLf & "    " & BodyShort(BodyOf(mPic(i), "C:\\T\\" & PicName(mPic(i)) & ".png", 0, 0, zoom)) & vbCrLf
                If Len(.rgSet) > 0 Then s = s & "    view/RESULTGRAPHIC first: {""Argument"":" & .rgSet & "}" & vbCrLf
            End If
        End With
    Next i
    s = s & "Ear-wall page: " & IIf(mGrp.Exists("KULAK"), "shown (the model has ear walls)", "hidden (the model has no ear walls)") & " on a results run" & vbCrLf
    PlanText = s & "Pictures to capture: " & nCap & " of " & mN & mLog
End Function

' the request with the element list shortened
Private Function BodyShort(ByVal b As String) As String
    Dim a As Long, e As Long, n As Long
    a = InStr(b, """E_LIST"":[")
    If a > 0 Then
        e = InStr(a, b, "]")
        n = UBound(Split(Mid$(b, a + 10, e - a - 10), ",")) + 1
        b = Left$(b, a + 9) & "<" & n & " elements>" & Mid$(b, e)
    End If
    BodyShort = b
End Function

' ---- small helpers ----
Private Function SheetOf(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetOf = ThisWorkbook.Worksheets(nm)
End Function

Private Function ModelTag(ByVal pj As String) As String
    Dim root As Object, k As Variant, rec As Object
    On Error GoTo Quit
    Set root = ParseJson(pj)
    If root Is Nothing Then GoTo Quit
    If Not root.Exists("PJCF") Then GoTo Quit
    For Each k In root("PJCF").keys
        Set rec = root("PJCF")(k)
        If rec.Exists("FILE_NAME") Then ModelTag = CStr(rec("FILE_NAME"))
        If rec.Exists("COMMENT") Then ModelTag = ModelTag & IIf(ModelTag = "", "", " (") & CStr(rec("COMMENT")) & IIf(ModelTag = CStr(rec("COMMENT")), "", ")")
        Exit For
    Next k
Quit:
    If ModelTag = "" Then ModelTag = "model without project data"
End Function

Private Function Api(ByVal method As String, ByVal ep As String, ByVal body As String, ByRef resp As String) As Boolean
    Api = McApi(method, ep, body, resp)
    If McStatus() = 404 Then McForgetKey
End Function

Private Sub Logp(ByVal msgLine As String)
    If mScripted Then Exit Sub
    On Error Resume Next
    UiLog msgLine
End Sub

Private Sub Prog(ByVal pct As Double, ByVal txt As String)
    If mScripted Then Exit Sub
    On Error Resume Next
    UiProgress pct, txt
End Sub

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
