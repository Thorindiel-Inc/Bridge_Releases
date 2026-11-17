Attribute VB_Name = "GSD_OKU"
Option Explicit
' ============================================================================
'  GSD_OKU - results of MIDAS GSD (General Section Designer) into the data sheets of the 8'7 workbook.
'  Owner 2026-11-17. GSD has no API: the owner runs GSD for a pier group and exports its report to Excel
'  ('<pier runs>_Section<n>.xlsx', e.g. 'P2-P3,P16-P17_Section1.xlsx'); this macro copies the report into
'  the group's data sheet (MIDAS_<pier runs>, found by its group number in M_TESIR_NO) and replaces the
'  pictures. It writes nothing to the exports.
'
'  Macro window (MAKROLAR) > MIDAS sonuclari > GSD sonuclari, or GSD_OKU_AC.
'  Folder: asked on every run (folder picker; it starts at the folder of the last run of this session).
'  File of a group: the .xlsx whose name, spaces ignored, is the group's pier runs (OA_GRUP_TABLO, 2nd
'  column) followed by '_'. A group without a file is reported and skipped.
'  Groups: window field 'Gruplar' = Tumu (every group that has a file) or Secili gruplar (a tick list,
'  PANEL.UiTicks; with an older PANEL a typed list such as 1,3-5).
'  What is copied, export sheet -> area of the data sheet (sheet-level names, created at the default):
'     Property       -> M_GSD_OZELLIK  A1:X60     (rows 1-60 only: the rows below are the force rows)
'     P-M Curve      -> M_GSD_PM       Z1:AW33
'     My-Mz Curve    -> M_GSD_MM       AY1:BV29
'     P-My-Mz Curve  -> M_GSD_ORAN     BX1:CU3839 (the ratio table that Tablo 64 reads; it grows and shrinks)
'  Values only. A cell that holds a formula on the data sheet is never written. Old values of an area
'  that the export leaves empty are cleared. A value that would land inside a merged cell (not its first
'  cell) is counted and reported: the layout of the sheet differs from the export.
'  Pictures: the picture of each of the four export sheets replaces the picture(s) of its area on the data
'  sheet: at the export's own place in the area, or at the place of the old picture when that is close to it
'  (within 8 columns and 4 rows), and at the old width. Then the two figures of the group sheet (found by
'  M_AYAK_ZINCIR; areas M_GSD_SEKIL_PM and M_GSD_SEKIL_MM) are made again as pictures of the data sheet
'  cells (heading row + picture) and replace the old figures at their place and size.
'  Check: the design loads of the export (Property, '3. Design Load Combination') are compared with the
'  rows MIDAS_AYAK read (M_TESIR_BLOK) by name; differences are reported, nothing is stopped.
'  Stamp: M_GSD_DAMGA (file, its date, time of the read).
'  Test entry: GSD_OKU_RUN <folder>, <groups: empty = all, or 1,3-5>, <report file> (no questions).
' ============================================================================

Private Const SCRIPT_ID As String = "gsd-report-reader"
Private Const SCRIPT_VERSION As String = "2026-11-17a"
Private Const SCRIPT_CHANGELOG As String = "new: GSD Excel exports (Property, P-M Curve, My-Mz Curve, P-My-Mz Curve; values and pictures) into the MIDAS_<group> data sheets of the 8'7 workbook, every group or ticked groups, design loads checked against the MIDAS_AYAK rows"
Private Const SCRIPT_NEEDS As String = "midas-core"   ' UPDATE adds this module when the workbook lacks it
' macro-window row(s) of this module (PANEL reads them): order|group|title|macro|sheet|text|fields[|modules]
'@TOOL 86|MIDAS sonu\u00E7lar\u0131|GSD sonu\u00E7lar\u0131|GSD_OKU_AC||GSD'nin Excel'e aktard\u0131\u011F\u0131 raporlar\u0131 (<ayaklar>_Section1.xlsx) grubun MIDAS_<grup> sayfas\u0131na okur: kesit \u00F6zellikleri, P-M ve My-Mz e\u011Frileri, oran tablosu ve resimler; grup sayfas\u0131ndaki iki grafi\u011Fi yeniler. Klas\u00F6r her \u00E7al\u0131\u015Ft\u0131rmada sorulur. GSD dosyalar\u0131na bir \u015Fey yazmaz.|grup~Gruplar~c~T\u00FCm\u00FC/Se\u00E7ili gruplar

Private Const MSG_TITLE As String = "GSD_OKU"
Private Const N_GTAB As String = "OA_GRUP_TABLO"
Private Const M_DATA_NO As String = "M_TESIR_NO"
Private Const M_CHAIN As String = "M_AYAK_ZINCIR"
Private Const GROUP_TPL As String = "SABLON_AYAK"
Private Const SH_PROP As String = "Property"
Private Const SH_PM As String = "P-M Curve"
Private Const SH_MM As String = "My-Mz Curve"
Private Const SH_RATIO As String = "P-My-Mz Curve"
Private Const LOAD_TOL As Double = 0.02       ' kN, kN.m: design load of the export against the MIDAS_AYAK row
Private Const PIC_TYPE As Long = 13           ' msoPicture
Private Const PIC_LINKED As Long = 11         ' msoLinkedPicture
Private Const NEAR_COLS As Long = 8           ' an old picture this close to the export's place keeps its place
Private Const NEAR_ROWS As Long = 4
Private Const PIC_COLS As Long = 16           ' a curve picture spans the first 16 columns of its sheet (A:P), the table follows

Private mScripted As Boolean
Private mLog As String
Private mLastFolder As String
Private mLastW As Double                      ' width of the last picture that took an old picture's width

' ============================================================================
Public Sub GSD_OKU_AC()
    Dim s As String
    mScripted = False
    s = RunPull("", "")
    UiMsg s, IIf(Left$(s, 8) = "STOPPED:", vbExclamation, vbInformation), MSG_TITLE
End Sub

Public Sub GSD_OKU_RUN(ByVal folder As String, ByVal groups As String, ByVal reportFile As String)
    mScripted = True
    WriteText reportFile, RunPull(folder, groups)
End Sub

' ============================================================================
Private Function RunPull(ByVal folder As String, ByVal groups As String) As String
    Dim wb As Workbook, gt As Variant, nG As Long, k As Long, files() As String, pick() As Boolean, nFile As Long, nPick As Long
    Dim s As String, sumTxt As String, cm As Long, act As Object, nOk As Long, one As String
    On Error GoTo Fail
    mLog = ""
    Set wb = ThisWorkbook
    gt = NmRange(wb, N_GTAB).Value
    nG = UBound(gt, 1)

    ' ---- folder ----
    If Not mScripted Then
        folder = AskFolder(wb)
        If Len(folder) = 0 Then RunPull = "STOPPED: no folder chosen: nothing was read.": Exit Function
    End If
    If Right$(folder, 1) = "\" Then folder = Left$(folder, Len(folder) - 1)
    If Len(Dir$(folder, vbDirectory)) = 0 Then RunPull = "STOPPED: folder not found: " & folder: Exit Function
    mLastFolder = folder

    ' ---- file of each group ----
    ReDim files(1 To nG): ReDim pick(1 To nG)
    nFile = MatchFiles(folder, gt, nG, files)
    If nFile = 0 Then RunPull = "STOPPED: no GSD export of a pier group in " & folder & " (expected '<pier runs>_Section1.xlsx', e.g. '" & FirstGroupName(gt, nG) & "_Section1.xlsx').": Exit Function

    ' ---- groups to read ----
    If mScripted Then
        If Len(Trim$(groups)) = 0 Then
            For k = 1 To nG: pick(k) = (Len(files(k)) > 0): Next k
        Else
            ParseList groups, nG, pick
        End If
        For k = 1 To nG
            If Len(CellTxt(gt(k, 2))) = 0 Then pick(k) = False
        Next k
    Else
        If UiIdx("grup", 1) = 2 Then
            If Not AskGroups(wb, gt, nG, files, pick) Then RunPull = "STOPPED: nothing was read.": Exit Function
        Else
            For k = 1 To nG: pick(k) = (Len(files(k)) > 0): Next k
        End If
    End If
    For k = 1 To nG
        If pick(k) Then
            If Len(files(k)) = 0 Then
                mLog = mLog & "Group " & k & " (" & CellTxt(gt(k, 2)) & "): no export file in the folder, skipped." & vbCrLf
                pick(k) = False
            Else
                nPick = nPick + 1
            End If
        End If
    Next k
    If nPick = 0 Then RunPull = "STOPPED: no picked group has an export file." & vbCrLf & mLog: Exit Function

    If Not mScripted Then
        s = "Folder: " & folder & vbCrLf & "Groups:"
        For k = 1 To nG
            If pick(k) Then s = s & vbCrLf & "  " & k & "  " & CellTxt(gt(k, 2)) & "  <-  " & files(k)
        Next k
        s = s & vbCrLf & vbCrLf & "The GSD values and pictures of these groups are replaced on their data sheets. Continue?"
        If UiMsg(s, vbOKCancel + vbQuestion, MSG_TITLE) <> vbOK Then RunPull = "STOPPED: nothing was read.": Exit Function
    End If

    ' ---- read ----
    Set act = ActiveSheet
    cm = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    For k = 1 To nG
        If pick(k) Then
            If UiStopped() Then mLog = mLog & "Stopped by the user before group " & k & "." & vbCrLf: Exit For
            Logp "GSD: group " & k & " (" & CellTxt(gt(k, 2)) & ")"
            one = PullGroup(wb, k, CellTxt(gt(k, 2)), folder & "\" & files(k))
            If Left$(one, 5) <> "FAIL:" Then nOk = nOk + 1
            sumTxt = sumTxt & one & vbCrLf
        End If
    Next k
    Application.Calculation = cm
    Application.Calculate
    On Error Resume Next
    act.Parent.Activate
    act.Activate
    On Error GoTo Fail
    Application.ScreenUpdating = True
    RunPull = "Done. Folder: " & folder & vbCrLf & nOk & " of " & nPick & " groups read." & vbCrLf & sumTxt & mLog
    Exit Function
Fail:
    s = Err.Description
    On Error Resume Next
    If cm <> 0 Then Application.Calculation = cm
    Application.ScreenUpdating = True
    RunPull = "STOPPED: " & s & IIf(Len(sumTxt) > 0, vbCrLf & sumTxt, "") & IIf(Len(mLog) > 0, vbCrLf & mLog, "")
End Function

' one group: its export into its data sheet; returns the summary line ("FAIL: ..." when it could not be read)
Private Function PullGroup(wb As Workbook, ByVal k As Long, ByVal gName As String, ByVal path As String) As String
    Dim ds As Worksheet, gs As Worksheet, xb As Workbook, s As String, note As String, ev As Boolean
    Dim pPM As Object, pMM As Object, st As Range, w As Object
    On Error GoTo Fail
    mLastW = 0
    Set ds = DataSheetOf(wb, k)
    If ds Is Nothing Then Err.Raise vbObjectError + 21, , "no data sheet: run 'Ayak grubu sayfalari' (ORTA_AYAK) first"
    ev = Application.EnableEvents
    Application.EnableEvents = False
    Set xb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Application.EnableEvents = ev
    For Each w In xb.Windows: w.Visible = False: Next w
    If Not HasSheet(xb, SH_PROP) Or Not HasSheet(xb, SH_PM) Or Not HasSheet(xb, SH_MM) Or Not HasSheet(xb, SH_RATIO) Then
        Err.Raise vbObjectError + 22, , "not a GSD report export (sheets " & SH_PROP & ", " & SH_PM & ", " & SH_MM & ", " & SH_RATIO & " expected)"
    End If

    ' ---- values ----
    s = "prop " & CopyBlock(xb.Worksheets(SH_PROP), McAnchor(ds, "M_GSD_OZELLIK", "A1:X60", True), False, note)
    s = s & ", P-M " & CopyBlock(xb.Worksheets(SH_PM), McAnchor(ds, "M_GSD_PM", "Z1:AW33", True), True, note)
    s = s & ", My-Mz " & CopyBlock(xb.Worksheets(SH_MM), McAnchor(ds, "M_GSD_MM", "AY1:BV29", True), True, note)
    s = s & ", ratios " & CopyBlock(xb.Worksheets(SH_RATIO), McAnchor(ds, "M_GSD_ORAN", "BX1:CU3839", True), True, note)

    ' ---- pictures of the data sheet ----
    SwapPicture xb.Worksheets(SH_PROP), ds, McAnchor(ds, "M_GSD_OZELLIK", "A1:X60", True), "GSD_KESIT", note
    Set pPM = SwapPicture(xb.Worksheets(SH_PM), ds, McAnchor(ds, "M_GSD_PM", "Z1:AW33", True), "GSD_PM", note, PIC_COLS)
    Set pMM = SwapPicture(xb.Worksheets(SH_MM), ds, McAnchor(ds, "M_GSD_MM", "AY1:BV29", True), "GSD_MM", note, PIC_COLS)
    SwapPicture xb.Worksheets(SH_RATIO), ds, McAnchor(ds, "M_GSD_ORAN", "BX1:CU3839", True), "GSD_PMM", note

    ' ---- design loads against the MIDAS_AYAK rows ----
    note = note & CheckLoads(xb.Worksheets(SH_PROP), ds)

    ' ---- stamp ----
    Set st = McAnchor(ds, "M_GSD_DAMGA", "DG5", True)
    If Len(CellTxt(st.Offset(0, -1).Value)) = 0 Then st.Offset(0, -1).Value = "GSD"
    st.Value = Mid$(path, InStrRev(path, "\") + 1) & " | " & Format$(FileDateTime(path), "yyyy-mm-dd hh:nn") & " | " & Format$(Now, "yyyy-mm-dd hh:nn")
    xb.Close SaveChanges:=False
    Set xb = Nothing

    ' ---- the two figures of the group sheet ----
    Set gs = GroupSheetOf(wb, k)
    If gs Is Nothing Then
        note = note & "  no group sheet: figures not renewed." & vbCrLf
    Else
        ds.Calculate
        GroupFigure ds, pPM, McAnchor(ds, "M_GSD_PM", "Z1:AW33", True), gs, "M_GSD_SEKIL_PM", "B216:L230", "GSD_PM", note
        GroupFigure ds, pMM, McAnchor(ds, "M_GSD_MM", "AY1:BV29", True), gs, "M_GSD_SEKIL_MM", "M216:X230", "GSD_MM", note
    End If
    PullGroup = ds.Name & ": " & s & IIf(Len(note) > 0, vbCrLf & note, "")
    If Right$(PullGroup, 2) = vbCrLf Then PullGroup = Left$(PullGroup, Len(PullGroup) - 2)
    Exit Function
Fail:
    s = Err.Description
    On Error Resume Next
    Application.EnableEvents = True
    If Not xb Is Nothing Then xb.Close SaveChanges:=False
    PullGroup = "FAIL: group " & k & " (" & gName & "): " & s
End Function

' values of the export sheet into the area (same cell positions). fit True: the export must fit the area;
' False: rows of the export below the area are not read (Property: the design loads). Returns 'n cells'.
Private Function CopyBlock(src As Worksheet, area As Range, ByVal fit As Boolean, ByRef note As String) As String
    Dim nR As Long, nC As Long, nAll As Long, lastT As Long, sv As Variant, tf As Variant, f As Range
    Dim r As Long, c As Long, v As Variant, cur As String, tgt As Range, nW As Long, nClr As Long, nKeep As Long, nLay As Long, isF As Boolean
    nC = area.Columns.Count
    With src.UsedRange
        nR = .Row + .Rows.Count - 1
        If .Column + .Columns.Count - 1 > nC Then note = note & "  " & src.Name & ": columns right of " & nC & " not read." & vbCrLf
    End With
    If nR > area.Rows.Count Then
        If fit Then Err.Raise vbObjectError + 23, , src.Name & ": " & nR & " rows, the area " & area.Address(False, False) & " holds " & area.Rows.Count
        nR = area.Rows.Count
    End If
    sv = src.Range(src.Cells(1, 1), src.Cells(nR, nC)).Value
    Set f = area.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not f Is Nothing Then lastT = f.Row - area.Row + 1
    nAll = IIf(lastT > nR, lastT, nR)
    tf = area.Resize(nAll).Formula
    For r = 1 To nAll
        For c = 1 To nC
            If r <= nR Then v = sv(r, c) Else v = Empty
            If IsError(v) Then v = Empty
            If IsError(tf(r, c)) Then cur = "" Else cur = CStr(tf(r, c))
            If Not IsEmpty(v) Then
                If VarType(v) = vbString Then
                    If Len(v) = 0 Then v = Empty
                End If
            End If
            If IsEmpty(v) And Len(cur) = 0 Then
                ' nothing on either side
            Else
                Set tgt = area.Cells(r, c)
                isF = False
                If Left$(cur, 1) = "=" And Len(cur) > 1 Then isF = tgt.HasFormula
                If isF Then
                    If Not IsEmpty(v) Then nKeep = nKeep + 1
                ElseIf tgt.MergeCells And tgt.Address <> tgt.MergeArea.Cells(1, 1).Address Then
                    If Not IsEmpty(v) Then nLay = nLay + 1
                ElseIf IsEmpty(v) Then
                    tgt.MergeArea.ClearContents
                    nClr = nClr + 1
                Else
                    If VarType(v) = vbString Then
                        If Left$(v, 1) = "=" Or Left$(v, 1) = "'" Then v = "'" & v
                    End If
                    tgt.Value = v
                    nW = nW + 1
                End If
            End If
        Next c
    Next r
    If nKeep > 0 Then note = note & "  " & src.Name & ": " & nKeep & " cells hold a formula on the data sheet and were kept." & vbCrLf
    If nLay > 0 Then note = note & "  " & src.Name & ": " & nLay & " values fall inside merged cells of the data sheet and were not written (its layout differs from the export)." & vbCrLf
    CopyBlock = nW & IIf(nClr > 0, " (-" & nClr & ")", "")
End Function

' the picture of the export sheet replaces the picture(s) whose top left cell is in the area; returns it
Private Function SwapPicture(src As Worksheet, ds As Worksheet, area As Range, ByVal nm As String, ByRef note As String, Optional ByVal capCols As Long = 0) As Object
    Dim sh As Object, pic As Object, i As Long, o As Object, had As Boolean, l As Double, t As Double, w As Double, plc As Long, n As Object, dst As Range
    Dim dist As Long, best As Long, near As Boolean, dc As Long, dr As Long, maxW As Double
    For Each sh In src.Shapes
        If sh.Type = PIC_TYPE Or sh.Type = PIC_LINKED Then Set pic = sh: Exit For
    Next sh
    If pic Is Nothing Then note = note & "  " & src.Name & ": no picture in the export, the old one stays." & vbCrLf: Exit Function
    ' the place of the export's picture inside the area; of the old pictures of the area the one nearest to it
    ' gives the width, and the place too when it is close (the owner moved it a little); the others are leftovers
    ' of an older layout
    Set dst = area.Cells(pic.TopLeftCell.Row, pic.TopLeftCell.Column)
    best = -1
    For i = ds.Shapes.Count To 1 Step -1
        Set o = ds.Shapes(i)
        If o.Type = PIC_TYPE Or o.Type = PIC_LINKED Then
            If Not Intersect(o.TopLeftCell, area) Is Nothing Then
                dc = Abs(o.TopLeftCell.Column - dst.Column): dr = Abs(o.TopLeftCell.Row - dst.Row)
                dist = dc + dr
                If best < 0 Or dist < best Then
                    best = dist
                    l = o.Left: t = o.Top: w = o.Width: plc = o.Placement: had = True
                    near = (dc <= NEAR_COLS And dr <= NEAR_ROWS)
                End If
                o.Delete
            End If
        End If
    Next i
    Set n = PasteShape(pic, Nothing, ds, dst)
    If n Is Nothing Then note = note & "  " & src.Name & ": the picture could not be pasted." & vbCrLf: Exit Function
    n.LockAspectRatio = True
    If had Then
        If near Then n.Left = l: n.Top = t
        n.Width = w
        n.Placement = plc
        mLastW = w
    ElseIf mLastW > 0 Then
        n.Width = mLastW
    End If
    ' capCols: the picture may not run over the table right of it (its width is cut to the first capCols columns)
    If capCols > 0 And capCols < area.Columns.Count Then
        maxW = area.Cells(1, capCols + 1).Left - n.Left - 1
        If maxW > 20 And n.Width > maxW Then n.Width = maxW
    End If
    On Error Resume Next
    n.Name = nm
    On Error GoTo 0
    Set SwapPicture = n
End Function

' figure of the group sheet: a picture of the data sheet cells from the heading row (3rd row of the area)
' down to the GSD picture, at the place and size of the old figure (or fitted into the named area)
Private Sub GroupFigure(ds As Worksheet, pic As Object, area As Range, gs As Worksheet, ByVal nmArea As String, ByVal defAddr As String, ByVal nm As String, ByRef note As String)
    Dim fa As Range, rg As Range, c1 As Long, c2 As Long, r2 As Long, i As Long, o As Object, had As Boolean
    Dim l As Double, t As Double, w As Double, h As Double, plc As Long, n As Object
    If pic Is Nothing Then Exit Sub
    On Error GoTo Fail
    Set fa = McAnchor(gs, nmArea, defAddr, True)
    If fa.Rows(1).EntireRow.Hidden Then note = note & "  " & gs.Name & ": figure rows are hidden, " & nm & " not renewed." & vbCrLf: Exit Sub
    c1 = area.Column + 1
    If pic.TopLeftCell.Column < c1 Then c1 = pic.TopLeftCell.Column
    c2 = pic.BottomRightCell.Column
    r2 = pic.BottomRightCell.Row
    If c2 > c1 And ds.Cells(1, c2).Left >= pic.Left + pic.Width - 2 Then c2 = c2 - 1
    If r2 > area.Row + 2 And ds.Cells(r2, 1).Top >= pic.Top + pic.Height - 2 Then r2 = r2 - 1
    If c2 > area.Column + PIC_COLS - 1 Then c2 = area.Column + PIC_COLS - 1
    Set rg = ds.Range(ds.Cells(area.Row + 2, c1), ds.Cells(r2, c2))
    For i = gs.Shapes.Count To 1 Step -1
        Set o = gs.Shapes(i)
        If o.Type = PIC_TYPE Or o.Type = PIC_LINKED Then
            If Not Intersect(o.TopLeftCell, fa) Is Nothing Then
                l = o.Left: t = o.Top: w = o.Width: h = o.Height: plc = o.Placement: had = True
                o.Delete
            End If
        End If
    Next i
    Set n = PasteShape(Nothing, rg, gs, fa.Cells(1, 1))
    If n Is Nothing Then note = note & "  " & gs.Name & ": figure " & nm & " could not be pasted." & vbCrLf: Exit Sub
    n.LockAspectRatio = True
    If had Then
        n.Left = l: n.Top = t: n.Width = w
        If n.Height > h * 1.02 Then n.Height = h
        n.Placement = plc
    Else
        n.Left = fa.Left: n.Top = fa.Top: n.Width = fa.Width
        If n.Height > fa.Height Then n.Height = fa.Height
    End If
    On Error Resume Next
    n.Name = nm
    Exit Sub
Fail:
    note = note & "  " & gs.Name & ": figure " & nm & " not renewed (" & Err.Description & ")." & vbCrLf
End Sub

' copies a shape (or a range as a picture) and pastes it on ws at dst; returns the new picture. The clipboard
' can be busy for a moment, so it tries a few times; the second try activates the sheet first.
Private Function PasteShape(pic As Object, rg As Range, ws As Worksheet, dst As Range) As Object
    Dim ids As Object, o As Object, tries As Long
    Set ids = CreateObject("Scripting.Dictionary")
    For Each o In ws.Shapes: ids(CStr(o.ID)) = True: Next o
    For tries = 1 To 4
        On Error Resume Next
        Err.Clear
        If tries >= 2 Then
            ws.Parent.Activate
            ws.Activate
            DoEvents
        End If
        If pic Is Nothing Then
            rg.CopyPicture Appearance:=xlPrinter, Format:=xlPicture
        Else
            pic.Copy
        End If
        If Err.Number = 0 Then ws.Paste dst
        On Error GoTo 0
        ' the pasted shape = the one whose ID is new (cell comments are shapes too and come last in Shapes)
        For Each o In ws.Shapes
            If Not ids.Exists(CStr(o.ID)) Then
                If o.Type = PIC_TYPE Or o.Type = PIC_LINKED Then
                    Set PasteShape = o
                    Application.CutCopyMode = False
                    Exit Function
                End If
            End If
        Next o
        Pause 0.3
    Next tries
    Application.CutCopyMode = False
End Function

' design loads of the export (name; Pu, My, Mz / Vy, Vz, T on two rows) against the rows of M_TESIR_BLOK
Private Function CheckLoads(src As Worksheet, ds As Worksheet) As String
    Dim blk As Range, bv As Variant, d As Object, i As Long, j As Long, r As Long, r0 As Long, nR As Long, sv As Variant
    Dim nm As String, nGsd As Long, nMiss As Long, nDif As Long, firstMiss As String, firstDif As String, g(1 To 6) As Double, rowB As Long, bad As Boolean, used As Long, nVz As Long
    On Error GoTo Quit
    Set blk = McAnchor(ds, "M_TESIR_BLOK", "CW26:DE3839", True)
    bv = blk.Value
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(bv, 1)
        nm = CellTxt(bv(i, 3))
        If Len(nm) > 0 Then d(nm) = i: used = i
    Next i
    If used = 0 Then CheckLoads = "  loads: the data sheet has no MIDAS_AYAK rows to compare with." & vbCrLf: Exit Function
    With src.UsedRange
        nR = .Row + .Rows.Count - 1
    End With
    sv = src.Range(src.Cells(1, 1), src.Cells(nR, 24)).Value
    For r = 1 To nR - 2
        If CellTxt(sv(r, 3)) = "No" And CellTxt(sv(r, 5)) = "Name" Then r0 = r + 2: Exit For
    Next r
    If r0 = 0 Then CheckLoads = "  loads: no 'Design Load Combination' table in the export." & vbCrLf: Exit Function
    For r = r0 To nR - 1 Step 2
        nm = CellTxt(sv(r, 5))
        If Len(nm) = 0 Then Exit For
        nGsd = nGsd + 1
        If Not d.Exists(nm) Then
            nMiss = nMiss + 1
            If nMiss = 1 Then firstMiss = nm
        Else
            rowB = d(nm)
            g(1) = NumOr0(sv(r, 10)): g(2) = NumOr0(sv(r, 15)): g(3) = NumOr0(sv(r, 20))
            g(4) = NumOr0(sv(r + 1, 10)): g(5) = NumOr0(sv(r + 1, 15)): g(6) = NumOr0(sv(r + 1, 20))
            bad = False
            ' Vz apart: the GSD export of 2026-11 prints Vy in its Vz column too (measured on two exports)
            For j = 1 To 6
                If Abs(g(j) - NumOr0(bv(rowB, 3 + j))) > LOAD_TOL Then
                    If j = 5 Then nVz = nVz + 1 Else bad = True
                End If
            Next j
            If bad Then
                nDif = nDif + 1
                If nDif = 1 Then firstDif = nm
            End If
            d.Remove nm
        End If
    Next r
    CheckLoads = "  loads: " & nGsd & " in the export, " & (nGsd - nMiss - nDif) & " equal to the MIDAS_AYAK rows"
    If nDif > 0 Then CheckLoads = CheckLoads & "; " & nDif & " DIFFER (first: " & firstDif & ")"
    If nMiss > 0 Then CheckLoads = CheckLoads & "; " & nMiss & " NOT on the data sheet (first: " & firstMiss & ")"
    If nVz > 0 Then CheckLoads = CheckLoads & "; Vz differs in " & nVz & " (Pu, My, Mz, Vy, T compared; the export repeats Vy as Vz)"
    If d.Count > 0 Then CheckLoads = CheckLoads & "; " & d.Count & " rows of the data sheet are not in the export"
    CheckLoads = CheckLoads & "." & vbCrLf
    Exit Function
Quit:
    CheckLoads = "  loads: not compared (" & Err.Description & ")." & vbCrLf
End Function

' ---- folder, files, groups ----
Private Function AskFolder(wb As Workbook) As String
    Dim fd As Object, start As String
    start = mLastFolder
    If Len(start) = 0 Then start = wb.path
    Set fd = Application.FileDialog(4)          ' msoFileDialogFolderPicker
    With fd
        .title = U("GSD Excel raporlar\u0131n\u0131n klas\u00F6r\u00FC (<ayaklar>_Section1.xlsx)")
        .AllowMultiSelect = False
        If Len(start) > 0 Then .InitialFileName = start & "\"
        If .Show <> -1 Then Exit Function
        AskFolder = CStr(.SelectedItems(1))
    End With
End Function

' files(k) = export of group k (its name, spaces ignored, starts with the group's pier runs and '_');
' returns how many groups have one
Private Function MatchFiles(ByVal folder As String, gt As Variant, ByVal nG As Long, files() As String) As Long
    Dim f As String, base As String, k As Long, key As String
    f = Dir$(folder & "\*.xlsx")
    Do While Len(f) > 0
        If Left$(f, 2) <> "~$" Then
            base = Norm(Left$(f, InStrRev(f, ".") - 1))
            For k = 1 To nG
                key = Norm(CellTxt(gt(k, 2)))
                If Len(key) > 0 Then
                    If base = key Or Left$(base, Len(key) + 1) = key & "_" Then
                        If Len(files(k)) = 0 Then
                            files(k) = f
                            MatchFiles = MatchFiles + 1
                        Else
                            mLog = mLog & "Group " & k & ": more than one export (" & files(k) & " is read, " & f & " is not)." & vbCrLf
                        End If
                    End If
                End If
            Next k
        End If
        f = Dir$()
    Loop
End Function

Private Function Norm(ByVal s As String) As String
    Norm = UCase$(Replace(Replace(s, " ", ""), ChrW(160), ""))
End Function

Private Function FirstGroupName(gt As Variant, ByVal nG As Long) As String
    Dim k As Long
    For k = 1 To nG
        If Len(CellTxt(gt(k, 2))) > 0 Then FirstGroupName = Replace(CellTxt(gt(k, 2)), " ", ""): Exit Function
    Next k
End Function

' tick list of the groups (PANEL.UiTicks); with an older PANEL a typed list. False on Cancel.
Private Function AskGroups(wb As Workbook, gt As Variant, ByVal nG As Long, files() As String, pick() As Boolean) As Boolean
    Dim items() As String, idx() As Long, n As Long, k As Long, def As String, ans As Variant, s As String, i As Long
    ReDim items(1 To nG): ReDim idx(1 To nG)
    For k = 1 To nG
        If Len(CellTxt(gt(k, 2))) > 0 Then
            n = n + 1: idx(n) = k
            items(n) = k & "   " & CellTxt(gt(k, 2)) & IIf(Len(files(k)) > 0, "", U("   (dosya yok)"))
            def = def & IIf(Len(files(k)) > 0, "1", "0")
        End If
    Next k
    If n = 0 Then Exit Function
    ReDim Preserve items(1 To n)
    On Error Resume Next
    ans = Application.Run("'" & Replace(wb.Name, "'", "''") & "'!PANEL.UiTicks", U("GSD sonu\u00E7lar\u0131 okunacak ayak gruplar\u0131n\u0131 i\u015Faretleyin."), items, MSG_TITLE, def)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        s = UiInputS(U("Okunacak grup numaralar\u0131 (\u00F6rnek: 1,3-5):"), MSG_TITLE, "")
        If Len(Trim$(s)) = 0 Then Exit Function
        ParseList s, nG, pick
        AskGroups = True
        Exit Function
    End If
    On Error GoTo 0
    s = CStr(ans)
    If Len(s) = 0 Then Exit Function
    For i = 1 To n
        If i <= Len(s) Then pick(idx(i)) = (Mid$(s, i, 1) = "1")
    Next i
    AskGroups = True
End Function

' "1,3-5" -> pick(1), pick(3), pick(4), pick(5)
Private Sub ParseList(ByVal s As String, ByVal nG As Long, pick() As Boolean)
    Dim part As Variant, a As Long, b As Long, p As Long, k As Long, t As String
    For Each part In Split(Replace(Replace(s, ";", ","), " ", ""), ",")
        t = CStr(part)
        If Len(t) > 0 Then
            p = InStr(2, t, "-")
            If p > 0 Then
                a = CLng(Val(Left$(t, p - 1))): b = CLng(Val(Mid$(t, p + 1)))
            Else
                a = CLng(Val(t)): b = a
            End If
            For k = a To b
                If k >= 1 And k <= nG Then pick(k) = True
            Next k
        End If
    Next part
End Sub

' ---- sheets ----
' the data sheet of group k: sheet-level name M_TESIR_NO = k (ORTA_AYAK), else an older sheet MIDAS_<k>
Private Function DataSheetOf(wb As Workbook, ByVal k As Long) As Worksheet
    Dim ws As Worksheet, nm As Name, old As Worksheet
    For Each ws In wb.Worksheets
        Set nm = Nothing
        On Error Resume Next
        Set nm = ws.Names(M_DATA_NO)
        On Error GoTo 0
        If Not nm Is Nothing Then
            If nm.RefersTo = "=" & k Then Set DataSheetOf = ws: Exit Function
        ElseIf StrComp(ws.Name, "MIDAS_" & k, vbTextCompare) = 0 Then
            Set old = ws
        End If
    Next ws
    Set DataSheetOf = old
End Function

' the group sheet of group k: the group number in the first cell of its sheet-level M_AYAK_ZINCIR
Private Function GroupSheetOf(wb As Workbook, ByVal k As Long) As Worksheet
    Dim ws As Worksheet, r As Range, v As Variant
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, GROUP_TPL, vbTextCompare) <> 0 Then
            Set r = Nothing
            On Error Resume Next
            Set r = ws.Names(M_CHAIN).RefersToRange
            On Error GoTo 0
            If Not r Is Nothing Then
                If r.Worksheet.Name = ws.Name Then
                    v = r.Cells(1, 1).Value
                    If Not IsError(v) Then
                        If IsNumeric(v) And Not IsEmpty(v) Then
                            If CDbl(v) = k Then Set GroupSheetOf = ws: Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next ws
End Function

Private Function HasSheet(xb As Workbook, ByVal nm As String) As Boolean
    Dim ws As Object
    On Error Resume Next
    Set ws = xb.Worksheets(nm)
    HasSheet = Not ws Is Nothing
End Function

' ---- helpers ----
Private Function CellTxt(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    CellTxt = Trim$(CStr(v))
End Function

Private Function U(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, "\u")
    Do While p > 0
        s = Left$(s, p - 1) & ChrW(CLng("&H" & Mid$(s, p + 2, 4))) & Mid$(s, p + 6)
        p = InStr(p + 1, s, "\u")
    Loop
    U = s
End Function

Private Sub Logp(ByVal msgLine As String)
    If mScripted Then Exit Sub
    On Error Resume Next
    UiLog msgLine
End Sub
