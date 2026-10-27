Option Explicit

' ============================================================================
'  GLOBAL_CAGIRMA - pull this workbook's INPUT values from GLOBAL_INPUT.xlsx
' ----------------------------------------------------------------------------
'  Replaces DEGER_CAGIRMA / KIRIS_CAGIRMA (hard-coded ranges, one divergent
'  copy per workbook). Here the cells to fill come from the TARGETS sheet of
'  the global file, matched BY KEY, so layout differences between workbooks
'  no longer matter.
'
'  NAME THE MODULE something else than GLOBAL_CAGIRMA (for example
'  modGlobalCagirma): a module with the same name as its macro breaks
'  Application.Run and button assignments. The updater finds it by SCRIPT_ID.
'
'  WHAT IT DOES
'   1. Workbook id = file name up to the first underscore (4A-7A, 8'6, 10).
'   2. Opens GLOBAL_INPUT.xlsx read-only (path asked once, then kept in the
'      workbook's custom property GLOBAL_INPUT_PATH - never in a cell). If the
'      file is already open it uses that open copy (unsaved edits included),
'      says so, and leaves it open.
'   3. Bridge = WORKBOOKS sheet default, or the one you type (GLOBAL_CAGIRMA_SEC).
'   4. For every TARGETS row of this workbook: writes the BRIDGES value of
'      that bridge as a VALUE (Value2: numbers are never converted through
'      text, so a decimal comma cannot corrupt them).
'
'  RULES
'   * a TYPED cell that holds a FORMULA is skipped and reported
'   * a DERIVED cell is meant to become a typed value: its formula IS replaced;
'     the report counts and lists these (also in the dry run) - there is no undo
'   * a blank (or empty-string) value in the global file never blanks a cell
'   * a missing key / sheet / cell, a multi-cell address, an error value and a
'     failed write are reported, never skipped silently
'   * GLOBAL_CAGIRMA_KONTROL does everything except the write (dry run)
'   * the global file is never saved; events stay off while the pull runs
'     (none of the 16 workbooks has sheet or workbook event code)
' ============================================================================

Private Const SCRIPT_VERSION As String = "2026-10-27b"
Private Const SCRIPT_CHANGELOG As String = "Review fixes: reuse an already open global file, skip empty strings, list replaced DERIVED formulas, guard multi-cell and failed writes, manual calc while writing"
Private Const SCRIPT_ID As String = "global-input-pull"

Private Const PROP_PATH As String = "GLOBAL_INPUT_PATH"
Private Const BRIDGES_HEADER_ROW As Long = 4
Private Const MAX_LISTED As Long = 10
Private Const MARKER As String = "###"

' Entry 1: default bridge of this workbook.
Public Sub GLOBAL_CAGIRMA()
    Call PullGlobalInput("", False)
End Sub

' Entry 2: ask which bridge.
Public Sub GLOBAL_CAGIRMA_SEC()
    Dim b As String
    b = InputBox("Bridge column of GLOBAL_INPUT (BRIDGES sheet header), e.g. KP-019:", "GLOBAL_CAGIRMA")
    If Len(Trim$(b)) = 0 Then Exit Sub
    Call PullGlobalInput(Trim$(b), False)
End Sub

' Entry 3: dry run, default bridge. Writes nothing.
Public Sub GLOBAL_CAGIRMA_KONTROL()
    Call PullGlobalInput("", True)
End Sub

' Entry 4: forget the stored path of the global file.
Public Sub GLOBAL_CAGIRMA_YOL_SIL()
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties(PROP_PATH).Delete
    On Error GoTo 0
    MsgBox "Stored path removed. The next pull asks for GLOBAL_INPUT.xlsx again.", vbInformation, "GLOBAL_CAGIRMA"
End Sub

Private Sub PullGlobalInput(ByVal bridgeOverride As String, ByVal dryRun As Boolean)

    Dim gWb As Workbook
    Dim wsBr As Worksheet, wsTg As Worksheet, wsWb As Worksheet
    Dim path As String, wbId As String, bridge As String
    Dim keyRow As Object
    Dim i As Long, lastRow As Long, bridgeCol As Long
    Dim tg As Variant
    Dim key As String, sh As String, addr As String, kind As String
    Dim v As Variant
    Dim target As Range
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim nWritten As Long, nTargets As Long, nReplaced As Long
    Dim nFormula As Long, nBlank As Long, nNoKey As Long, nNoSheet As Long, nError As Long, nBadCell As Long, nWriteFail As Long
    Dim issues As String, nIssues As Long
    Dim replacedList As String, nListed As Long
    Dim stage As String
    Dim report As String, msg As String
    Dim wasOpen As Boolean, quiet As Boolean, pathIsNew As Boolean
    Dim prevScreen As Boolean, prevEvents As Boolean, prevCalc As Long

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    On Error GoTo Fail

    stage = "workbook id"
    wbId = WorkbookId(ThisWorkbook.Name)

    stage = "find the global file"
    path = GlobalPath(pathIsNew)
    If Len(path) = 0 Then Exit Sub
    If StrComp(path, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        MsgBox "The selected file is this workbook itself. Select GLOBAL_INPUT.xlsx.", vbExclamation, "GLOBAL_CAGIRMA"
        If pathIsNew Then Call ForgetPath
        Exit Sub
    End If

    stage = "open the global file"
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    For Each wb In Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then
            Set gWb = wb
            wasOpen = True
        End If
    Next wb
    If gWb Is Nothing Then Set gWb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)

    stage = "read the global sheets"
    Set wsBr = gWb.Worksheets("BRIDGES")
    Set wsTg = gWb.Worksheets("TARGETS")
    Set wsWb = gWb.Worksheets("WORKBOOKS")
    If pathIsNew Then Call StorePath(path)
    stage = "calculate the global sheet"
    wsBr.Calculate

    stage = "choose the bridge"
    bridge = bridgeOverride
    If Len(bridge) = 0 Then bridge = DefaultBridge(wsWb, wbId)
    If Len(bridge) = 0 Then
        MsgBox "Workbook id """ & wbId & """ is not in the WORKBOOKS sheet (or has no default bridge)." & vbCrLf & _
               "Use GLOBAL_CAGIRMA_SEC to type a bridge.", vbExclamation, "GLOBAL_CAGIRMA"
        quiet = True
        GoTo Done
    End If
    bridgeCol = 0
    For i = 3 To wsBr.Cells(BRIDGES_HEADER_ROW, wsBr.Columns.Count).End(xlToLeft).Column
        If StrComp(SafeStr(wsBr.Cells(BRIDGES_HEADER_ROW, i).Value2), bridge, vbTextCompare) = 0 Then
            bridgeCol = i
            Exit For
        End If
    Next i
    If bridgeCol = 0 Then
        MsgBox "Bridge """ & bridge & """ is not a column of the BRIDGES sheet.", vbExclamation, "GLOBAL_CAGIRMA"
        quiet = True
        GoTo Done
    End If

    stage = "index the keys"
    Set keyRow = CreateObject("Scripting.Dictionary")
    keyRow.CompareMode = 1
    lastRow = wsBr.Cells(wsBr.Rows.Count, 1).End(xlUp).Row
    For i = BRIDGES_HEADER_ROW + 1 To lastRow
        key = SafeStr(wsBr.Cells(i, 1).Value2)
        If Len(key) > 0 Then
            If Left$(key, 3) <> MARKER Then
                If keyRow.Exists(key) Then
                    Call AddIssue(issues, nIssues, "duplicate key in BRIDGES (first row used): " & key)
                Else
                    keyRow.Add key, i
                End If
            End If
        End If
    Next i

    stage = "read TARGETS"
    lastRow = wsTg.Cells(wsTg.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "TARGETS is empty.", vbExclamation, "GLOBAL_CAGIRMA"
        quiet = True
        GoTo Done
    End If
    tg = wsTg.Range(wsTg.Cells(2, 1), wsTg.Cells(lastRow, 5)).Value2

    stage = "write the values"
    If Not dryRun Then Application.Calculation = xlCalculationManual
    For i = 1 To UBound(tg, 1)
        If StrComp(SafeStr(tg(i, 2)), wbId, vbTextCompare) = 0 Then
            nTargets = nTargets + 1
            key = SafeStr(tg(i, 1))
            sh = SafeStr(tg(i, 3))
            addr = SafeStr(tg(i, 4))
            kind = SafeStr(tg(i, 5))

            If Not keyRow.Exists(key) Then
                nNoKey = nNoKey + 1
                Call AddIssue(issues, nIssues, "key not in BRIDGES: " & key)
                GoTo NextTarget
            End If
            Set ws = Nothing
            On Error Resume Next
            Set ws = ThisWorkbook.Worksheets(sh)
            On Error GoTo Fail
            If ws Is Nothing Then
                nNoSheet = nNoSheet + 1
                Call AddIssue(issues, nIssues, "sheet missing: " & sh & " (" & key & ")")
                GoTo NextTarget
            End If
            Set target = Nothing
            On Error Resume Next
            Set target = ws.Range(addr)
            On Error GoTo Fail
            If target Is Nothing Then
                nBadCell = nBadCell + 1
                Call AddIssue(issues, nIssues, "bad cell: " & sh & "!" & addr)
                GoTo NextTarget
            End If
            If target.Cells.CountLarge <> 1 Then
                nBadCell = nBadCell + 1
                Call AddIssue(issues, nIssues, "not a single cell: " & sh & "!" & addr)
                GoTo NextTarget
            End If
            If target.MergeCells Then
                If target.Address <> target.MergeArea.Cells(1, 1).Address Then
                    nBadCell = nBadCell + 1
                    Call AddIssue(issues, nIssues, "merged, not the top-left cell: " & sh & "!" & addr)
                    GoTo NextTarget
                End If
            End If

            v = wsBr.Cells(CLng(keyRow(key)), bridgeCol).Value2
            If IsError(v) Then
                nError = nError + 1
                Call AddIssue(issues, nIssues, "error value in global file for " & key)
                GoTo NextTarget
            End If
            If IsEmpty(v) Then
                nBlank = nBlank + 1
                GoTo NextTarget
            End If
            If VarType(v) = vbString Then
                If Len(v) = 0 Then
                    nBlank = nBlank + 1
                    GoTo NextTarget
                End If
            End If
            If target.HasFormula Then
                If StrComp(kind, "DERIVED", vbTextCompare) <> 0 Then
                    nFormula = nFormula + 1
                    Call AddIssue(issues, nIssues, "formula kept (listed TYPED): " & sh & "!" & addr & " " & key)
                    GoTo NextTarget
                End If
                nReplaced = nReplaced + 1
                If nListed < MAX_LISTED Then
                    replacedList = replacedList & "  " & sh & "!" & addr & " " & key & vbCrLf
                    nListed = nListed + 1
                End If
            End If

            If Not dryRun Then
                ' text that Excel would re-parse (number, date, formula) is stored as text
                If VarType(v) = vbString Then
                    If IsNumeric(v) Or IsDate(v) Or Left$(v, 1) = "=" Then v = "'" & v
                End If
                On Error Resume Next
                target.Value2 = v
                If Err.Number <> 0 Then
                    nWriteFail = nWriteFail + 1
                    Call AddIssue(issues, nIssues, "write failed " & sh & "!" & addr & " " & key & ": " & Err.Description)
                    Err.Clear
                    On Error GoTo Fail
                    GoTo NextTarget
                End If
                On Error GoTo Fail
            End If
            nWritten = nWritten + 1
        End If
NextTarget:
    Next i

Done:
    Application.Calculation = prevCalc
    If Not gWb Is Nothing Then
        If Not wasOpen Then gWb.Close SaveChanges:=False
    End If
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen

    If quiet Then Exit Sub
    If nTargets = 0 Then
        MsgBox "No TARGETS rows for workbook id """ & wbId & """.", vbInformation, "GLOBAL_CAGIRMA"
        Exit Sub
    End If

    If dryRun Then report = "DRY RUN (nothing written)" Else report = "Pull done"
    report = report & "  [" & SCRIPT_VERSION & "]" & vbCrLf & _
             "workbook " & wbId & "   bridge " & bridge & vbCrLf & _
             "global file: " & path & vbCrLf & String(40, "-") & vbCrLf & _
             "targets: " & nTargets & vbCrLf
    If wasOpen Then report = report & "NOTE: the global file was already open - its current (maybe unsaved) values were used" & vbCrLf
    If dryRun Then report = report & "would write: " & nWritten & vbCrLf Else report = report & "written: " & nWritten & vbCrLf
    report = report & "DERIVED formula cells replaced by values: " & nReplaced & vbCrLf & _
             "blank in global file (left alone): " & nBlank & vbCrLf & _
             "formula kept: " & nFormula & "   write failed: " & nWriteFail & vbCrLf & _
             "key missing: " & nNoKey & "   sheet missing: " & nNoSheet & "   bad cell: " & nBadCell & "   error value: " & nError
    If nReplaced > 0 Then report = report & vbCrLf & vbCrLf & "Formulas replaced (first " & nListed & "):" & vbCrLf & replacedList
    If nIssues > 0 Then report = report & vbCrLf & "First issues:" & vbCrLf & issues
    Debug.Print report
    If nIssues > 0 Then
        MsgBox report, vbExclamation, "GLOBAL_CAGIRMA"
    Else
        MsgBox report, vbInformation, "GLOBAL_CAGIRMA"
    End If
    Exit Sub

Fail:
    msg = "GLOBAL_CAGIRMA stopped at step '" & stage & "': " & Err.Number & " " & Err.Description
    Debug.Print msg
    On Error Resume Next
    Application.Calculation = prevCalc
    If Not gWb Is Nothing Then
        If Not wasOpen Then gWb.Close SaveChanges:=False
    End If
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    MsgBox msg & vbCrLf & vbCrLf & "Written so far: " & nWritten, vbCritical, "GLOBAL_CAGIRMA"
End Sub

Private Sub AddIssue(ByRef issues As String, ByRef nIssues As Long, ByVal text As String)
    nIssues = nIssues + 1
    Debug.Print "issue: " & text
    If nIssues <= MAX_LISTED Then issues = issues & "  " & text & vbCrLf
    If nIssues = MAX_LISTED + 1 Then issues = issues & "  ... (all issues are in the Immediate window)" & vbCrLf
End Sub

' CStr that survives error values and Null.
Private Function SafeStr(ByVal x As Variant) As String
    If IsError(x) Then Exit Function
    If IsNull(x) Then Exit Function
    SafeStr = CStr(x)
End Function

' "4A-7A_YHT_XXXXX_UST_KDME_L30.xlsm" -> "4A-7A"
Private Function WorkbookId(ByVal fileName As String) As String
    Dim p As Long
    p = InStr(1, fileName, "_", vbBinaryCompare)
    If p > 1 Then
        WorkbookId = Left$(fileName, p - 1)
    Else
        WorkbookId = fileName
    End If
End Function

Private Function DefaultBridge(ByVal wsWb As Worksheet, ByVal wbId As String) As String
    Dim i As Long, lastRow As Long
    lastRow = wsWb.Cells(wsWb.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        If StrComp(SafeStr(wsWb.Cells(i, 1).Value2), wbId, vbTextCompare) = 0 Then
            DefaultBridge = SafeStr(wsWb.Cells(i, 3).Value2)
            Exit Function
        End If
    Next i
End Function

' Stored path of GLOBAL_INPUT.xlsx, or ask once. isNew = True when just picked
' (stored only after the file proved to have the three sheets).
Private Function GlobalPath(ByRef isNew As Boolean) As String
    Dim p As String
    Dim res As Variant
    isNew = False
    On Error Resume Next
    p = CStr(ThisWorkbook.CustomDocumentProperties(PROP_PATH).Value)
    On Error GoTo 0
    If Len(p) > 0 Then
        If Len(Dir$(p)) > 0 Then
            GlobalPath = p
            Exit Function
        End If
    End If
    res = Application.GetOpenFilename("Excel Files (*.xlsx;*.xlsm), *.xlsx;*.xlsm", , "Select GLOBAL_INPUT.xlsx")
    If VarType(res) = vbBoolean Then Exit Function
    isNew = True
    GlobalPath = CStr(res)
End Function

Private Sub StorePath(ByVal p As String)
    Call ForgetPath
    ' 4 = msoPropertyTypeString (literal: no Office library reference needed)
    ThisWorkbook.CustomDocumentProperties.Add Name:=PROP_PATH, LinkToContent:=False, Type:=4, Value:=p
End Sub

Private Sub ForgetPath()
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties(PROP_PATH).Delete
    On Error GoTo 0
End Sub
