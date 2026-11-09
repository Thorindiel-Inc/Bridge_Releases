Attribute VB_Name = "MIDAS_MOD"
Option Explicit
' ============================================================================
'  MIDAS_MOD - modal results of the seismic model into sheet MOD (chapter 8 workbook)
'
'  Macro window (MAKROLAR) > MIDAS sonuclari > Modal analiz sonuclari, or MIDAS_MOD_OKU.
'  Window field 'islem': 1 read + pictures (default), 2 pictures only (no read: the modes come
'  from MOD, whose Q1 stamp must name the open model), 3 read only; field 'res': picture width
'  4000 / 3000 / 2000 px or screen size. The pictures (Z_YATAY) are taken by MIDAS_SEKIL, called
'  late bound (MIDAS_SEKIL_CAPTURE), so this module also runs where MIDAS_SEKIL is missing.
'  Reads the model OPEN in Civil NX through the MAPI, read only (nothing is written
'  to the model): post/TABLE EIGENVALUEMODE and its sub-tables
'     EIGENVALUE ANALYSIS                      -> MOD C6:J105   (mode, rad/s, Hz, s, tolerance)
'     MODAL PARTICIPATION MASSES PRINTOUT (1)  -> MOD C109:O208 (% mass and sum, TRAN-X..ROTN-Z)
'     MODAL PARTICIPATION MASSES PRINTOUT (2)  -> MOD C211:O310 (mass and sum)
'     MODAL PARTICIPATION FACTOR PRINTOUT      -> MOD C314:N413 (D F H J L N)
'     MODAL DIRECTION FACTOR PRINTOUT          -> MOD C417:N516 (D F H J L N)
'  The addresses are the ones of the old pasted MIDAS text, so 1_RAPOR (Tablo 54, 55,
'  T122 / T123) reads them unchanged. Each block is cleared first (values only, the
'  formats stay); a model with fewer modes leaves the rest of a block empty, a model
'  with more than 100 modes writes the first 100 and says so.
'  Model check: db/PJCF of a MIDAS_KOPRU seismic model ends with SEISMIC; any other
'  model is read only after a yes in the question box.
'  MAPI key: INPUT!T3 (G_MAPI_KEY, pulled by DEGER CAGIR), else MIDAS_MAPI_KEY, else asked.
'  Test entry without Civil NX: MIDAS_MOD_TEST <saved post/TABLE answer .json>.
' ============================================================================

Private Const SCRIPT_ID As String = "midas-modal-read"
Private Const SCRIPT_VERSION As String = "2026-11-09b"
Private Const SCRIPT_CHANGELOG As String = "One tool with the mode-shape pictures: window field Islem = read + pictures, pictures only, read only; field res = picture width (pictures by MIDAS_SEKIL); new: periods, frequencies, mass participation and factors of the seismic model into sheet MOD"

Private Const API_BASE As String = "https://moa-engineers.midasit.com:443/civil/"
Private Const MSG_TITLE As String = "MIDAS_MOD"
Private Const KEY_SHEET As String = "INPUT"
Private Const KEY_CELL As String = "T3"
Private Const KEY_NAME As String = "G_MAPI_KEY"
Private Const OUT_SHEET As String = "MOD"
Private Const N_MAX As Long = 100
Private Const R_EIG As Long = 6, R_PCT As Long = 109, R_MAS As Long = 211, R_FAC As Long = 314, R_DIR As Long = 417
Private Const STAMP_CELL As String = "Q1"
Private Const RES_W As String = "4000,3000,2000,0"  ' window field res: picture width, 0 = screen size

Private mKey As String
Private http As Object

' ============================================================================
Public Sub MIDAS_MOD_OKU()
    Dim resp As String, msg As String, pj As String, mode As Long, pxW As Long, pics As String, ok As Boolean
    On Error GoTo Fail
    ' field islem: 1 read + pictures, 2 pictures only, 3 read only (owner 2026-11-09: one tool, like MIDAS_RTI)
    mode = UiIdx("islem", 1)
    If mode < 1 Or mode > 3 Then mode = 1
    pxW = PicWidth(UiIdx("res", 1))
    If mode = 2 Then
        ' pictures only: MIDAS_SEKIL asks its own questions (model, MOD stamp, start)
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

' pictures by MIDAS_SEKIL.MIDAS_SEKIL_CAPTURE, late bound so this module compiles without it;
' a summary text in any case
Private Function Capture(ByVal pxW As Long, ByVal ask As Boolean, ByVal startQ As Boolean, ByVal modelChecked As Boolean) As String
    Dim r As Variant
    On Error GoTo NoModule
    r = Application.Run("'" & Replace(ThisWorkbook.Name, "'", "''") & "'!MIDAS_SEKIL.MIDAS_SEKIL_CAPTURE", pxW, ask, startQ, modelChecked)
    Capture = CStr(r)
    Exit Function
NoModule:
    Capture = "STOPPED: the pictures need module MIDAS_SEKIL in this workbook (" & Err.Description & ")."
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
    ws.Range(STAMP_CELL).Value = "MIDAS: " & tag & " / " & Format$(Now, "yyyy-mm-dd hh:nn")
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
        "  1_RAPOR T122 Tb = " & CellText("1_RAPOR", "T122") & ", T123 Te = " & CellText("1_RAPOR", "T123") & notes
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

Private Function CellText(ByVal sh As String, ByVal addr As String) As String
    On Error Resume Next
    CellText = CStr(ThisWorkbook.Worksheets(sh).Range(addr).Text)
End Function

' ---- API ----
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
    cellKey = Trim$(CStr(ThisWorkbook.Names(KEY_NAME).RefersToRange.Cells(1, 1).Value))
    If cellKey = "" Then cellKey = Trim$(CStr(ThisWorkbook.Worksheets(KEY_SHEET).Range(KEY_CELL).Value))
    On Error GoTo 0
    If cellKey <> "" Then mKey = cellKey
    If mKey = "" Then mKey = Environ$("MIDAS_MAPI_KEY")
    If mKey = "" Then mKey = UiInputS("MIDAS MAPI key (kept in memory for this Excel session only):", MSG_TITLE)
    GetKey = (mKey <> "")
End Function

Private Function ReadUtf8(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText
    st.Close
End Function

' ---- minimal JSON parser (from MAK_LL): objects -> Scripting.Dictionary, arrays -> 1-based Variant(),
'      numbers -> Double (Val), strings -> String, true/false/null -> Boolean/Null ----
Private Function ParseJson(ByVal s As String) As Object
    Dim p As Long, v As Variant
    p = 1
    SkipWs s, p
    If Mid$(s, p, 1) <> "{" Then Exit Function
    ParseValue s, p, v
    Set ParseJson = v
End Function

Private Sub ParseValue(ByRef s As String, ByRef p As Long, ByRef v As Variant)
    Dim ch As String
    SkipWs s, p
    ch = Mid$(s, p, 1)
    Select Case ch
        Case "{"
            Dim d As Object, k As String, item As Variant
            Set d = CreateObject("Scripting.Dictionary")
            p = p + 1: SkipWs s, p
            If Mid$(s, p, 1) = "}" Then p = p + 1: Set v = d: Exit Sub
            Do
                SkipWs s, p
                k = ParseString(s, p)
                SkipWs s, p
                p = p + 1                       ' :
                ParseValue s, p, item
                If IsObject(item) Then Set d(k) = item Else d(k) = item
                SkipWs s, p
                ch = Mid$(s, p, 1): p = p + 1
                If ch = "}" Then Exit Do
            Loop
            Set v = d
        Case "["
            Dim arr() As Variant, n As Long, el As Variant
            p = p + 1: SkipWs s, p
            If Mid$(s, p, 1) = "]" Then p = p + 1: v = Empty: Exit Sub
            Do
                ParseValue s, p, el
                n = n + 1
                ReDim Preserve arr(1 To n)
                If IsObject(el) Then Set arr(n) = el Else arr(n) = el
                SkipWs s, p
                ch = Mid$(s, p, 1): p = p + 1
                If ch = "]" Then Exit Do
            Loop
            v = arr
        Case """"
            v = ParseString(s, p)
        Case "t": v = True: p = p + 4
        Case "f": v = False: p = p + 5
        Case "n": v = Null: p = p + 4
        Case Else
            Dim st As Long
            st = p
            Do While p <= Len(s) And InStr(1, "+-0123456789.eE", Mid$(s, p, 1), vbBinaryCompare) > 0: p = p + 1: Loop
            v = Val(Mid$(s, st, p - st))
    End Select
End Sub

Private Function ParseString(ByRef s As String, ByRef p As Long) As String
    Dim o As String, ch As String
    p = p + 1                                   ' opening quote
    Do While p <= Len(s)
        ch = Mid$(s, p, 1)
        If ch = """" Then p = p + 1: Exit Do
        If ch = "\" Then
            p = p + 1: ch = Mid$(s, p, 1)
            Select Case ch
                Case "n": o = o & vbLf
                Case "t": o = o & vbTab
                Case "u": o = o & ChrW(CLng("&H" & Mid$(s, p + 1, 4))): p = p + 4
                Case Else: o = o & ch
            End Select
        Else
            o = o & ch
        End If
        p = p + 1
    Loop
    ParseString = o
End Function

Private Sub SkipWs(ByRef s As String, ByRef p As Long)
    Do While p <= Len(s)
        Select Case Mid$(s, p, 1)
            Case " ", vbTab, vbCr, vbLf: p = p + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub
