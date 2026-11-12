Attribute VB_Name = "MAK_LL"
Option Explicit
' ============================================================================
'  MAK_LL - live-load envelope of the girders for 4-7 Tablo 27, read from MIDAS
'
'  Button on sheet LL (OnAction MAK_LL.MAK_LL). Reads the model OPEN in Civil NX
'  through the MAPI (read only - nothing is written to the model):
'    - girder elements = elements with a COMPOSITE section, girder nodes at Z = 0
'    - one moving-load case, picked in a pop-up from the model's cases
'    - beam forces (MV:max) and (MV:min) at both ends of every girder element
'  For each Tablo 27 station (4-7 1_RAPOR, column K, m FROM THE BEARING) the
'  girder node at x = bearing + K and the mirrored node (L - bearing - K) are used
'  on every girder (owner 2026-11-03: all girders, one envelope, both halves):
'    M = max(|My max|, |My min|), V = max(|Vz max|, |Vz min|) over both element
'    ends that meet at the node. No interpolation: a station without a node stops
'    the run (MIDAS_USTYAPI 2026-11-03s adds the station nodes).
'  Writes M to LL!X4.. and V to LL!Y4.. (read by 1_RAPOR Tablo 27) and a detail
'  table + run info on sheet LL. The old pasted MIDAS table and its helper
'  columns are cleared on the first run (owner 2026-11-03: new clean layout).
'  Bearing x = 4-7 G_SARKIKLIK (cm from the girder end).
'  MAPI key: read by MIDAS_CORE (name G_MAPI_KEY, then INPUT!T3, then MIDAS_MAPI_KEY, then asked; memory only).
' ============================================================================

Private Const SCRIPT_ID As String = "mak-ll"
Private Const SCRIPT_VERSION As String = "2026-11-12b"
Private Const SCRIPT_CHANGELOG As String = "Its macro-window row is declared in this module ('@TOOL line), so a changed or new tool needs no PANEL release; Shared helpers (key, HTTP, JSON, files) are in the module MIDAS_CORE, which UPDATE adds; MAPI key order: G_MAPI_KEY, INPUT!T3, MIDAS_MAPI_KEY, prompt; Refuses the DYNAMIC model; Syntax error in the last table line fixed; Moving-load case from a drop-down of the model cases, default LM71; Inputs and messages through the macro window; Span sheets named 4A_, 4B_ ...; Table below the strand tables on 4<letter>_KILIF, reads 4<letter>_KDME"
Private Const SCRIPT_NEEDS As String = "midas-core"   ' UPDATE adds this module when the workbook lacks it
' macro-window row(s) of this module (PANEL reads them): order|group|title|macro|sheet|text|fields[|modules]
'@TOOL 100|Kiri\u015F tasar\u0131m\u0131|MIDAS'tan hareketli y\u00FCk|MAK_LL.MAK_LL|4?_KILIF|Civil NX'te a\u00E7\u0131k a\u015Famal\u0131 modelden hareketli y\u00FCk kuvvetlerini okur, a\u00E7\u0131kl\u0131\u011F\u0131n LL tablosuna yazar (Tablo 27).|

Private Const MSG_TITLE As String = "MAK_LL"
Private Const LL_TOP As Long = 72          ' title row of the table on the 4<letter>_KILIF sheet
Private Const OUT_ROW As Long = 75          ' first station row of X (M) / Y (V), read by 1_RAPOR
Private Const COL_M As Long = 15           ' X
Private Const COL_V As Long = 16           ' Y
Private Const TOL As Double = 0.001        ' node position tolerance (m)
Private Const DEF_CASE As String = "LM71"  ' preselected moving-load case


' station data
Private nSt As Long, stName() As String, stK() As Double, stX1() As Double, stX2() As Double
Private resM() As Double, resV() As Double
Private govM() As String, govV() As String

' ============================================================================
Public Sub MAK_LL()
    Dim ws As Worksheet, wb As Workbook, calcWas As Long, msg As String
    On Error GoTo Fail
    Set wb = ThisWorkbook
    Set ws = ActiveSheet
    If UCase$(Right$(ws.Name, 6)) <> "_KILIF" Then Err.Raise vbObjectError + 9, , "Run MAK_LL from a span sheet such as 4A_KILIF."
    calcWas = Application.Calculation
    msg = RunEnvelope(wb, ws, "", "")
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    If msg <> "" Then UiMsg msg, IIf(Left$(msg, 4) = "Done", vbInformation, vbExclamation), MSG_TITLE
    Exit Sub
Fail:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    UiMsg "Stopped: " & Err.Description & vbCrLf & "The LL table was not changed.", vbCritical, MSG_TITLE
End Sub

' Scripted entry (tests): no dialogs; caseName given; report written to reportFile.
Public Sub MAK_LL_RUN(ByVal caseName As String, ByVal reportFile As String, Optional ByVal apiKey As String = "", Optional ByVal llSheet As String = "4A_KILIF")
    Dim msg As String
    On Error GoTo Fail
    If apiKey <> "" Then McSetKey apiKey
    msg = RunEnvelope(ThisWorkbook, ThisWorkbook.Worksheets(llSheet), caseName, reportFile)
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    WriteText reportFile, msg
    Exit Sub
Fail:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    WriteText reportFile, "FAIL: " & Err.Description
End Sub

Public Sub MAK_LL_FORGET_KEY()
    McForgetKey
End Sub

' ============================================================================
Private Function RunEnvelope(wb As Workbook, ws As Worksheet, ByVal caseName As String, ByVal reportFile As String) As String
    Dim scripted As Boolean, resp As String, sark As Double
    Dim nodes As Object, elems As Object, sects As Object, compSect As Object
    Dim gElem As Collection, gNodeX As Object, gLines As Object, Lg As Double
    Dim k As Variant, e As Object, nd As Variant, i As Long, j As Long
    Dim ni As Long, nj As Long, xi As Double, xj As Double, yi As Double
    Dim cases As Collection, keys As String, rowsMax As Variant, rowsMin As Variant

    scripted = (reportFile <> "")
    ReadStations wb, ws
    sark = Num(NmRange(wb, "G_SARKIKLIK").Cells(1).Value, "4-7 G_SARKIKLIK") / 100#
    If sark <= 0 Or sark > 2 Then Err.Raise vbObjectError + 1, , "4-7 G_SARKIKLIK = " & sark * 100 & " cm is out of range."

    If Not GetKey(wb) Then RunEnvelope = "Stopped: no MAPI key (G_MAPI_KEY, INPUT!T3, MIDAS_MAPI_KEY or the prompt).": Exit Function
    McCloseHttp
    Application.StatusBar = "MAK_LL: reading the model from Civil NX..."
    If IsDynamicModel() Then
        RunEnvelope = "Stopped: the model open in Civil NX is the DYNAMIC model. Its supports hold X at slab level and give false moments at the bearings." & vbCrLf & "Open the staged model and run MAK_LL again. The table was not changed."
        Exit Function
    End If
    Set nodes = TableGet("NODE")
    If nodes Is Nothing Then RunEnvelope = "Stopped: Civil NX does not answer or the model has no nodes (is the API connected?).": Exit Function
    Set elems = TableGet("ELEM")
    Set sects = TableGet("SECT")
    If elems Is Nothing Or sects Is Nothing Then RunEnvelope = "Stopped: could not read elements / sections from Civil NX.": Exit Function

    ' girder elements: COMPOSITE sections, both nodes at Z = 0, along X
    Set compSect = CreateObject("Scripting.Dictionary")
    For Each k In sects.keys
        If UCase$(CStr(GetField(sects(k), "SECTTYPE"))) = "COMPOSITE" Then compSect(CStr(k)) = True
    Next k
    If compSect.Count = 0 Then RunEnvelope = "Stopped: the model has no COMPOSITE section - is this the USTYAPI superstructure model?": Exit Function
    Set gElem = New Collection
    Set gNodeX = CreateObject("Scripting.Dictionary")      ' node -> x of girder nodes
    Set gLines = CreateObject("Scripting.Dictionary")      ' girder line y -> count
    For Each k In elems.keys
        Set e = elems(k)
        If compSect.Exists(CStr(GetField(e, "SECT"))) Then
            nd = GetField(e, "NODE")
            ni = CLng(nd(1)): nj = CLng(nd(2))
            If Abs(NodeVal(nodes, ni, "Z")) < TOL And Abs(NodeVal(nodes, nj, "Z")) < TOL Then
                gElem.Add CLng(k)
                xi = NodeVal(nodes, ni, "X"): xj = NodeVal(nodes, nj, "X"): yi = NodeVal(nodes, ni, "Y")
                gNodeX(ni) = xi: gNodeX(nj) = xj
                gLines(Format(yi, "0.000")) = 1
                If xi > Lg Then Lg = xi
                If xj > Lg Then Lg = xj
            End If
        End If
    Next k
    If gElem.Count = 0 Then RunEnvelope = "Stopped: no girder elements (COMPOSITE sections at Z = 0) found.": Exit Function

    ' station x in the model, both halves
    For i = 1 To nSt
        stX1(i) = sark + stK(i)
        stX2(i) = Lg - sark - stK(i)
        If stX1(i) > Lg / 2 + TOL Then Err.Raise vbObjectError + 2, , "Station " & stName(i) & " (" & stK(i) & " m from the bearing) lies beyond midspan (girder " & Lg & " m)."
    Next i

    ' moving-load case
    If caseName = "" Then
        Set cases = MovingCases()
        If cases.Count = 0 Then RunEnvelope = "Stopped: the model has no moving-load case.": Exit Function
        caseName = PickCase(wb, cases)
        If caseName = "" Then RunEnvelope = "": Exit Function
    End If

    ' beam forces of all girder elements, both ends
    For i = 1 To gElem.Count: keys = keys & IIf(i = 1, "", ",") & gElem(i): Next i
    Application.StatusBar = "MAK_LL: reading " & caseName & " (max)..."
    rowsMax = ForceRows(keys, caseName & "(MV:max)")
    Application.StatusBar = "MAK_LL: reading " & caseName & " (min)..."
    rowsMin = ForceRows(keys, caseName & "(MV:min)")

    ' envelope per station
    Dim hits As Long, missTxt As String, lineCnt As Object, y As Double
    For i = 1 To nSt
        resM(i) = -1: resV(i) = -1: govM(i) = "": govV(i) = ""
        Set lineCnt = CreateObject("Scripting.Dictionary")
        For Each k In gNodeX.keys
            If Abs(gNodeX(k) - stX1(i)) < TOL Or Abs(gNodeX(k) - stX2(i)) < TOL Then
                y = NodeVal(nodes, CLng(k), "Y")
                lineCnt(Format(y, "0.000")) = 1
                Envelope rowsMax, CLng(k), "max", i, gNodeX(k), y
                Envelope rowsMin, CLng(k), "min", i, gNodeX(k), y
            End If
        Next k
        If lineCnt.Count < gLines.Count Then
            missTxt = missTxt & "  " & stName(i) & " (x = " & Format(stX1(i), "0.000") & " m): node on " & lineCnt.Count & " of " & gLines.Count & " girders" & vbCrLf
        ElseIf resM(i) < 0 Then
            missTxt = missTxt & "  " & stName(i) & ": no " & caseName & " result at its nodes" & vbCrLf
        End If
    Next i
    If missTxt <> "" Then
        RunEnvelope = "Stopped - LL table not changed. Stations without a girder node / result:" & vbCrLf & missTxt & _
            "Rebuild the model with MIDAS_USTYAPI 2026-11-03s or later (it adds nodes at the Tablo 27 stations) and run the analysis."
        Exit Function
    End If

    WriteSheet ws, caseName, gLines.Count, gElem.Count, Lg, sark
    RunEnvelope = "Done: " & caseName & ", " & gLines.Count & " girders, " & nSt & " stations." & vbCrLf & _
        "M max " & Format(MaxOf(resM), "0.0") & " kNm, V max " & Format(MaxOf(resV), "0.0") & " kN written to O" & OUT_ROW & ":P" & (OUT_ROW + nSt - 1) & "." & vbCrLf & _
        SCRIPT_ID & " " & SCRIPT_VERSION
End Function

' rows: array of Variant(Elem, Load, Part, ..., Shear-z, ..., Moment-y); Part "I[41]" / "J[42]" = node
Private Sub Envelope(rows As Variant, ByVal nodeNo As Long, ByVal sideTxt As String, ByVal i As Long, ByVal x As Double, ByVal y As Double)
    Dim r As Long, tag As String, v As Double, m As Double, part As String
    If IsEmpty(rows) Then Exit Sub
    tag = "[" & nodeNo & "]"
    For r = LBound(rows) To UBound(rows)
        part = CStr(rows(r)(2))
        If InStr(1, part, tag, vbBinaryCompare) > 0 Then
            v = Abs(rows(r)(4)): m = Abs(rows(r)(5))
            If m > resM(i) Then resM(i) = m: govM(i) = "girder y " & Format(y, "0.000") & ", x " & Format(x, "0.000") & ", elem " & rows(r)(0) & " " & Left$(part, 1) & ", " & sideTxt
            If v > resV(i) Then resV(i) = v: govV(i) = "girder y " & Format(y, "0.000") & ", x " & Format(x, "0.000") & ", elem " & rows(r)(0) & " " & Left$(part, 1) & ", " & sideTxt
        End If
    Next r
End Sub

' ---- sheet ----
Private Sub WriteSheet(ws As Worksheet, ByVal caseName As String, ByVal nGird As Long, ByVal nEl As Long, ByVal Lg As Double, ByVal sark As Double)
    Dim i As Long, r As Long, wasProt As Boolean
    Application.Calculation = xlCalculationManual
    ' clean layout (owner 2026-11-03): clear everything except the button; X/Y rewritten below
    ws.Range(ws.Cells(LL_TOP, 2), ws.Cells(LL_TOP + 36, COL_V)).ClearContents
    ws.Cells(LL_TOP, 2).Value = "MAK_LL - live-load envelope of the girders (read from MIDAS)"
    ws.Cells(LL_TOP, 2).Font.Bold = True
    ws.Cells(LL_TOP + 1, 2).Value = "Case: " & caseName & "   |   " & nGird & " girders, " & nEl & " girder elements, girder " & Format(Lg, "0.000") & " m, bearing at " & Format(sark, "0.000") & " m   |   " & Format(Now, "yyyy-mm-dd hh:nn") & "   |   " & SCRIPT_ID & " " & SCRIPT_VERSION
    r = OUT_ROW - 1
    ws.Cells(r, 2).Resize(1, 7).Value = Array("Station", "x from bearing (m)", "Model x (m)", "Model x mirror (m)", "|M| max (kNm)", "|V| max (kN)", "")
    ws.Cells(r, 9).Value = "Governing M"
    ws.Cells(r, 14).Value = "Governing V"
    ws.Cells(r, COL_M).Value = "M"
    ws.Cells(r, COL_V).Value = "V"
    ws.Range(ws.Cells(r, 2), ws.Cells(r, COL_V)).Font.Bold = True
    For i = 1 To nSt
        r = OUT_ROW + i - 1
        ws.Cells(r, 2).Value = stName(i)
        ws.Cells(r, 3).Value = stK(i)
        ws.Cells(r, 4).Value = Round(stX1(i), 4)
        ws.Cells(r, 5).Value = Round(stX2(i), 4)
        ws.Cells(r, 6).Value = Round(resM(i), 2)
        ws.Cells(r, 7).Value = Round(resV(i), 2)
        ws.Cells(r, 9).Value = govM(i)
        ws.Cells(r, 14).Value = govV(i)
        ws.Cells(r, COL_M).Value = Round(resM(i), 2)
        ws.Cells(r, COL_V).Value = Round(resV(i), 2)
    Next i
    r = OUT_ROW + nSt + 1
    ws.Cells(r, 2).Value = "M = max(|My(MV:max)|, |My(MV:min)|), V = max(|Vz(MV:max)|, |Vz(MV:min)|) at the station nodes of every girder, both element ends, both halves (x and L - x)."
    ws.Cells(r + 1, 2).Value = "O" & OUT_ROW & ":P" & (OUT_ROW + nSt - 1) & " are read by the report Tablo 27. Stations = 1_RAPOR Tablo 27 column K (m from the bearing)."
    Application.Calculation = xlCalculationAutomatic
End Sub

' ---- 4-7 1_RAPOR Tablo 27 stations: header 'Yeri' (H) + 'V (kN)' (N); rows below until D is empty ----
Private Sub ReadStations(wb As Workbook, ll As Worksheet)
    Dim ws As Worksheet, r As Long, hdr As Long, last As Long
    Set ws = wb.Worksheets(Left$(ll.Name, Len(ll.Name) - 6) & "_KDME")
    last = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row
    For r = 1 To last
        If Left$(Txt(ws.Cells(r, 8).Value), 4) = "Yeri" And Left$(Txt(ws.Cells(r, 14).Value), 6) = "V (kN)" Then hdr = r: Exit For
    Next r
    If hdr = 0 Then Err.Raise vbObjectError + 3, , "4-7 1_RAPOR: Tablo 27 header ('Yeri' in column H, 'V (kN)' in column N) not found."
    nSt = 0
    ReDim stName(1 To 30): ReDim stK(1 To 30): ReDim stX1(1 To 30): ReDim stX2(1 To 30)
    ReDim resM(1 To 30): ReDim resV(1 To 30): ReDim govM(1 To 30): ReDim govV(1 To 30)
    For r = hdr + 1 To hdr + 30
        If Txt(ws.Cells(r, 4).Value) = "" Then Exit For
        nSt = nSt + 1
        stName(nSt) = Txt(ws.Cells(r, 4).Value)
        stK(nSt) = Num(ws.Cells(r, 11).Value, "4-7 1_RAPOR Tablo 27 station x (K" & r & ")")
    Next r
    If nSt = 0 Then Err.Raise vbObjectError + 3, , "4-7 1_RAPOR: Tablo 27 has no stations."
    ReDim Preserve resM(1 To nSt): ReDim Preserve resV(1 To nSt)
End Sub

' ---- moving-load cases ----
Private Function MovingCases() As Collection
    Dim c As New Collection, t As Variant, d As Object, k As Variant, nm As String, seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    For Each t In Array("MVLDeu", "MVLD")
        Set d = TableGet(CStr(t))
        If Not d Is Nothing Then
            For Each k In d.keys
                nm = CStr(GetField(d(k), "LCNAME"))
                If nm <> "" And Not seen.Exists(nm) Then seen(nm) = True: c.Add nm
            Next k
        End If
    Next t
    Set MovingCases = c
End Function

Private Function PickCase(wb As Workbook, cases As Collection) As String
    ' drop-down of the moving-load cases found in the open model; default LM71 (owner 2026-11-06),
    ' else the third case, else the first
    Dim items() As String, i As Long, v As Long, def As Long
    ReDim items(0 To cases.Count - 1)
    def = IIf(cases.Count >= 3, 3, 1)
    For i = 1 To cases.Count
        items(i - 1) = cases(i)
        If cases(i) = DEF_CASE Then def = i
    Next i
    v = UiChoice("Moving-load case for the girder live-load envelope:", items, MSG_TITLE, def)
    If v > 0 Then PickCase = cases(v)
End Function

' ---- beam forces: rows (Elem, Load, Part, Axial-unused, Shear-z, Moment-y) ----
Private Function ForceRows(ByVal keys As String, ByVal lc As String) As Variant
    Dim body As String, resp As String, root As Object, tb As Object, head As Variant, data As Variant
    Dim iE As Long, iL As Long, iP As Long, iV As Long, iM As Long, j As Long, r As Long, out() As Variant
    body = "{""Argument"":{""TABLE_NAME"":""F"",""TABLE_TYPE"":""BEAMFORCE"",""UNIT"":{""FORCE"":""KN"",""DIST"":""M""}," & _
           """NODE_ELEMS"":{""KEYS"":[" & keys & "]},""LOAD_CASE_NAMES"":[" & QS(lc) & "],""PARTS"":[""PartI"",""PartJ""]}}"
    If Not Api("POST", "post/TABLE", body, resp) Then Err.Raise vbObjectError + 4, , "Beam forces " & lc & ": " & Left$(resp, 300)
    Set root = ParseJson(resp)
    If Not root.Exists("F") Then Err.Raise vbObjectError + 4, , "Beam forces " & lc & ": no table (has the moving-load analysis been run?). " & Left$(resp, 200)
    Set tb = root("F")
    head = tb("HEAD"): data = tb("DATA")
    For j = 1 To UBound(head)
        Select Case CStr(head(j))
            Case "Elem": iE = j
            Case "Load": iL = j
            Case "Part": iP = j
            Case "Shear-z": iV = j
            Case "Moment-y": iM = j
        End Select
    Next j
    If iE * iP * iV * iM = 0 Then Err.Raise vbObjectError + 4, , "Beam forces " & lc & ": unexpected columns."
    If IsEmpty(data) Then Exit Function
    ReDim out(1 To UBound(data))
    For r = 1 To UBound(data)
        out(r) = Array(CStr(data(r)(iE)), CStr(data(r)(iL)), CStr(data(r)(iP)), 0, Val(CStr(data(r)(iV))), Val(CStr(data(r)(iM))))
    Next r
    ForceRows = out
End Function

' ---- API ----
' MIDAS_USTYAPI ends the PJCF comment with DYNAMIC on the dynamic model (owner 2026-11-06: live load from the
' staged model only). An unreadable PJCF does not stop the macro.
Private Function IsDynamicModel() As Boolean
    Dim pj As Object, k As Variant
    On Error GoTo Done
    Set pj = TableGet("PJCF")
    If pj Is Nothing Then Exit Function
    For Each k In pj.keys
        If Right$(UCase$(Trim$(CStr(GetField(pj(k), "COMMENT")))), 7) = "DYNAMIC" Then IsDynamicModel = True
    Next k
Done:
End Function

Private Function TableGet(ByVal t As String) As Object
    Dim resp As String, root As Object, k As Variant
    If Not Api("GET", "db/" & t, "", resp) Then Exit Function
    Set root = ParseJson(resp)
    If root Is Nothing Then Exit Function
    For Each k In root.keys
        If UCase$(CStr(k)) = UCase$(t) Then Set TableGet = root(k): Exit Function
    Next k
End Function

Private Function Api(ByVal method As String, ByVal ep As String, ByVal body As String, ByRef resp As String) As Boolean
    Api = McApi(method, ep, body, resp)
    If McStatus() = 404 Then McForgetKey
End Function

Private Function GetKey(wb As Workbook) As Boolean
    GetKey = McGetKey(wb, True, MSG_TITLE)
End Function

' ---- helpers ----
Private Function GetField(o As Variant, ByVal f As String) As Variant
    If IsObject(o) Then
        If o.Exists(f) Then
            If IsObject(o(f)) Then Set GetField = o(f) Else GetField = o(f)
        End If
    End If
End Function

Private Function NodeVal(nodes As Object, ByVal n As Long, ByVal f As String) As Double
    If Not nodes.Exists(CStr(n)) Then Err.Raise vbObjectError + 5, , "Node " & n & " is not in the model."
    NodeVal = CDbl(nodes(CStr(n))(f))
End Function

Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    Txt = Trim$(CStr(v))
End Function

Private Function MaxOf(a() As Double) As Double
    Dim i As Long
    For i = 1 To nSt
        If a(i) > MaxOf Then MaxOf = a(i)
    Next i
End Function

Private Sub WriteText(ByVal path As String, ByVal txtOut As String)
    Dim f As Integer
    If path = "" Then Exit Sub
    f = FreeFile
    Open path For Output As #f
    Print #f, txtOut
    Close #f
End Sub
