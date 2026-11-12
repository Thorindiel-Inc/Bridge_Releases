Attribute VB_Name = "MIDAS_CORE"
Option Explicit
Option Private Module
' ============================================================================
'  MIDAS_CORE - shared helpers of the MIDAS macros (MIDAS_KOPRU, MIDAS_USTYAPI,
'  MIDAS_KENARAYAK, MIDAS_RTI, MIDAS_MOD, MAK_LL).
' ----------------------------------------------------------------------------
'  One copy of what every macro needs:
'    MAPI key      McGetKey, McSetKey, McForgetKey
'    HTTP          McApi (optional retry), McStatus, McCloseHttp
'    JSON          ParseJson, QS, JN, JNum, TopKeys, CountOf, TableHasRecords
'    cells         NmRange, Num, NumOr0
'    files         MakeDir, WriteText, ReadUtf8, WriteUtf8, WaitForFile, Pause
'
'  A macro that uses this module declares
'    Private Const SCRIPT_NEEDS As String = "midas-core"
'  so UPDATE adds this module to a workbook that does not have it yet.
'
'  MAPI key order (owner 2026-11-12): workbook name G_MAPI_KEY, then INPUT!T3,
'  then a key given earlier in this Excel session (scripted entry or prompt),
'  then the Windows variable MIDAS_MAPI_KEY, then the prompt. The key lives in
'  memory only; it is never written to a file or a log.
'
'  Option Private Module: the procedures are visible to the other modules of
'  this workbook only, not as worksheet functions and not in the Alt+F8 list.
'  Needs PANEL (UiInputS, UiLog).
' ============================================================================

Private Const SCRIPT_ID As String = "midas-core"
Private Const SCRIPT_VERSION As String = "2026-11-12a"
Private Const SCRIPT_CHANGELOG As String = "New: one shared copy of the MAPI key, HTTP, JSON, number and file helpers of the MIDAS macros"

Private Const API_BASE As String = "https://moa-engineers.midasit.com:443/civil/"
Private Const KEY_NAME As String = "G_MAPI_KEY"
Private Const KEY_SHEET As String = "INPUT"
Private Const KEY_CELL As String = "T3"
Private Const KEY_ENV As String = "MIDAS_MAPI_KEY"

Private mKey As String
Private mHttp As Object
Private mStatus As Long

Public Function McVersion() As String
    McVersion = SCRIPT_VERSION
End Function

' ---------------------------------------------------------------------------
'  MAPI key
' ---------------------------------------------------------------------------
' True when a key is known. wb = the workbook that holds the key cell;
' ask = False for scripted runs (no prompt).
Public Function McGetKey(wb As Workbook, Optional ByVal ask As Boolean = True, Optional ByVal title As String = "MIDAS") As Boolean
    Dim cellKey As String
    On Error Resume Next
    cellKey = Trim$(CStr(wb.Names(KEY_NAME).RefersToRange.Cells(1, 1).Value))
    If cellKey = "" Then cellKey = Trim$(CStr(wb.Worksheets(KEY_SHEET).Range(KEY_CELL).Value))
    On Error GoTo 0
    If cellKey <> "" Then mKey = cellKey
    If mKey = "" Then mKey = Trim$(Environ$(KEY_ENV))
    If mKey = "" And ask Then mKey = Trim$(UiInputS("MIDAS MAPI key (kept in memory for this Excel session only):", title))
    McGetKey = (mKey <> "")
End Function

' Key of a scripted entry; an empty text forgets the key in memory.
Public Sub McSetKey(ByVal key As String)
    mKey = Trim$(key)
End Sub

Public Sub McForgetKey()
    mKey = ""
End Sub

' ---------------------------------------------------------------------------
'  HTTP
' ---------------------------------------------------------------------------
' One MAPI call. True = HTTP 2xx and no "error" in the answer.
' maxTries > 1: an answer with HTTP 5xx or a transport error is sent again
' (the MAPI goes through the MIDAS relay server: a large answer can arrive
' broken, 'invalid or unrecognized response', owner 2026-11-09).
' quiet = no log line for a retry.
Public Function McApi(ByVal method As String, ByVal ep As String, ByVal body As String, ByRef resp As String, _
                      Optional ByVal maxTries As Long = 1, Optional ByVal recvMs As Long = 600000, _
                      Optional ByVal quiet As Boolean = False) As Boolean
    Dim tries As Long
Again:
    mStatus = 0
    On Error GoTo Bad
    If mHttp Is Nothing Then Set mHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    mHttp.Open method, API_BASE & ep, False
    mHttp.SetTimeouts 30000, 60000, recvMs, recvMs
    mHttp.SetRequestHeader "MAPI-Key", mKey
    mHttp.SetRequestHeader "Content-Type", "application/json"
    If method = "GET" Then mHttp.Send Else mHttp.Send body
    resp = mHttp.ResponseText
    mStatus = mHttp.Status
    If mStatus >= 500 And tries < maxTries - 1 Then
        tries = tries + 1
        RetryNote ep, "HTTP " & mStatus, tries, maxTries, quiet
        GoTo Again
    End If
    McApi = (mStatus >= 200 And mStatus < 300 And InStr(1, resp, """error""", vbTextCompare) = 0)
    Exit Function
Bad:
    resp = "HTTP error: " & Err.Description
    Set mHttp = Nothing
    mStatus = 0
    If tries < maxTries - 1 Then
        tries = tries + 1
        RetryNote ep, Err.Description, tries, maxTries, quiet
        Resume Again
    End If
    McApi = False
End Function

' HTTP status of the last McApi call; 0 = no answer.
Public Function McStatus() As Long
    McStatus = mStatus
End Function

Public Sub McCloseHttp()
    Set mHttp = Nothing
End Sub

Private Sub RetryNote(ByVal ep As String, ByVal why As String, ByVal tries As Long, ByVal maxTries As Long, ByVal quiet As Boolean)
    If Not quiet Then UiLog "! " & ep & ": no valid answer (" & why & "), sent again (" & (tries + 1) & " / " & maxTries & ")"
    Application.Wait Now + TimeSerial(0, 0, 3 * tries)
End Sub

' ---------------------------------------------------------------------------
'  JSON: writing
' ---------------------------------------------------------------------------
' Number as JSON text: point decimal on any Excel language, no bare ".5".
Public Function JN(ByVal x As Double) As String
    Dim s As String
    If Abs(x) < 0.000000000001 Then JN = "0": Exit Function
    s = Trim$(Str$(x))
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    JN = s
End Function

' The same, rounded to 6 decimals.
Public Function JNum(ByVal v As Double) As String
    Dim s As String
    s = Trim$(Str$(Round(v, 6)))
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    JNum = s
End Function

' Text as a JSON string: quotes and backslashes escaped, everything outside
' printable ASCII as \uXXXX.
Public Function QS(ByVal s As String) As String
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

' ---------------------------------------------------------------------------
'  JSON: reading
' ---------------------------------------------------------------------------
' Keys of the second level of an answer ({"TABLE":{"1":..,"2":..}} -> "1","2").
Public Function TopKeys(ByVal resp As String) As String
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

Public Function CountOf(ByVal s As String, ByVal what As String) As Long
    CountOf = (Len(s) - Len(Replace(s, what, ""))) / Len(what)
End Function

Public Function TableHasRecords(ByVal resp As String, ByVal tbl As String) As Boolean
    TableHasRecords = (InStr(1, resp, """" & tbl & """:{""", vbBinaryCompare) > 0)
End Function

' JSON text -> Scripting.Dictionary (objects), 1-based arrays, values.
' Nothing when the text does not start with "{".
Public Function ParseJson(ByVal s As String) As Object
    Dim p As Long, v As Variant
    p = 1
    SkipWs s, p
    If Mid$(s, p, 1) <> "{" Then Exit Function
    ParseValue s, p, v
    Set ParseJson = v
End Function

Private Sub SkipWs(ByRef s As String, ByRef p As Long)
    Do While p <= Len(s)
        Select Case Mid$(s, p, 1)
            Case " ", vbTab, vbCr, vbLf: p = p + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

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

' ---------------------------------------------------------------------------
'  Cells
' ---------------------------------------------------------------------------
Public Function NmRange(wb As Workbook, ByVal nm As String) As Range
    On Error Resume Next
    Set NmRange = wb.Names(nm).RefersToRange
    On Error GoTo 0
    If NmRange Is Nothing Then Err.Raise vbObjectError + 20, , "Name " & nm & " missing in " & wb.Name & "."
End Function

Public Function Num(ByVal v As Variant, ByVal what As String) As Double
    If IsError(v) Then Err.Raise vbObjectError + 21, , what & ": error value."
    If IsEmpty(v) Or Not IsNumeric(v) Then Err.Raise vbObjectError + 21, , what & ": not a number ('" & CStr(v) & "')."
    Num = CDbl(v)
End Function

Public Function NumOr0(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsNumeric(v) And Not IsEmpty(v) Then NumOr0 = CDbl(v)
End Function

' ---------------------------------------------------------------------------
'  Files
' ---------------------------------------------------------------------------
Public Sub MakeDir(ByVal p As String)
    If Len(Dir$(p, vbDirectory)) = 0 Then MkDir p
End Sub

' ANSI text, no line end added.
Public Sub WriteText(ByVal p As String, ByVal txt As String)
    Dim f As Integer: f = FreeFile
    Open p For Output As #f
    Print #f, txt;
    Close #f
End Sub

Public Function ReadUtf8(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText
    st.Close
End Function

Public Sub WriteUtf8(ByVal path As String, ByVal txt As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText txt
    st.SaveToFile path, 2
    st.Close
End Sub

' True when file f exists and its size has stopped changing, within secs.
Public Function WaitForFile(ByVal f As String, ByVal secs As Long) As Boolean
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

Public Sub Pause(ByVal s As Single)
    Dim t0 As Single
    t0 = Timer
    Do While Timer - t0 < s And Timer >= t0
        DoEvents
    Loop
End Sub
