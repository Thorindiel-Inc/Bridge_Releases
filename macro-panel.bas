Attribute VB_Name = "PANEL"
Option Explicit
Private Type MONINFO
    cb As Long
    mL As Long: mT As Long: mR As Long: mB As Long
    wL As Long: wT As Long: wR As Long: wB As Long
    flags As Long
End Type
#If Win64 Then
Private Declare PtrSafe Function GetWindowLongPtrW Lib "user32" (ByVal h As LongPtr, ByVal i As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtrW Lib "user32" (ByVal h As LongPtr, ByVal i As Long, ByVal v As LongPtr) As LongPtr
#Else
Private Declare PtrSafe Function GetWindowLongPtrW Lib "user32" Alias "GetWindowLongW" (ByVal h As LongPtr, ByVal i As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtrW Lib "user32" Alias "SetWindowLongW" (ByVal h As LongPtr, ByVal i As Long, ByVal v As LongPtr) As LongPtr
#End If
Private Declare PtrSafe Function FindWindowW Lib "user32" (ByVal c As LongPtr, ByVal t As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal h As LongPtr, ByVal a As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal f As Long) As Long
Private Declare PtrSafe Function MonitorFromWindow Lib "user32" (ByVal h As LongPtr, ByVal f As Long) As LongPtr
Private Declare PtrSafe Function GetMonitorInfoW Lib "user32" (ByVal m As LongPtr, ByRef i As MONINFO) As Long
' ============================================================================
'  PANEL - the macro window of the workbook. The button MAKROLAR on INPUT opens it (PANEL_AC).
'  Left: every tool, by group. Right: what it does, its input fields, Run, progress, results log.
'  Add a tool = add one line to Tools():
'     Group|Title|Macro|Sheet|Description|key~label~type~default;key~label~type~default
'     Sheet: empty, or the sheet the macro runs on; ? = span letter (4?_KILIF)
'     type: n number, c list (default = a/b/c, first entry is the default), s filled spans
'  Fields are drawn when the tool is picked, the last values are remembered in the workbook.
'  The window opens at a share of the monitor Excel is on and can be resized: Layout scales fonts and
'  controls with Zoom and gives the extra room to the results log. Everything is laid out in BASE_W x BASE_H.
'  The tool list is rows of labels: an MSForms ListBox drew at the wrong size on scaled screens.
'  Macros read them with UiNum / UiIdx. They call UiMsg, UiInput, UiInputS, UiChoice, UiProgress,
'  UiLog instead of MsgBox / InputBox: long texts scroll, nothing is cut off.
'  frmPanel and frmAsk are built here from code, so this module ships through the UPDATE feed.
'  Building needs "Trust access to the VBA project object model", as UPDATE does.
'  Turkish letters are written as \uXXXX (this file stays ASCII), see U().
' ============================================================================

Private Const SCRIPT_ID As String = "macro-panel"
Private Const SCRIPT_VERSION As String = "2026-11-09g"
Private Const SCRIPT_CHANGELOG As String = "RTI sonuclari: no LM71 scale field (the rail stress pictures are pasted by hand); Modal analiz sonuclari: read and mode pictures in one tool (Mod sekilleri row removed); Durdur: the Run button stops the running macro at its next step; RTI row text fits; RTI sonuclari: Islem and resolution fields; live log: time, mark and duration of every step sent or read, questions and answers, fields of the run, failures marked; Makro guncellemesi: the window hides and UPDATE starts after it (UPDATE rewriting PANEL from inside the window crashed Excel); New tool: Mod sekilleri (MIDAS_SEKIL, model and mode-shape pictures into Z_YATAY); New tool: RTI sonuclari (MIDAS_RTI); tool list shows only the macros this workbook has (one module for the main and the chapter 8 workbook); chapter 8 tools: Degerleri cagir, Modal analiz sonuclari; Ustyapi modeli: gate for the staged LM71 single-track and a=1 cases"
Private Const FORM_VERSION As String = "2026-11-06d"
Private Const F_PANEL As String = "frmPanel"
Private Const F_ASK As String = "frmAsk"
Private Const C_DARK As Long = 5059095      ' RGB(23, 50, 77)
Private Const C_ACC As Long = 10248727      ' RGB(23, 98, 156)
Private Const C_BACK As Long = 16512763
Private Const C_LINE As Long = 14472137
Private Const C_WHITE As Long = 16777215
Private Const C_RED As Long = 2039992
Private Const C_AMBER As Long = 23464
Private Const BASE_W As Double = 900
Private Const BASE_H As Double = 600
Private Const LIST_W As Double = 232
Private Const RX As Double = 262
Private Const ROWS_MAX As Long = 24
Private Const ROW_H As Double = 22
Private Const FONT_BASE As Double = 12
Private Const FIELD_TOP As Double = 124
Private Const C_GREEN As Long = 5081622      ' RGB(22, 138, 77)
Private Const C_TEXT As Long = 3353901      ' RGB(45, 45, 51)
Private Const C_MUTED As Long = 9866370      ' RGB(130, 140, 150)
Private Const C_STRIPE As Long = 16513269      ' RGB(245, 248, 251)
Private Const C_AMBERBG As Long = 14808574      ' RGB(254, 245, 225)
Private Const C_REDBG As Long = 15395325      ' RGB(253, 233, 234)
Private Const C_BLUEBG As Long = 16445668      ' RGB(228, 240, 250)
Private Const ASK_W As Double = 660
Private Const ASK_HEAD As Double = 62
Private Const ASK_KEY As Double = 138
Private Const ASK_KEYLEN As Long = 22
Private Const ASK_BTN As Double = 104
Private Const L_PLAIN As Long = 0
Private Const L_OK As Long = 1
Private Const L_WARN As Long = 2
Private Const L_BAD As Long = 3
Private Const L_HEAD As Long = 4
Private Const L_PICK As Long = 5
Private Const FIELD_ROW As Double = 29
Private Const LOG_TOP As Double = 276

Private pForm As Object
Private toolList As Variant, rowTool() As Long, curTool As Long, curRow As Long
Private inLayout As Boolean, closed As Boolean, barPct As Double, zoomNow As Double, logW As Double
Private fKey() As String, fTyp() As String, fCnt As Long
Private vals As Object
Private uiBtn As Long, uiText As String, uiPick As Long, pickDef As Long, running As Boolean
Private stepTxt As String, stepT As Double, runBad As Boolean   ' live log: open step, its start, a failure was logged
Private stopReq As Boolean, stopRaised As Boolean, runBack As Long   ' Durdur: asked, raised once; colour of the Run button

Private Function Tools() As Variant
    Tools = Array( _
        "Girdiler|DXF'ten g\u00FCzergah oku|DXF_OKU_OLUSTUR||Plan DXF dosyas\u0131ndaki eksen ve aks \u00E7izgilerini okur, aks koordinatlar\u0131n\u0131 ve a\u00E7\u0131kl\u0131k yar\u0131\u00E7aplar\u0131n\u0131 INPUT sayfas\u0131na yazar.|", _
        "Girdiler|De\u011Ferleri \u00E7a\u011F\u0131r|DEGER_CAGIRMA.DEGER_CAGIRMA||Ana \u00E7al\u0131\u015Fma kitab\u0131ndan G_ adl\u0131 girdileri (MAPI anahtar\u0131 dahil) ve \u00F6nceki b\u00F6l\u00FCm\u00FCn son tablo, \u015Fekil ve sayfa numaralar\u0131n\u0131 bu dosyaya \u00E7eker.|", _
        "Girdiler|A\u00E7\u0131kl\u0131k b\u00F6l\u00FCmlerini g\u00FCncelle|BOLUMLERI_GUNCELLE||Dolu a\u00E7\u0131kl\u0131k say\u0131s\u0131na g\u00F6re 3. ve 4. b\u00F6l\u00FCm sayfalar\u0131n\u0131 ekler veya siler; sayfa s\u0131ras\u0131n\u0131, numaralar\u0131, ba\u015Fl\u0131klar\u0131 ve listeleri g\u00FCnceller.|", _
        "MIDAS modelleri|\u00DCstyap\u0131 modeli|MIDAS_USTYAPI_OLUSTUR||Civil NX'te a\u00E7\u0131k modele tek a\u00E7\u0131kl\u0131k \u0131zgara modelini kurar. A\u00E7\u0131k model silinir.|model~Model~c~A\u015Famal\u0131 (tasar\u0131m)/Dinamik (modlar);span~A\u00E7\u0131kl\u0131k~s~;strip~\u015Eerit aral\u0131\u011F\u0131 (m)~n~1.5;link~Link boyu (m)~n~0.35;kfix~Sabit yay rijitli\u011Fi (kN/m)~n~100000;mvek~Tek hat ve \u03B1=1 LM71 (a\u015Famal\u0131)~c~Hay\u0131r/Evet", _
        "MIDAS modelleri|K\u00F6pr\u00FC genel modeli|MIDAS_KOPRU_OLUSTUR||Civil NX'te a\u00E7\u0131k modele t\u00FCm k\u00F6pr\u00FC modelini kurar. A\u00E7\u0131k model silinir.|model~Model~c~Statik/Sismik/RTI;mesh~Eleman boyu (m)~n~1.5;link~Mesnet link boyu (m)~n~0.35;kfix~Sabit yatay rijitlik kh (kN/m)~n~100000;ll~Hareketli y\u00FCk (RTI)~c~Do\u011Frusal (MV)/Do\u011Frusal olmayan", _
        "MIDAS sonu\u00E7lar\u0131|Modal analiz sonu\u00E7lar\u0131|MIDAS_MOD_OKU||Sismik modelden periyot, frekans ve k\u00FCtle kat\u0131l\u0131mlar\u0131n\u0131 MOD'a okur; en b\u00FCy\u00FCk X ve Y modlar\u0131n\u0131n \u015Fekillerini Z_YATAY'a resimler. Modele bir \u015Fey yazmaz.|islem~\u0130\u015Flem~c~Oku + resimle/Yaln\u0131z resimle/Yaln\u0131z oku;res~\u00C7\u00F6z\u00FCn\u00FCrl\u00FCk (px)~c~4000 px/3000 px/2000 px/Ekran boyutu", _
        "MIDAS sonu\u00E7lar\u0131|RTI sonu\u00E7lar\u0131|MIDAS_RTI_OKU||Civil NX'teki RTI modelinden ray gerilmesini ve yer de\u011Fi\u015Ftirmeleri RTI_SONUC'a, resimleri 3B_RTI_Y'ye yazar. Sonu\u00E7lar\u0131 de\u011Fi\u015Ftirmez.|islem~\u0130\u015Flem~c~Oku + resimle/Yaln\u0131z resimle/Yaln\u0131z oku;res~\u00C7\u00F6z\u00FCn\u00FCrl\u00FCk (px)~c~4000 px/3000 px/2000 px/Ekran boyutu", _
        "Kiri\u015F tasar\u0131m\u0131|MIDAS'tan hareketli y\u00FCk|MAK_LL.MAK_LL|4?_KILIF|Civil NX'te a\u00E7\u0131k a\u015Famal\u0131 modelden hareketli y\u00FCk kuvvetlerini okur, a\u00E7\u0131kl\u0131\u011F\u0131n LL tablosuna yazar (Tablo 27).|", _
        "Rapor|\u0130\u00E7indekiler|ICINDEKILER_AYARLA||\u0130\u00E7indekiler listesini sayfa numaralar\u0131 ve ba\u011Flant\u0131lar\u0131yla yeniden olu\u015Fturur.|", _
        "Rapor|\u015Eekil ve tablo listeleri|FIG_TAB_AYARLA||\u015Eekil ve tablo listelerini sayfa numaralar\u0131 ve ba\u011Flant\u0131lar\u0131yla yeniden olu\u015Fturur.|", _
        "Yard\u0131m|K\u0131lavuzlar|KILAVUZ_AC||Makro k\u0131lavuzlar\u0131n\u0131 taray\u0131c\u0131da a\u00E7ar.|", _
        "Yard\u0131m|Makro g\u00FCncellemesi|CheckBridgeMacroUpdates||Makrolar\u0131 yay\u0131nla kar\u015F\u0131la\u015Ft\u0131r\u0131r, yeni s\u00FCr\u00FCmleri kurar.|")
End Function

' tools whose macro is in this workbook; without access to the VBA project every tool is kept
Private Function Present(all As Variant) As Variant
    Dim i As Long, n As Long, out() As String, p As Variant
    ReDim out(0 To UBound(all))
    For i = LBound(all) To UBound(all)
        p = Split(all(i), "|")
        If HasMacro(CStr(p(2))) Then out(n) = all(i): n = n + 1
    Next i
    If n = 0 Then Present = all: Exit Function
    ReDim Preserve out(0 To n - 1)
    Present = out
End Function

Private Function HasMacro(ByVal m As String) As Boolean
    Dim c As Object, cs As Object, modN As String, proc As String, p As Long, ln As Long
    On Error Resume Next
    Set cs = ThisWorkbook.VBProject.VBComponents
    If cs Is Nothing Then HasMacro = True: Exit Function
    p = InStr(m, ".")
    If p > 0 Then modN = Left$(m, p - 1): proc = mid$(m, p + 1) Else proc = m
    For Each c In cs
        If c.Type = 1 And (modN = "" Or StrComp(c.Name, modN, vbTextCompare) = 0) Then
            ln = 0
            Err.Clear
            ln = c.CodeModule.ProcStartLine(proc, 0)
            If Err.Number = 0 And ln > 0 Then HasMacro = True: Exit Function
        End If
    Next c
End Function

Private Function CHK() As String
    CHK = ChrW(&H2714)
End Function

Private Function CRS() As String
    CRS = ChrW(&H2716)
End Function

' \uXXXX -> letter
Private Function U(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, "\u")
    Do While p > 0
        s = Left$(s, p - 1) & ChrW(CLng("&H" & mid$(s, p + 2, 4))) & mid$(s, p + 6)
        p = InStr(p + 1, s, "\u")
    Loop
    U = s
End Function

Public Sub PANEL_AC()
    On Error GoTo Fail
    If closed Then
        Set pForm = Nothing
        closed = False
    End If
    If Not pForm Is Nothing Then
        pForm.Show vbModeless
        Exit Sub
    End If
    EnsureForms
    Set pForm = VBA.UserForms.Add(F_PANEL)
    pForm.Caption = U("Makrolar")
    pForm.Controls("lblTitle").Caption = U("K\u00F6pr\u00FC makrolar\u0131")
    pForm.Controls("lblSub").Caption = U("Bir ara\u00E7 se\u00E7in, alanlar\u0131 doldurun, \u00C7al\u0131\u015Ft\u0131r'a bas\u0131n")
    pForm.Controls("lblVer").Caption = SCRIPT_VERSION
    pForm.Controls("cmdRun").Caption = U("\u00C7al\u0131\u015Ft\u0131r")
    pForm.Controls("cmdCopy").Caption = U("Kopyala")
    pForm.Controls("cmdSave").Caption = U("Kaydet")
    pForm.Controls("cmdClear").Caption = U("Temizle")
    pForm.Controls("cmdClose").Caption = U("Kapat")
    pForm.Controls("lblLog").Caption = U("Sonu\u00E7lar")
    FillTools
    Layout
    pForm.Show vbModeless
    FitWindow
    Exit Sub
Fail:
    MsgBox "The macro window could not open: " & Err.Description & vbCrLf & vbCrLf & _
           "File > Options > Trust Center > Macro Settings > Trust access to the VBA project object model.", vbExclamation, "PANEL"
End Sub

Public Sub PANEL_KUR()
    EnsureForms
End Sub

' test: opens the window on a tool, returns nothing
Public Sub PANEL_TEST(ByVal row As Long, ByVal txt As String)
    PANEL_AC
    EvRow row
    UiLog txt
    UiProgress 45, "test"
End Sub

' test: a fake run, Durdur pressed at step 3; the fake macro has its own handler whose cleanup
' reports progress again (the stop must raise only once); returns the states and the log
Public Function PANEL_STOPTEST() As String
    Dim s As String
    PANEL_AC
    pForm.Controls("txtLog").Text = ""
    running = True: stepTxt = "": runBad = False: stopReq = False: stopRaised = False
    SetBusy True
    s = "busy caption=" & pForm.Controls("cmdRun").Caption & ", enabled=" & pForm.Controls("cmdRun").Enabled
    s = s & vbLf & "fake macro: " & FakeMacro()
    s = s & vbLf & "after: caption=" & pForm.Controls("cmdRun").Caption & ", enabled=" & pForm.Controls("cmdRun").Enabled & ", UiStopped=" & UiStopped()
    FlushStep ""
    LogLine IIf(runBad Or stopReq, ChrW(215), ChrW(9632)), IIf(stopReq, "Durduruldu", "Bitti"), -1
    running = False: SetBusy False
    s = s & vbLf & "idle caption=" & pForm.Controls("cmdRun").Caption & ", enabled=" & pForm.Controls("cmdRun").Enabled
    PANEL_STOPTEST = s & vbLf & "LOG:" & vbLf & pForm.Controls("txtLog").Text
End Function

Private Function FakeMacro() As String
    Dim i As Long, done As Long
    On Error GoTo Fail
    For i = 1 To 10
        UiProgress i * 10, "table " & i & "/10  db/STAG"
        done = i
        If i = 3 Then EvRun
    Next i
    FakeMacro = "ran to the end (" & done & " steps): the stop did NOT act"
    Exit Function
Fail:
    FakeMacro = "stopped after step " & done & ": " & Err.Description
    On Error Resume Next
    UiProgress 100, "cleanup: units back"
    FakeMacro = FakeMacro & "; cleanup progress err=" & Err.Number
    UiMsg "STOPPED: " & Err.Description, vbExclamation, "Test"
End Function

' test: fakes a run from the window and returns the time of 500 steps in ms, then the log
Public Function PANEL_LOGTEST() As String
    Dim i As Long, t As Double
    PANEL_AC
    pForm.Controls("txtLog").Text = ""
    running = True: stepTxt = "": runBad = False
    LogLine ChrW(9658), "Test" & FieldText(), -1
    UiProgress 1, "table 1/3  db/NODE  12 kB"
    UiProgress 2, "table 1/3  db/NODE  12 kB"
    UiProgress 3, "read-back nodes"
    UiLog "Read-back:" & vbLf & "  nodes 10 / 10" & vbLf & "  elements 9 / 10  FAIL" & vbLf & "  ! cap fit: bearing off the web" & vbLf & "  key from INPUT!T3"
    UiProgress 4, "save"
    UiLog "Done: 3 tables"
    t = Timer
    For i = 1 To 500: UiProgress 5, "table " & i & "/500  db/STAG": Next i
    FlushStep IIf(runBad, ChrW(215), "")
    LogLine IIf(runBad, ChrW(215), ChrW(9632)), "Bitti", -1
    running = False
    PANEL_LOGTEST = Format$((Timer - t) * 1000#, "0") & " ms" & vbLf & pForm.Controls("txtLog").Text
End Function

' test: picks a row and returns the names and values of its fields; run = True presses Run and returns the log
Public Function PANEL_PROBE(ByVal row As Long, ByVal run As Boolean) As String
    Dim i As Long, s As String
    PANEL_AC
    EvRow row
    For i = 0 To fCnt - 1
        s = s & fKey(i) & "=" & pForm.Controls("p_" & fKey(i)).Value & ";"
    Next i
    If run Then
        EvRun
        s = s & " LOG[" & pForm.Controls("txtLog").Text & "] STEP[" & pForm.Controls("lblStep").Caption & "]"
    End If
    PANEL_PROBE = s
End Function

' test: titles of the tools this workbook shows
Public Function PANEL_TOOLS() As String
    Dim v As Variant, i As Long, s As String
    v = Present(Tools())
    For i = LBound(v) To UBound(v): s = s & U(CStr(Split(v(i), "|")(1))) & "; ": Next i
    PANEL_TOOLS = s
End Function

' test: a yes / no dialog, the answer goes to the log
Public Sub PANEL_ASKTEST()
    UiLog "answer " & UiMsg("Test sorusu" & vbLf & String(200, "x"), vbYesNo + vbQuestion, "Deneme")
    UiLog "pick " & UiChoice("Hangi?", Array("Bir", "Iki"), "Deneme")
End Sub

' test: state of the window
Public Function PANEL_STATE() As String
    If PanelOpen() Then
        PANEL_STATE = "open inside " & pForm.InsideWidth & " x " & pForm.InsideHeight & " zoom " & pForm.Zoom & " tool " & pForm.Controls("lblName").Caption & " log[" & pForm.Controls("txtLog").Text & "]"
    Else
        PANEL_STATE = "closed"
    End If
End Function

' MSForms windows have a fixed frame: this gives the window a sizing frame and a maximise button, then
' sizes it in pixels to a share of the monitor Excel is on (Application.Width is in other units than the form).
' The size change fires UserForm_Resize, so Layout follows.
Private Sub FitWindow()
    Dim h As LongPtr, s As LongPtr, cap As String, tmp As String, mi As MONINFO, w As Long, ht As Long
    On Error Resume Next
    cap = pForm.Caption
    tmp = "PNL" & Format(Timer * 1000, "0")
    pForm.Caption = tmp
    h = FindWindowW(StrPtr("ThunderDFrame"), StrPtr(tmp))
    pForm.Caption = cap
    If h = 0 Then Exit Sub
    s = GetWindowLongPtrW(h, -16)
    SetWindowLongPtrW h, -16, s Or &H40000 Or &H10000
    SetWindowPos h, 0, 0, 0, 0, 0, &H27
    mi.cb = LenB(mi)
    If GetMonitorInfoW(MonitorFromWindow(Application.Hwnd, 2), mi) = 0 Then Exit Sub
    w = (mi.wR - mi.wL) * 0.66
    ht = (mi.wB - mi.wT) * 0.76
    SetWindowPos h, 0, mi.wL + ((mi.wR - mi.wL) - w) \ 2, mi.wT + ((mi.wB - mi.wT) - ht) \ 2, w, ht, &H14
End Sub

Public Sub EvResize()
    If pForm Is Nothing Or inLayout Then Exit Sub
    Layout
End Sub

' zoom from the window size; the room left over on one axis goes to the right side and the log
Private Sub Layout()
    Dim z As Double, w As Double, h As Double, rw As Double, c As Variant, realW As Double, realH As Double
    inLayout = True
    On Error GoTo Done
    ' InsideWidth / InsideHeight are in zoomed units, so real size = inside x current zoom
    realW = pForm.InsideWidth * pForm.Zoom / 100
    realH = pForm.InsideHeight * pForm.Zoom / 100
    z = realW / BASE_W
    If realH / BASE_H < z Then z = realH / BASE_H
    If z < 0.7 Then z = 0.7
    If z > 2.5 Then z = 2.5
    pForm.Zoom = CLng(z * 100)
    zoomNow = CLng(z * 100) / 100
    w = realW / zoomNow
    h = realH / zoomNow
    logW = w
    rw = w - RX - 14
    pForm.Controls("lblHeadBack").Width = w
    pForm.Controls("lblVer").Left = w - 134
    pForm.Controls("lblListBack").Height = h - 64 - 14
    pForm.Controls("lblName").Width = rw - 122
    pForm.Controls("lblDesc").Width = rw - 122
    pForm.Controls("cmdRun").Left = w - 14 - 110
    pForm.Controls("lblBarBack").Width = rw
    pForm.Controls("lblBar").Width = rw * barPct / 100 + 0.1
    pForm.Controls("lblStep").Width = rw
    pForm.Controls("txtLog").Width = rw
    pForm.Controls("txtLog").Height = h - LOG_TOP - 48
    For Each c In Array("cmdCopy", "cmdSave", "cmdClear", "cmdClose")
        pForm.Controls(CStr(c)).top = h - 40
    Next c
    pForm.Controls("cmdClose").Left = w - 14 - 90
    PlaceFields
Done:
    inLayout = False
End Sub

Private Sub PlaceFields()
    Dim i As Long, half As Double, l As Double, t As Double
    half = (logW - RX - 14) / 2
    For i = 0 To fCnt - 1
        l = RX + (i Mod 2) * half
        t = FIELD_TOP + (i \ 2) * FIELD_ROW
        With pForm.Controls("q_" & fKey(i))
            .Left = l: .top = t + 3: .Width = half - 136: .Height = 20
        End With
        With pForm.Controls("p_" & fKey(i))
            .Left = l + half - 132: .top = t: .Width = 122: .Height = 25
        End With
    Next i
End Sub

Private Sub FillTools()
    Dim i As Long, p As Variant, g As String, n As Long
    toolList = Present(Tools())
    ReDim rowTool(0 To ROWS_MAX - 1)
    For i = LBound(toolList) To UBound(toolList)
        p = Split(toolList(i), "|")
        If p(0) <> g Then
            g = p(0)
            RowText n, UpperTr(g), -1
            n = n + 1
        End If
        RowText n, "   " & U(CStr(p(1))), i
        n = n + 1
    Next i
    For i = n To ROWS_MAX - 1
        rowTool(i) = -1
        pForm.Controls("r" & i).Visible = False
    Next i
    curRow = -1
    EvRow 1
End Sub

Private Sub RowText(ByVal n As Long, ByVal txt As String, ByVal tool As Long)
    If n >= ROWS_MAX Then Err.Raise 513, , "More list rows than ROWS_MAX"
    rowTool(n) = tool
    With pForm.Controls("r" & n)
        .Caption = txt
        .ForeColor = IIf(tool < 0, C_ACC, C_DARK)
    End With
End Sub

Private Function UpperTr(ByVal s As String) As String
    UpperTr = UCase$(Replace(U(s), ChrW(305), "I"))
End Function

Public Sub EvRow(ByVal k As Long)
    Dim p As Variant
    If pForm Is Nothing Or running Then Exit Sub
    If k < 0 Or k >= ROWS_MAX Then Exit Sub
    If rowTool(k) < 0 Then Exit Sub
    If curRow >= 0 Then
        With pForm.Controls("r" & curRow)
            .BackColor = C_WHITE
            .ForeColor = C_DARK
        End With
    End If
    curRow = k
    With pForm.Controls("r" & k)
        .BackColor = C_ACC
        .ForeColor = C_WHITE
    End With
    curTool = rowTool(k)
    p = Split(toolList(curTool), "|")
    pForm.Controls("lblName").Caption = U(CStr(p(1)))
    pForm.Controls("lblDesc").Caption = U(CStr(p(4)))
    DrawFields CStr(p(5))
End Sub

' input fields of the tool, two per row
Private Sub DrawFields(ByVal spec As String)
    Dim i As Long, a As Variant, q As Variant, c As Object, o As Variant, keep As String
    For i = pForm.Controls.Count - 1 To 0 Step -1
        If Left$(pForm.Controls(i).Name, 2) = "p_" Or Left$(pForm.Controls(i).Name, 2) = "q_" Then pForm.Controls.Remove pForm.Controls(i).Name
    Next i
    fCnt = 0
    If spec <> "" Then
        a = Split(spec, ";")
        ReDim fKey(0 To UBound(a)): ReDim fTyp(0 To UBound(a))
        For i = 0 To UBound(a)
            q = Split(a(i), "~")
            fKey(i) = q(0): fTyp(i) = q(2): fCnt = fCnt + 1
            Set c = pForm.Controls.Add("Forms.Label.1", "q_" & q(0), True)
            c.BackStyle = 0: c.Font.Name = "Segoe UI": c.Font.Size = FONT_BASE: c.Caption = U(CStr(q(1)))
            If q(2) = "n" Then
                Set c = pForm.Controls.Add("Forms.TextBox.1", "p_" & q(0), True)
                c.Text = Remembered(CStr(q(0)), CStr(q(3)))
                c.TextAlign = 3
            Else
                Set c = pForm.Controls.Add("Forms.ComboBox.1", "p_" & q(0), True)
                c.style = 2
                If q(2) = "s" Then
                    FillSpans c
                Else
                    For Each o In Split(q(3), "/"): c.AddItem U(CStr(o)): Next o
                End If
                keep = Remembered(CStr(q(0)), "0")
                If c.ListCount > 0 Then c.ListIndex = IIf(Val(keep) < c.ListCount, Val(keep), 0)
            End If
            c.Font.Name = "Segoe UI": c.Font.Size = FONT_BASE
            c.SpecialEffect = 0: c.BorderStyle = 1: c.BorderColor = C_LINE
        Next i
    End If
    If zoomNow > 0 Then PlaceFields
End Sub

Private Sub FillSpans(c As Object)
    Dim r As Range, i As Long
    On Error Resume Next
    Set r = ThisWorkbook.nameS("G_HESAP_ACIKLIK").RefersToRange
    On Error GoTo 0
    If r Is Nothing Then Exit Sub
    For i = 1 To r.Cells.Count
        If IsNumeric(r.Cells(i).Value) And Not IsEmpty(r.Cells(i).Value) Then
            If r.Cells(i).Value > 0 Then c.AddItem U("A\u00E7\u0131kl\u0131k ") & i & ": " & r.Cells(i).Value & " m"
        End If
    Next i
End Sub

Private Function Remembered(ByVal key As String, ByVal def As String) As String
    Dim n As String
    n = "PNL_" & curTool & "_" & key
    Remembered = def
    On Error Resume Next
    Remembered = CStr(ThisWorkbook.CustomDocumentProperties(n).Value)
End Function

Private Sub Remember(ByVal key As String, ByVal v As String)
    Dim n As String
    n = "PNL_" & curTool & "_" & key
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties(n).Value = v
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.CustomDocumentProperties.Add Name:=n, LinkToContent:=False, Type:=4, Value:=v
    End If
End Sub

' reads the fields; False if a number is wrong
Private Function ReadFields() As Boolean
    Dim i As Long, c As Object, t As String
    Set vals = CreateObject("Scripting.Dictionary")
    For i = 0 To fCnt - 1
        Set c = pForm.Controls("p_" & fKey(i))
        If fTyp(i) = "n" Then
            t = Replace(Trim$(c.Text), ",", ".")
            If Len(t) = 0 Or Val(t) <= 0 Or Not IsNumeric(Replace(t, ".", "0")) Then
                UiMsg U("Say\u0131 girin: ") & pForm.Controls("q_" & fKey(i)).Caption, vbExclamation, U("Makrolar")
                Exit Function
            End If
            vals(fKey(i)) = Val(t)
            Remember fKey(i), c.Text
        Else
            vals(fKey(i)) = c.ListIndex + 1
            If fTyp(i) = "s" Then vals(fKey(i)) = SpanNo(CStr(c.Text))
            Remember fKey(i), CStr(c.ListIndex)
        End If
    Next i
    ReadFields = True
End Function

Private Function SpanNo(ByVal t As String) As Long
    Dim p As Long
    p = InStr(t, ":")
    If p > 0 Then SpanNo = Val(mid$(t, InStrRev(t, " ", p) + 1))
End Function

Public Function UiHas(ByVal key As String) As Boolean
    If Not vals Is Nothing Then UiHas = vals.Exists(key)
End Function

Public Function UiNum(ByVal key As String, ByVal def As Double) As Double
    UiNum = def
    If UiHas(key) Then UiNum = vals(key)
End Function

Public Function UiIdx(ByVal key As String, ByVal def As Long) As Long
    UiIdx = def
    If UiHas(key) Then UiIdx = vals(key)
End Function

Public Sub EvRun()
    Dim k As Long, p As Variant, sh As String, t As Single, msg As String
    ' while a macro runs, the button is Durdur: the click asks the macro to stop at its next step
    If running Then
        RequestStop
        Exit Sub
    End If
    k = curRow
    If k < 0 Then Exit Sub
    p = Split(toolList(rowTool(k)), "|")
    On Error GoTo Fail
    If StrComp(CStr(p(2)), "CheckBridgeMacroUpdates", vbTextCompare) = 0 Then
        ' UPDATE rewrites this module too, and rewriting a module whose code is running crashed
        ' Excel (2026-11-09). So the window hides and the updater starts from Excel's idle loop,
        ' after this click has returned and no PANEL code runs.
        closed = True
        pForm.Hide
        Application.OnTime Now, "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!CheckBridgeMacroUpdates"
        Exit Sub
    End If
    If Not ReadFields() Then Exit Sub
    If p(3) <> "" Then
        sh = PickSheet(CStr(p(3)))
        If sh = "" Then Exit Sub
        ThisWorkbook.Worksheets(sh).Activate
    End If
    running = True
    SetBusy True
    stepTxt = "": runBad = False: stopReq = False: stopRaised = False
    LogLine ChrW(9658), U(CStr(p(1))) & FieldText(), -1
    SetBar 0, U("\u00C7al\u0131\u015F\u0131yor...")
    t = Timer
    Application.run "'" & ThisWorkbook.Name & "'!" & p(2)
    FlushStep IIf(runBad, ChrW(215), "")
    SetBar 100, U("Bitti: ") & Format(Timer - t, "0") & " s"
    LogLine IIf(runBad Or stopReq, ChrW(215), ChrW(9632)), IIf(stopReq, U("Durduruldu: "), IIf(runBad, U("Bitti, hata var: "), U("Bitti: "))) & Format(Timer - t, "0") & " s", -1
    AddText ""
Done:
    running = False
    stepTxt = ""
    Set vals = Nothing
    SetBusy False
    Exit Sub
Fail:
    msg = Err.Description
    FlushStep ChrW(215)
    LogLine ChrW(215), U("DURDU: ") & msg, -1
    AddText ""
    SetBar 0, U("Durdu")
    Resume Done
End Sub

Private Function PickSheet(ByVal pat As String) As String
    Dim ws As Worksheet, c As New Collection, i As Long, a() As String
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like pat Then c.Add ws.Name
    Next ws
    If c.Count = 0 Then
        UiMsg U("Sayfa bulunamad\u0131: ") & pat, vbExclamation, U("Makrolar")
    ElseIf c.Count = 1 Then
        PickSheet = c(1)
    Else
        ReDim a(1 To c.Count)
        For i = 1 To c.Count: a(i) = c(i): Next i
        i = UiChoice(U("Hangi a\u00E7\u0131kl\u0131k?"), a, U("Makrolar"))
        If i > 0 Then PickSheet = a(i)
    End If
End Function

Private Sub SetBusy(ByVal busy As Boolean)
    If pForm Is Nothing Then Exit Sub
    With pForm.Controls("cmdRun")
        If busy Then
            runBack = .BackColor
            .Caption = "Durdur": .BackColor = C_RED
        Else
            .Caption = U("\u00C7al\u0131\u015Ft\u0131r")
            If runBack <> 0 Then .BackColor = runBack
        End If
        .Enabled = True
    End With
    pForm.Repaint
End Sub

' Durdur pressed: the running macro stops at its next progress step (UiProgress raises once);
' a Civil NX request already sent is finished first
Private Sub RequestStop()
    If stopReq Then Exit Sub
    stopReq = True
    With pForm.Controls("cmdRun")
        .Caption = "Duruyor...": .Enabled = False
    End With
    LogLine ChrW(215), U("Durdurma istendi: makro bir sonraki ad\u0131mda duracak (g\u00F6nderilmi\u015F bir Civil NX iste\u011Fi \u00F6nce biter)."), -1
    pForm.Repaint
End Sub

' for macros that want to check between their own steps
Public Function UiStopped() As Boolean
    DoEvents
    UiStopped = (running And stopReq)
End Function

Public Sub EvClear()
    pForm.Controls("txtLog").Text = ""
End Sub

Public Sub EvSave()
    Dim v As Variant, f As Integer
    v = Application.GetSaveAsFilename("makro_sonuc.txt", "Text (*.txt),*.txt")
    If VarType(v) = vbBoolean Then Exit Sub
    f = FreeFile
    Open CStr(v) For Output As #f
    Print #f, pForm.Controls("txtLog").Text;
    Close #f
End Sub

Public Sub EvCopy()
    With pForm.Controls("txtLog")
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
End Sub

' asked by the form before it closes itself. The form must unload itself: Unload from this module
' inside its click or close event did nothing and the error was hidden.
Public Function EvCanClose() As Boolean
    If running Then
        UiLog U("Makro \u00E7al\u0131\u015F\u0131rken pencere kapat\u0131lamaz.")
        Exit Function
    End If
    closed = True
    EvCanClose = True
End Function

' ---- what the other modules call ----
Private Function PanelOpen() As Boolean
    PanelOpen = Not pForm Is Nothing And Not closed
End Function

' every log line: time, a mark from the start of the line (x stop / fail, triangle warning,
' tick done, bullet otherwise), the text. Lines 2.. of one call are indented under it.
Public Sub UiLog(ByVal txt As String)
    Dim ln As Variant, first As Boolean, mk As String
    If Not PanelOpen() Then Exit Sub
    If Len(Trim$(Replace(Replace(txt, vbCr, ""), vbLf, ""))) = 0 Then AddText "": Exit Sub
    first = True
    For Each ln In Split(Replace(txt, vbCrLf, vbLf), vbLf)
        mk = MarkOf(CStr(ln))
        If mk = ChrW(215) And running Then runBad = True
        If first Then
            FlushStep IIf(mk = ChrW(215), mk, "")
            LogLine IIf(mk = "", ChrW(8226), mk), CStr(ln), -1
            first = False
        ElseIf Len(Trim$(ln)) > 0 Then
            AddText Space$(9) & IIf(mk = "", " ", mk) & " " & ln
        End If
    Next ln
    pForm.Repaint
    DoEvents
End Sub

' progress bar; during a run from the window a new step text also closes the last step in the log.
' After Durdur it raises one error, so the macro stops through its own error handling (cleanup runs).
Public Sub UiProgress(ByVal pct As Double, ByVal txt As String)
    If Not PanelOpen() Then Exit Sub
    If running And stopReq And Not stopRaised Then
        stopRaised = True
        FlushStep ChrW(215)
        Err.Raise vbObjectError + 900, "PANEL", U("Durduruldu (Durdur d\u00FC\u011Fmesi). Yar\u0131m kalan bir MIDAS modeli eksiktir: yeniden kurun.")
    End If
    If running And txt <> stepTxt Then
        FlushStep ""
        stepTxt = txt: stepT = Timer
    End If
    SetBar pct, txt
End Sub

' live log: the open step with its mark (up = sent, down = read, tick = other) and its time
Private Sub FlushStep(ByVal mark As String)
    Dim d As Double
    If stepTxt = "" Then Exit Sub
    If mark = "" Then mark = StepMark(stepTxt)
    d = Timer - stepT
    If d < 0 Then d = d + 86400#
    LogLine mark, stepTxt, d
    stepTxt = ""
End Sub

Private Function StepMark(ByVal s As String) As String
    s = LCase$(s)
    If InStr(s, "read") > 0 Or InStr(s, "check") > 0 Or InStr(s, "oku") > 0 Then
        StepMark = ChrW(8595)
    ElseIf InStr(s, "table") > 0 Or InStr(s, "import") > 0 Or InStr(s, "delete") > 0 Or InStr(s, "send") > 0 Then
        StepMark = ChrW(8593)
    Else
        StepMark = ChrW(8730)
    End If
End Function

' mark of a message line, read from its start only (INPUT!T3 inside a line is no warning)
Private Function MarkOf(ByVal s As String) As String
    Dim u As String
    u = UCase$(Trim$(s))
    If Left$(u, 4) = "FAIL" Or Left$(u, 7) = "STOPPED" Or Left$(u, 5) = "DURDU" Or Left$(u, 10) = "NOT SAVED:" Or _
       Right$(u, 4) = "FAIL" Or InStr(u, ": FAIL") > 0 Or InStr(u, "  FAIL") > 0 Then
        MarkOf = ChrW(215)
    ElseIf Left$(u, 1) = "!" Or Left$(u, 4) = "WARN" Or Left$(u, 5) = "UYARI" Or Left$(u, 9) = "NOT SAVED" Then
        MarkOf = ChrW(9650)
    ElseIf Left$(u, 4) = "DONE" Or Left$(u, 5) = "SAVED" Or Left$(u, 5) = "BITTI" Then
        MarkOf = ChrW(8730)
    End If
End Function

Private Sub LogLine(ByVal mark As String, ByVal txt As String, ByVal secs As Double)
    Dim s As String
    s = Format$(Now, "hh:nn:ss") & " " & mark & " " & txt
    If secs >= 0 Then
        If Len(s) < 58 Then s = s & Space$(58 - Len(s))
        s = s & " " & Format$(secs, "0.0") & " s"
    End If
    AddText s
End Sub

Private Sub AddText(ByVal s As String)
    If Not PanelOpen() Then Exit Sub
    With pForm.Controls("txtLog")
        .Text = .Text & s & vbLf
        .SelStart = Len(.Text)
    End With
End Sub

' the field values of the tool, for the start line of a run
Private Function FieldText() As String
    Dim i As Long, s As String
    On Error Resume Next
    For i = 0 To fCnt - 1
        s = s & IIf(s = "", "", ", ") & pForm.Controls("q_" & fKey(i)).Caption & " " & pForm.Controls("p_" & fKey(i)).Text
    Next i
    If s <> "" Then FieldText = "  (" & s & ")"
End Function

Private Sub SetBar(ByVal pct As Double, ByVal txt As String)
    If Not PanelOpen() Then Exit Sub
    If pct < 0 Then pct = 0
    If pct > 100 Then pct = 100
    barPct = pct
    pForm.Controls("lblBar").Width = pForm.Controls("lblBarBack").Width * pct / 100 + 0.1
    pForm.Controls("lblStep").Caption = Format(pct, "0") & " %   " & txt
    pForm.Repaint
    DoEvents
End Sub

' same arguments as MsgBox. With the window open a plain message goes to the log.
Public Function UiMsg(ByVal prompt As String, Optional ByVal buttons As Long = vbOKOnly, Optional ByVal title As String = "") As Long
    Dim kind As Long, style As Long
    kind = buttons And 7
    style = buttons And 112
    If kind = vbOKOnly And running And stopReq And PanelOpen() Then UiLog prompt: UiMsg = vbOK: Exit Function
    If kind = vbOKOnly And style <> vbCritical And style <> vbExclamation And PanelOpen() Then
        UiLog prompt
        UiMsg = vbOK
        Exit Function
    End If
    If PanelOpen() And kind = vbOKOnly Then UiLog prompt
    AskDialog prompt, IIf(title = "", "Makro", title), kind, style, False, "", Empty, ((buttons And 768) \ 256) + 1
    Select Case kind
        Case vbYesNo: UiMsg = IIf(uiBtn = 1, vbYes, vbNo)
        Case vbYesNoCancel: UiMsg = Choose(uiBtn + 1, vbCancel, vbYes, vbNo, vbCancel)
        Case vbOKCancel: UiMsg = IIf(uiBtn = 1, vbOK, vbCancel)
        Case Else: UiMsg = vbOK
    End Select
    If kind <> vbOKOnly And PanelOpen() Then LogLine "?", FirstLine(prompt) & "  -> " & AnswerText(UiMsg), -1
End Function

Private Function FirstLine(ByVal s As String) As String
    s = Replace(s, vbCr, "")
    If InStr(s, vbLf) > 0 Then s = Left$(s, InStr(s, vbLf) - 1)
    FirstLine = s
End Function

Private Function AnswerText(ByVal a As Long) As String
    Select Case a
        Case vbYes: AnswerText = "Evet"
        Case vbNo: AnswerText = U("Hay\u0131r")
        Case vbOK: AnswerText = "Tamam"
        Case Else: AnswerText = U("\u0130ptal")
    End Select
End Function

' like Application.InputBox: False on Cancel; number when isNumber
Public Function UiInput(ByVal prompt As String, Optional ByVal title As String = "", Optional ByVal def As Variant = "", Optional ByVal isNumber As Boolean = False) As Variant
    Do
        AskDialog prompt, IIf(title = "", "Makro", title), vbOKCancel, vbQuestion, True, CStr(def), Empty
        If uiBtn <> 1 Then UiInput = False: Exit Function
        If Not isNumber Then UiInput = uiText: Exit Function
        If IsNumeric(Replace(Replace(uiText, ",", ""), ".", "")) Then
            UiInput = Val(Replace(uiText, ",", "."))
            Exit Function
        End If
    Loop
End Function

' like InputBox: empty text on Cancel
Public Function UiInputS(ByVal prompt As String, Optional ByVal title As String = "", Optional ByVal def As String = "") As String
    AskDialog prompt, IIf(title = "", "Makro", title), vbOKCancel, vbQuestion, True, def, Empty
    If uiBtn = 1 Then UiInputS = uiText
End Function

' list of choices, position counted from 1, 0 on Cancel
Public Function UiChoice(ByVal prompt As String, items As Variant, Optional ByVal title As String = "", Optional ByVal def As Long = 1) As Long
    pickDef = def
    AskDialog prompt, IIf(title = "", "Makro", title), vbOKCancel, vbQuestion, False, "", items
    If uiBtn = 1 Then UiChoice = uiPick
End Function

Public Sub AskDone(ByVal btn As Long, ByVal txt As String, ByVal pick As Long)
    uiBtn = btn: uiText = txt: uiPick = pick
End Sub

Private Sub AskDialog(ByVal prompt As String, ByVal title As String, ByVal kind As Long, ByVal style As Long, ByVal withInput As Boolean, ByVal def As String, items As Variant, Optional ByVal defBtn As Long = 1)
    Dim f As Object, fr As Object, lb As Object, ln As Variant, ver As String, ask As String, t As String
    Dim i As Long, last As Long, tone As Long, tp As Double, z As Double, hMax As Double, bodyH As Double, withList As Boolean
    EnsureForms
    Set f = VBA.UserForms.Add(F_ASK)
    withList = IsArray(items)
    z = AskZoom()
    f.Zoom = CLng(z * 100)
    f.Caption = title
    ' the version line goes to the header, a closing question to the box above the buttons
    ln = Split(Replace(prompt, vbCrLf, vbLf), vbLf)
    last = -1
    For i = 0 To UBound(ln)
        t = Trim$(ln(i))
        If IsVersionLine(t) Then ver = t: ln(i) = "": t = ""
        If t <> "" Then last = i
    Next i
    If last >= 0 And kind <> vbOKOnly And Not withInput And Not withList Then
        t = Trim$(ln(last))
        If Right$(t, 1) = "?" Then ask = t: ln(last) = ""
    End If
    tone = AskTone(style, ln)
    AskHead f, title, tone, ver, withInput Or withList
    f.Controls("txtMsg").Text = Replace(prompt, vbCrLf, vbLf)
    ' body: rows when they fit. A long text goes to a text box: many rows paint slowly and a frame does not
    ' scroll with the wheel.
    Set fr = f.Controls("fraBody")
    tp = ASK_HEAD + 12
    hMax = 420 / z
    If AskEstimate(ln) > hMax Then
        fr.Visible = False
        bodyH = hMax
        With f.Controls("txtLong")
            .Text = AskPlain(ln)
            .top = tp
            .Height = hMax
            .Visible = True
            .SelStart = 0
        End With
        tp = tp + hMax + 10
    Else
        bodyH = AskRows(fr, ln)
        fr.Visible = (bodyH > 0)
        If bodyH > 0 Then
            fr.top = tp
            fr.Height = bodyH + 2
            tp = tp + fr.Height + 10
        End If
    End If
    ' question
    If ask <> "" Then tp = AskCallout(f, ask, tone, tp)
    ' input
    With f.Controls("txtIn")
        .Visible = withInput
        .top = tp
        .Text = def
    End With
    If withInput Then tp = tp + 36
    Set lb = f.Controls("lstPick")
    lb.Visible = withList
    If withList Then
        lb.Clear
        For i = LBound(items) To UBound(items): lb.AddItem CStr(items(i)): Next i
        lb.ListIndex = IIf(pickDef >= 1 And pickDef <= lb.ListCount, pickDef - 1, 0)
        lb.top = tp
        tp = tp + 36
    End If
    ' buttons
    f.Controls("lblRule").top = tp + 2
    tp = tp + 12
    Select Case kind
        Case vbYesNo
            AskBtn f, CHK & U("  Evet"), CRS & U("  Hay\u0131r"), "", tp
        Case vbYesNoCancel
            AskBtn f, CHK & U("  Evet"), CRS & U("  Hay\u0131r"), U("\u0130ptal"), tp
        Case vbOKCancel
            AskBtn f, CHK & U("  Tamam"), U("\u0130ptal"), "", tp
        Case Else
            AskBtn f, CHK & U("  Tamam"), "", "", tp
    End Select
    AskKeys f, kind, defBtn, tone
    If withInput Then f.tag = "txtIn"
    If withList Then f.tag = "lstPick"
    f.Controls("cmdCopy").Visible = (bodyH > 0 And Not withInput And Not withList)
    f.Controls("cmdCopy").top = tp
    On Error Resume Next
    If withInput Then f.Controls("txtIn").TabIndex = 0
    If withList Then f.Controls("lstPick").TabIndex = 0
    On Error GoTo 0
    f.Width = ASK_W * z + 6
    f.Height = (tp + 46) * z + 34
    uiBtn = 0: uiText = "": uiPick = 0
    f.Show vbModal
    Unload f
End Sub

' 0 question, 1 information, 2 finished, 3 warning, 4 error
Private Function AskTone(ByVal style As Long, ln As Variant) As Long
    Dim i As Long, t As String
    Select Case style
        Case vbCritical: AskTone = 4
        Case vbExclamation: AskTone = 3
        Case vbQuestion: AskTone = 0
        Case Else
            AskTone = 1
            For i = 0 To UBound(ln)
                t = Trim$(ln(i))
                If t <> "" Then
                    If LineClass(t, False) = L_OK Then AskTone = 2
                    Exit For
                End If
            Next i
    End Select
End Function

Private Sub AskHead(f As Object, ByVal title As String, ByVal tone As Long, ByVal ver As String, ByVal isInput As Boolean)
    Dim col As Long, g As String, s As String
    Select Case tone
        Case 4: col = C_RED: g = CRS: s = U("\u0130\u015Flem durdu")
        Case 3: col = C_AMBER: g = U("\u26A0"): s = U("Uyar\u0131")
        Case 2: col = C_GREEN: g = CHK: s = U("Tamamland\u0131")
        Case 1: col = C_DARK: g = U("\u2139"): s = "Bilgi"
        Case Else: col = C_DARK: g = "?": s = IIf(isInput, U("Giri\u015F gerekiyor"), U("Onay\u0131n\u0131z gerekiyor"))
    End Select
    f.Controls("lblBand").BackColor = col
    With f.Controls("lblIcon")
        .Caption = g
        .ForeColor = col
    End With
    f.Controls("lblHead").Caption = title
    f.Controls("lblSub").Caption = s
    f.Controls("lblVer").Caption = ver
End Sub

' height the rows would need
Private Function AskEstimate(ln As Variant) As Double
    Dim i As Long, t As String, h As Double
    For i = 0 To UBound(ln)
        t = Trim$(ln(i))
        If t = "" Then h = h + 6 Else h = h + 8 + 18 * (1 + Len(t) \ 95)
    Next i
    AskEstimate = h
End Function

' the same lines with their marks, for the text box
Private Function AskPlain(ln As Variant) As String
    Dim i As Long, t As String, g As String, s As String, cls As Long, sec As Boolean
    For i = 0 To UBound(ln)
        t = Trim$(ln(i))
        If t = "" Then
            sec = False
            s = s & vbCrLf
        Else
            cls = LineClass(t, sec)
            If cls = L_HEAD And InStr(UCase$(t), "WARNING") > 0 Then sec = True
            Select Case cls
                Case L_OK: g = CHK & "  "
                Case L_WARN: g = U("\u26A0") & "  "
                Case L_BAD: g = CRS & "  "
                Case L_PICK: g = U("\u25B6") & "  "
                Case L_HEAD: g = ""
                Case Else: g = U("\u2022") & "  "
            End Select
            s = s & g & t & vbCrLf
        End If
    Next i
    AskPlain = s
End Function

' draws the lines as rows inside the frame, returns their height
Private Function AskRows(fr As Object, ln As Variant) As Double
    Dim i As Long, n As Long, cls As Long, t As String, y As Double, sec As Boolean, gap As Boolean
    gap = True
    For i = 0 To UBound(ln)
        t = Trim$(ln(i))
        If t = "" Then
            If Not gap Then y = y + 6: gap = True
            sec = False
        Else
            cls = LineClass(t, sec)
            If cls = L_HEAD And InStr(UCase$(t), "WARNING") > 0 Then sec = True
            y = y + AskRow(fr, n, cls, t, y)
            n = n + 1
            gap = False
        End If
    Next i
    AskRows = y
End Function

Private Function AskRow(fr As Object, ByVal n As Long, ByVal cls As Long, ByVal t As String, ByVal y As Double) As Double
    Dim bg As Object, ic As Object, a As Object, b As Object, sk As String, sv As String
    Dim w As Double, h As Double, x As Double, g As String, fore As Long, fill As Long
    w = ASK_W - 48
    fore = C_TEXT
    fill = IIf(n Mod 2 = 1, C_STRIPE, C_WHITE)
    Select Case cls
        Case L_OK: g = CHK: fore = C_GREEN
        Case L_WARN: g = U("\u26A0"): fore = C_AMBER: fill = C_AMBERBG
        Case L_BAD: g = CRS: fore = C_RED: fill = C_REDBG
        Case L_PICK: g = U("\u25B6"): fore = C_ACC: fill = C_BLUEBG
        Case L_HEAD: g = "": fore = C_DARK: fill = C_WHITE
        Case Else: g = U("\u2022"): fore = C_MUTED
    End Select
    If cls = L_WARN And Left$(t, 2) = "- " Then t = mid$(t, 3)
    Set bg = RowLabel(fr, "b" & n, 0, y, ASK_W - 30, 10, "")
    bg.BackStyle = 1: bg.BackColor = fill
    If g <> "" Then
        Set ic = RowLabel(fr, "i" & n, 8, y + 3, 20, 20, g)
        ic.Font.Name = "Segoe UI Symbol"
        ic.ForeColor = fore
        ic.TextAlign = 2
    End If
    x = IIf(cls = L_HEAD, 10, 32)
    If SplitKey(t, cls, sk, sv) Then
        Set a = RowLabel(fr, "k" & n, x, y + 4, ASK_KEY, 18, sk)
        a.Font.Bold = True
        a.ForeColor = IIf(cls = L_PLAIN Or cls = L_OK, C_DARK, fore)
        h = FitLabel(a, ASK_KEY)
        Set b = RowLabel(fr, "v" & n, x + ASK_KEY + 6, y + 4, w - x - ASK_KEY - 14, 18, sv)
        If FitLabel(b, w - x - ASK_KEY - 14) > h Then h = b.Height
    Else
        Set b = RowLabel(fr, "v" & n, x, y + 4, w - x - 8, 18, t)
        If cls = L_HEAD Then
            b.Font.Bold = True
            b.ForeColor = IIf(InStr(UCase$(t), "WARNING") > 0, C_AMBER, C_DARK)
        End If
        h = FitLabel(b, w - x - 8)
    End If
    bg.Height = h + 8
    AskRow = h + 8
End Function

Private Function RowLabel(fr As Object, ByVal nm As String, ByVal l As Double, ByVal tp As Double, ByVal w As Double, ByVal h As Double, ByVal txt As String) As Object
    Dim x As Object
    Set x = fr.Controls.Add("Forms.Label.1", nm, True)
    x.BackStyle = 0
    x.Font.Name = "Segoe UI"
    x.Font.Size = 11
    x.ForeColor = C_TEXT
    x.WordWrap = True
    x.Left = l: x.top = tp: x.Width = w: x.Height = h
    x.Caption = txt
    Set RowLabel = x
End Function

' height of a wrapped label at this width
Private Function FitLabel(x As Object, ByVal w As Double) As Double
    x.AutoSize = True
    x.AutoSize = False
    x.Width = w
    If x.Height < 18 Then x.Height = 18
    FitLabel = x.Height
End Function

Private Function AskCallout(f As Object, ByVal q As String, ByVal tone As Long, ByVal tp As Double) As Double
    Dim bk As Object, ic As Object, tx As Object, h As Double, fill As Long, col As Long
    fill = IIf(tone = 3, C_AMBERBG, IIf(tone = 4, C_REDBG, C_BLUEBG))
    col = IIf(tone = 3, C_AMBER, IIf(tone = 4, C_RED, C_ACC))
    Set bk = f.Controls("lblAskBack")
    Set ic = f.Controls("lblAskIcon")
    Set tx = f.Controls("lblAskTxt")
    tx.Caption = q
    tx.ForeColor = col
    h = FitLabel(tx, ASK_W - 76)
    bk.top = tp: bk.Height = h + 14: bk.BackColor = fill: bk.BorderColor = col
    ic.top = tp + 5: ic.ForeColor = col
    tx.top = tp + 7
    bk.Visible = True: ic.Visible = True: tx.Visible = True
    AskCallout = tp + h + 24
End Function

Private Sub AskBtn(f As Object, ByVal a As String, ByVal b As String, ByVal c As String, ByVal tp As Double)
    Dim x As Double
    x = ASK_W - 14
    If c <> "" Then x = x - ASK_BTN: f.Controls("cmd3").Left = x: x = x - 8
    If b <> "" Then x = x - ASK_BTN: f.Controls("cmd2").Left = x: x = x - 8
    x = x - ASK_BTN: f.Controls("cmd1").Left = x
    f.Controls("cmd1").Caption = a
    f.Controls("cmd2").Caption = b: f.Controls("cmd2").Visible = (b <> "")
    f.Controls("cmd3").Caption = c: f.Controls("cmd3").Visible = (c <> "")
    f.Controls("cmd1").top = tp: f.Controls("cmd2").top = tp: f.Controls("cmd3").top = tp
End Sub

' Enter = the default button of the MsgBox call, Esc = the way out. The default button is the filled one.
Private Sub AskKeys(f As Object, ByVal kind As Long, ByVal defBtn As Long, ByVal tone As Long)
    Dim n As Long, i As Long, c As Object
    n = IIf(kind = vbYesNoCancel, 3, IIf(kind = vbOKOnly, 1, 2))
    If defBtn > n Or defBtn < 1 Then defBtn = 1
    For i = 1 To 3
        Set c = f.Controls("cmd" & i)
        c.Default = (i = defBtn)
        c.Cancel = (i = n)
        c.BackColor = IIf(i = defBtn, IIf(tone = 4, C_RED, C_ACC), C_WHITE)
        c.ForeColor = IIf(i = defBtn, C_WHITE, C_TEXT)
    Next i
    f.tag = "cmd" & defBtn
    On Error Resume Next
    f.Controls("cmd" & defBtn).TabIndex = 0
End Sub

' A text box draws its scroll bar only after it had the focus once, so it gets it here and the focus then
' goes to the default button (or the input).
Public Sub EvAskShown(f As Object)
    On Error Resume Next
    If f.Controls("txtLong").Visible Then
        f.Controls("txtLong").SetFocus
        f.Controls("txtLong").SelStart = 0
    End If
    If f.tag <> "" Then f.Controls(f.tag).SetFocus
End Sub

Public Sub EvAskCopy(f As Object)
    With f.Controls("txtMsg")
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
        .Copy
    End With
    f.Controls("cmdCopy").Caption = U("Kopyaland\u0131")
    f.Controls("cmdCopy").SetFocus
End Sub

' "macro-id 2026-11-06a"
Private Function IsVersionLine(ByVal t As String) As Boolean
    Dim p As Long
    p = InStr(t, " ")
    If p < 3 Then Exit Function
    If InStr(p + 1, t, " ") > 0 Then Exit Function
    IsVersionLine = (Left$(t, p - 1) Like "*-*") And (mid$(t, p + 1) Like "20##-##-##*")
End Function

' a tick only where the line itself reports a passed check; a plain fact gets a bullet
Private Function LineClass(ByVal t As String, ByVal inWarn As Boolean) As Long
    Dim s As String
    s = " " & UCase$(t) & " "
    If Left$(t, 6) = "YES = " Or Left$(t, 5) = "NO = " Then LineClass = L_PICK: Exit Function
    If Right$(t, 1) = ":" And Len(t) <= 60 Then LineClass = L_HEAD: Exit Function
    If InStr(s, " FAIL") > 0 Or InStr(s, " STOPPED") > 0 Or InStr(s, " ERROR") > 0 Then LineClass = L_BAD: Exit Function
    If InStr(s, "WARNING") > 0 Or (inWarn And Left$(t, 2) = "- ") Then LineClass = L_WARN: Exit Function
    If InStr(s, " NOT ") > 0 Or InStr(s, " NO ") > 0 Then Exit Function
    If InStr(s, " OK ") > 0 Or InStr(s, " OK,") > 0 Or InStr(s, " OK.") > 0 Or InStr(s, " OK:") > 0 Or InStr(s, " OK)") > 0 Then LineClass = L_OK: Exit Function
    If InStr(s, " DONE") > 0 Or InStr(s, " WRITTEN") > 0 Or InStr(s, " SAVED") > 0 Then LineClass = L_OK: Exit Function
    If InStr(s, " MATCH") > 0 Or InStr(s, " IN ORDER") > 0 Or InStr(s, " UP TO DATE") > 0 Then LineClass = L_OK
End Function

Private Function SplitKey(ByVal t As String, ByVal cls As Long, sk As String, sv As String) As Boolean
    Dim p As Long
    If cls = L_HEAD Or cls = L_WARN Then Exit Function
    If cls = L_PICK Then
        p = InStr(t, " = ")
        sk = IIf(Left$(t, 3) = "YES", "Evet", U("Hay\u0131r"))
        sv = mid$(t, p + 3)
        SplitKey = True
        Exit Function
    End If
    p = InStr(t, ": ")
    If p < 3 Or p > ASK_KEYLEN Then Exit Function
    If Left$(t, 1) = "(" Or InStr(Left$(t, p), "\") > 0 Then Exit Function
    sk = Left$(t, p - 1)
    sv = mid$(t, p + 2)
    SplitKey = True
End Function

' dialogs follow the zoom of the window
Private Function AskZoom() As Double
    AskZoom = 1
    If PanelOpen() And zoomNow > 1 Then AskZoom = zoomNow
End Function

' ---- building the forms ----
Private Sub EnsureForms()
    If FormVersion(F_PANEL) = FORM_VERSION And FormVersion(F_ASK) = FORM_VERSION Then Exit Sub
    DropForm F_PANEL
    DropForm F_ASK
    BuildAsk
    BuildPanel
End Sub

Private Function FormVersion(ByVal nm As String) As String
    Dim c As Object
    On Error Resume Next
    Set c = ThisWorkbook.VBProject.VBComponents(nm)
    If Not c Is Nothing Then FormVersion = mid$(c.CodeModule.lines(1, 1), 3)
End Function

Private Sub DropForm(ByVal nm As String)
    Dim c As Object
    On Error Resume Next
    Set c = ThisWorkbook.VBProject.VBComponents(nm)
    If Not c Is Nothing Then ThisWorkbook.VBProject.VBComponents.Remove c
End Sub

Private Function Ctl(d As Object, ByVal typ As String, ByVal nm As String, ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal h As Double) As Object
    Set Ctl = d.Controls.Add("Forms." & typ & ".1", nm, True)
    Ctl.Left = l: Ctl.top = t: Ctl.Width = w: Ctl.Height = h
    On Error Resume Next
    Ctl.Font.Name = "Segoe UI"
    Ctl.Font.Size = FONT_BASE
End Function

Private Function NewForm(ByVal nm As String, ByVal w As Double, ByVal h As Double) As Object
    Dim c As Object
    Set c = ThisWorkbook.VBProject.VBComponents.Add(3)
    c.Name = nm
    c.Properties("Width") = w
    c.Properties("Height") = h
    c.Properties("BackColor") = C_BACK
    c.Properties("ShowModal") = False
    Set NewForm = c
End Function

Private Sub BuildAsk()
    Dim c As Object, d As Object, x As Object, s As String
    Set c = NewForm(F_ASK, ASK_W, 240)
    Set d = c.Designer
    Set x = Ctl(d, "Label", "lblBand", 0, 0, ASK_W, ASK_HEAD)
    x.BackColor = C_DARK: x.Caption = ""
    Set x = Ctl(d, "Label", "lblIcon", 14, 12, 38, 38)
    x.BackColor = C_WHITE: x.Font.Name = "Segoe UI Symbol": x.Font.Size = 20: x.Font.Bold = True: x.TextAlign = 2
    Set x = Ctl(d, "Label", "lblHead", 64, 8, ASK_W - 200, 26)
    x.BackStyle = 0: x.ForeColor = C_WHITE: x.Font.Bold = True: x.Font.Size = 15
    Set x = Ctl(d, "Label", "lblSub", 64, 35, ASK_W - 200, 20)
    x.BackStyle = 0: x.ForeColor = C_LINE: x.Font.Size = 10.5
    Set x = Ctl(d, "Label", "lblVer", ASK_W - 214, 35, 200, 20)
    x.BackStyle = 0: x.ForeColor = C_LINE: x.Font.Size = 10: x.TextAlign = 3
    Set x = Ctl(d, "Frame", "fraBody", 14, 74, ASK_W - 28, 80)
    x.Caption = "": x.BackColor = C_WHITE: x.BorderStyle = 1: x.BorderColor = C_LINE: x.SpecialEffect = 0
    Set x = Ctl(d, "TextBox", "txtLong", 14, 74, ASK_W - 28, 80)
    x.MultiLine = True: x.WordWrap = True: x.Locked = True: x.ScrollBars = 2: x.BorderStyle = 1: x.BorderColor = C_LINE: x.SpecialEffect = 0: x.BackColor = C_WHITE
    x.Font.Name = "Segoe UI Symbol": x.Font.Size = 11: x.Visible = False
    Set x = Ctl(d, "Label", "lblAskBack", 14, 160, ASK_W - 28, 34)
    x.Caption = "": x.BorderStyle = 1: x.Visible = False
    Set x = Ctl(d, "Label", "lblAskIcon", 22, 165, 24, 24)
    x.BackStyle = 0: x.Caption = "?": x.Font.Name = "Segoe UI Symbol": x.Font.Size = 14: x.Font.Bold = True: x.TextAlign = 2: x.Visible = False
    Set x = Ctl(d, "Label", "lblAskTxt", 50, 167, ASK_W - 76, 20)
    x.BackStyle = 0: x.WordWrap = True: x.Font.Size = 12: x.Font.Bold = True: x.Visible = False
    ' holds the plain text for Kopyala, outside the visible area
    Set x = Ctl(d, "TextBox", "txtMsg", -3000, 0, 100, 20)
    x.MultiLine = True
    Set x = Ctl(d, "TextBox", "txtIn", 14, 110, ASK_W - 28, 26)
    x.BackColor = C_WHITE: x.BorderStyle = 1: x.BorderColor = C_LINE: x.SpecialEffect = 0
    Set x = Ctl(d, "ComboBox", "lstPick", 14, 110, ASK_W - 28, 26)
    x.style = 2: x.ListRows = 12: x.BorderStyle = 1: x.BorderColor = C_LINE: x.SpecialEffect = 0
    Set x = Ctl(d, "Label", "lblRule", 0, 150, ASK_W, 1)
    x.BackColor = C_LINE: x.Caption = ""
    Set x = Ctl(d, "CommandButton", "cmdCopy", 14, 160, 92, 32)
    x.Caption = "Kopyala": x.Font.Size = 10.5
    Set x = Ctl(d, "CommandButton", "cmd1", 300, 160, ASK_BTN, 32): x.Font.Name = "Segoe UI Symbol": x.Font.Bold = True
    Set x = Ctl(d, "CommandButton", "cmd2", 410, 160, ASK_BTN, 32): x.Font.Name = "Segoe UI Symbol"
    Set x = Ctl(d, "CommandButton", "cmd3", 520, 160, ASK_BTN, 32): x.Font.Name = "Segoe UI Symbol"
    s = "' " & FORM_VERSION & vbCrLf & _
        "Private Sub cmd1_Click(): PANEL.AskDone 1, txtIn.Text, lstPick.ListIndex + 1: Me.Hide: End Sub" & vbCrLf & _
        "Private Sub cmd2_Click(): PANEL.AskDone 2, """", 0: Me.Hide: End Sub" & vbCrLf & _
        "Private Sub cmd3_Click(): PANEL.AskDone 3, """", 0: Me.Hide: End Sub" & vbCrLf & _
        "Private Sub cmdCopy_Click(): PANEL.EvAskCopy Me: End Sub" & vbCrLf & _
        "Private Sub UserForm_Activate(): PANEL.EvAskShown Me: End Sub" & vbCrLf & _
        "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf & _
        "    If CloseMode = 0 Then Cancel = True: PANEL.AskDone 0, """", 0: Me.Hide" & vbCrLf & _
        "End Sub" & vbCrLf
    c.CodeModule.DeleteLines 1, c.CodeModule.CountOfLines
    c.CodeModule.AddFromString s
End Sub

Private Sub BuildPanel()
    Dim c As Object, d As Object, x As Object, s As String, i As Long
    Set c = NewForm(F_PANEL, BASE_W, BASE_H)
    Set d = c.Designer
    Set x = Ctl(d, "Label", "lblHeadBack", 0, 0, BASE_W, 52)
    x.BackColor = C_DARK: x.Caption = ""
    Set x = Ctl(d, "Label", "lblTitle", 16, 5, 500, 26)
    x.BackStyle = 0: x.ForeColor = C_WHITE: x.Font.Bold = True: x.Font.Size = 16
    Set x = Ctl(d, "Label", "lblSub", 16, 31, 600, 18)
    x.BackStyle = 0: x.ForeColor = C_LINE: x.Font.Size = 10.5
    Set x = Ctl(d, "Label", "lblVer", BASE_W - 134, 31, 120, 18)
    x.BackStyle = 0: x.ForeColor = C_LINE: x.Font.Size = 10.5: x.TextAlign = 3
    Set x = Ctl(d, "Label", "lblListBack", 14, 64, LIST_W, BASE_H - 78)
    x.BackColor = C_WHITE: x.BorderStyle = 1: x.BorderColor = C_LINE: x.Caption = ""
    For i = 0 To ROWS_MAX - 1
        Set x = Ctl(d, "Label", "r" & i, 16, 68 + i * ROW_H, LIST_W - 4, ROW_H)
        x.BackColor = C_WHITE: x.WordWrap = False: x.Caption = ""
    Next i
    Set x = Ctl(d, "Label", "lblName", RX, 60, 500, 26)
    x.BackStyle = 0: x.Font.Size = 15: x.Font.Bold = True: x.ForeColor = C_DARK
    Set x = Ctl(d, "Label", "lblDesc", RX, 86, 500, 36)
    x.BackStyle = 0: x.WordWrap = True: x.Font.Size = 11
    Set x = Ctl(d, "CommandButton", "cmdRun", BASE_W - 124, 62, 110, 38)
    x.Font.Bold = True: x.Font.Size = 13: x.BackColor = C_ACC: x.ForeColor = C_WHITE
    Set x = Ctl(d, "Label", "lblBarBack", RX, 216, 624, 9)
    x.BackColor = C_LINE: x.Caption = ""
    Set x = Ctl(d, "Label", "lblBar", RX, 216, 0.1, 9)
    x.BackColor = C_ACC: x.Caption = ""
    Set x = Ctl(d, "Label", "lblStep", RX, 228, 624, 20)
    x.BackStyle = 0: x.Font.Size = 11
    Set x = Ctl(d, "Label", "lblLog", RX, 252, 200, 20)
    x.BackStyle = 0: x.Font.Bold = True
    Set x = Ctl(d, "TextBox", "txtLog", RX, LOG_TOP, 624, 276)
    x.MultiLine = True: x.WordWrap = True: x.Locked = True: x.ScrollBars = 2: x.BorderStyle = 1: x.BorderColor = C_LINE: x.SpecialEffect = 0: x.BackColor = C_WHITE
    x.Font.Name = "Consolas": x.Font.Size = 11.5
    Set x = Ctl(d, "CommandButton", "cmdCopy", RX, 560, 90, 30)
    Set x = Ctl(d, "CommandButton", "cmdSave", RX + 96, 560, 90, 30)
    Set x = Ctl(d, "CommandButton", "cmdClear", RX + 192, 560, 90, 30)
    Set x = Ctl(d, "CommandButton", "cmdClose", BASE_W - 104, 560, 90, 30): x.Cancel = True
    s = "' " & FORM_VERSION & vbCrLf & _
        "Private Sub cmdRun_Click(): PANEL.EvRun: End Sub" & vbCrLf & _
        "Private Sub cmdCopy_Click(): PANEL.EvCopy: End Sub" & vbCrLf & _
        "Private Sub cmdSave_Click(): PANEL.EvSave: End Sub" & vbCrLf & _
        "Private Sub cmdClear_Click(): PANEL.EvClear: End Sub" & vbCrLf & _
        "Private Sub cmdClose_Click()" & vbCrLf & _
        "    If PANEL.EvCanClose() Then Unload Me" & vbCrLf & _
        "End Sub" & vbCrLf & _
        "Private Sub UserForm_Resize(): PANEL.EvResize: End Sub" & vbCrLf & _
        "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf & _
        "    If CloseMode = 0 Then" & vbCrLf & _
        "        If Not PANEL.EvCanClose() Then Cancel = True" & vbCrLf & _
        "    End If" & vbCrLf & _
        "End Sub" & vbCrLf
    For i = 0 To ROWS_MAX - 1
        s = s & "Private Sub r" & i & "_Click(): PANEL.EvRow " & i & ": End Sub" & vbCrLf
    Next i
    c.CodeModule.DeleteLines 1, c.CodeModule.CountOfLines
    c.CodeModule.AddFromString s
End Sub
