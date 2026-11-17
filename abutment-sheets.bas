Attribute VB_Name = "KENAR_AYAK"
Option Explicit
' Abutment report workbooks (8'6_..._KA1, 8'6_..._KA2) and the foundation workbook (8'8) - 2026-11-17a
' TEMEL_SAYFALARI: shows the sheet set of this abutment's foundation type, hides the other set,
' then rebuilds the contents and the figure / table lists (RAPOR_LISTE skips hidden sheets).
'   foundation type = workbook name KA_TEMEL_TIPI (abutment workbooks) or T_TEMEL_TIPI (foundation
'   workbook), both from the pier table of the main workbook, column Temel Tip:
'     KAZIKLI  -> piled set           2A_KAZIK / 2_KAZIK, GEO_KAZIK, KAZIK_UC, KAZIK_UC_EYL, MIDAS_CIKTISI_KAZIK
'     YUZEYSEL -> spread-footing set  2A_STABILITE, GEO_YUZEYSEL
'   A sheet of a set that the workbook does not hold is skipped.
'   Nothing is deleted: a hidden set comes back when the type changes and the tool runs again.
' TEMEL_SAYFALARI_SESSIZ: the same without a message box; returns the log (scripted tests).

Private Const SCRIPT_ID As String = "abutment-sheets"
Private Const SCRIPT_VERSION As String = "2026-11-17a"
Private Const SCRIPT_CHANGELOG As String = "Also for the foundation workbook 8'8 (type from T_TEMEL_TIPI, piled sheet 2_KAZIK); Its macro-window row is declared in this module ('@TOOL line), so a changed or new tool needs no PANEL release; New: shows the sheet set of the abutment's foundation type (KAZIKLI / YUZEYSEL), hides the other set, rebuilds the lists"
' macro-window row(s) of this module (PANEL reads them): order|group|title|macro|sheet|text|fields[|modules]
'@TOOL 30|Girdiler|Temel tipine g\u00F6re sayfalar|KENAR_AYAK.TEMEL_SAYFALARI||Temel tipine (KAZIKLI / Y\u00DCZEYSEL) g\u00F6re kaz\u0131kl\u0131 veya y\u00FCzeysel temel sayfalar\u0131n\u0131 g\u00F6sterir, di\u011Ferini gizler (silmez); i\u00E7indekiler ve \u015Fekil / tablo listelerini yeniler.|

Private Const SET_KAZIKLI As String = "2A_KAZIK|2_KAZIK|GEO_KAZIK|KAZIK_UC|KAZIK_UC_EYL|MIDAS_CIKTISI_KAZIK"
Private Const SET_YUZEYSEL As String = "2A_STABILITE|GEO_YUZEYSEL"
Private Const TYPE_NAMES As String = "KA_TEMEL_TIPI|T_TEMEL_TIPI"   ' the first one the workbook holds

Public Sub TEMEL_SAYFALARI()
    Dim s As String
    s = TEMEL_SAYFALARI_SESSIZ()
    If Left$(s, 5) = "FAIL:" Then
        ShowMsg mid$(s, 6), vbExclamation
    Else
        ShowMsg s, vbInformation
    End If
End Sub

Public Function TEMEL_SAYFALARI_SESSIZ() As String
    Dim t As String, showSet As String, hideSet As String, lg As String, r As Variant
    t = TemelTipi()
    Select Case t
        Case "KAZIKLI": showSet = SET_KAZIKLI: hideSet = SET_YUZEYSEL
        Case "YUZEYSEL": showSet = SET_YUZEYSEL: hideSet = SET_KAZIKLI
        Case Else
            TEMEL_SAYFALARI_SESSIZ = "FAIL:" & U("Temel tipi okunamad\u0131: '") & t & "'." & vbLf & _
                U("KA_TEMEL_TIPI / T_TEMEL_TIPI KAZIKLI veya Y\u00DCZEYSEL olmal\u0131 (ana \u00E7al\u0131\u015Fma kitab\u0131, INPUT ayak tablosu, Temel Tip s\u00FCtunu).")
            Exit Function
    End Select
    ' show first, so the workbook always keeps a visible sheet
    lg = SetVisible(showSet, True) & SetVisible(hideSet, False)
    On Error Resume Next
    r = Application.Run("'" & Replace(ThisWorkbook.Name, "'", "''") & "'!RAPOR_LISTE.LISTELERI_AYARLA_SESSIZ")
    If Err.Number <> 0 Then
        lg = lg & U("Listeler yenilenemedi: ") & Err.Description & vbLf
        Err.Clear
    Else
        lg = lg & U("\u0130\u00E7indekiler ve \u015Fekil / tablo listeleri yenilendi.") & vbLf
        If Len(CStr(r)) > 0 Then lg = lg & CStr(r)
    End If
    On Error GoTo 0
    TEMEL_SAYFALARI_SESSIZ = U("Temel tipi: ") & t & vbLf & lg
End Function

' Foundation type in capitals without Turkish dots: KAZIKLI, YUZEYSEL, or the text found
Private Function TemelTipi() As String
    Dim v As Variant, s As String, nm As Variant, found As Boolean
    For Each nm In Split(TYPE_NAMES, "|")
        On Error Resume Next
        v = ThisWorkbook.Names(CStr(nm)).RefersToRange.Cells(1, 1).Value
        found = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
        If found Then Exit For
    Next nm
    If Not found Then TemelTipi = U("(ad yok: ") & Replace(TYPE_NAMES, "|", " / ") & ")": Exit Function
    If IsError(v) Then TemelTipi = U("(hata de\u011Feri)"): Exit Function
    s = UCase$(Trim$(CStr(v)))
    s = Replace(s, ChrW(220), "U"): s = Replace(s, ChrW(252), "U")
    s = Replace(s, ChrW(304), "I"): s = Replace(s, ChrW(305), "I")
    TemelTipi = s
End Function

Private Function SetVisible(ByVal setList As String, ByVal vis As Boolean) As String
    Dim nm As Variant, ws As Worksheet, lg As String
    For Each nm In Split(setList, "|")
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        On Error GoTo 0
        If ws Is Nothing Then
            ' not in this workbook (the sets cover the abutment and the foundation workbooks)
        ElseIf vis Then
            If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
            lg = lg & U("g\u00F6r\u00FCn\u00FCr: ") & nm & vbLf
        Else
            If ws.Visible = xlSheetVisible Then ws.Visible = xlSheetHidden
            lg = lg & U("gizli: ") & nm & vbLf
        End If
    Next nm
    SetVisible = lg
End Function

' The macro window (module PANEL) is optional: without it, a plain message box.
Private Sub ShowMsg(msg As String, icon As VbMsgBoxStyle)
    Dim title As String
    title = U("Temel tipine g\u00F6re sayfalar")
    On Error Resume Next
    Application.Run "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PANEL.UiMsg", msg, icon, title
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox msg, icon, title
    End If
End Sub

Private Function U(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, "\u")
    Do While p > 0
        s = Left$(s, p - 1) & ChrW(CLng("&H" & mid$(s, p + 2, 4))) & mid$(s, p + 6)
        p = InStr(p + 1, s, "\u")
    Loop
    U = s
End Function
