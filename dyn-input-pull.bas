Attribute VB_Name = "DEGER_CAGIRMA"
Option Explicit
' 3A (dynamic) workbook - 2026-10-30b
' Button "DEGER CAGIR" on INPUT (macro DEGER_CAGIRMA.DEGER_CAGIRMA).
' Pulls the INPUT values from the chapter 1-2 workbook (1-2_YHT_XXXXX_INT-LDS.xlsm).
' Both INPUT sheets have the same layout, so every cell of the source INPUT (A1:Z165) goes
' to the same cell here, except:
'   - the 3A-only cells in KEEP_AREAS (chapter counters L1:O5, Phi3 A92:B92,
'     selected span block H6:J7, and D43:H46 where 3A has no Enleme block)
'   - cells that hold a formula here (kept; listed if the source value differs)
' Then the last Tablo / Sekil / Sayfa numbers of 1-2 (its INPUT!M2:O2) go to M4:O4.
' The source path is remembered in INPUT!L1.
' Replaces the old DEGER_CAGIRMA (fixed LDS ranges + form Dosya_Sec).

Private Const SCRIPT_ID As String = "dyn-input-pull"
Private Const SCRIPT_VERSION As String = "2026-10-30b"
Private Const SCRIPT_CHANGELOG As String = "3A pull from the 1-2 INPUT; keeps H6:J7 and D43:H46; numbers to M4:O4"

Private Const KEEP_AREAS As String = "L1:O5,A92:B92,H6:J7,D43:H46"
Private Const START_CELL As String = "M4"    ' first Tablo / Sekil / Sayfa numbers of this chapter
Private Const COPY_AREA As String = "A1:Z165"

Public Sub DEGER_CAGIRMA()
    Dim wsT As Worksheet, wbS As Workbook, wsS As Worksheet, keep As Range
    Dim c As Range, t As Range, v As Variant, tv As Variant
    Dim path As String, opened As Boolean, busy As Boolean, calcMode As XlCalculation
    Dim nW As Long, nF As Long, nE As Long, nL As Long, i As Long
    Dim fList As String, lList As String, msg As String

    On Error GoTo Fail
    Set wsT = ThisWorkbook.Worksheets("INPUT")
    path = PickFile(CStr(wsT.Range("L1").Value))
    If path = "" Then Exit Sub

    Set wbS = OpenSource(path, opened)
    If wbS Is Nothing Then Exit Sub
    If wbS Is ThisWorkbook Then
        MsgBox "This is the 3A workbook itself; choose the 1-2 (INT-LDS) file.", vbExclamation, "DEGER_CAGIRMA"
        Exit Sub
    End If
    On Error Resume Next
    Set wsS = wbS.Worksheets("INPUT")
    On Error GoTo Fail
    If wsS Is Nothing Then
        MsgBox wbS.Name & " has no INPUT sheet.", vbExclamation, "DEGER_CAGIRMA"
        GoTo Done
    End If
    Set keep = wsT.Range(KEEP_AREAS)

    ' layout check: typed labels in column A must match
    For Each c In wsS.Range("A1:A165")
        Set t = wsT.Range(c.Address)
        If Intersect(t, keep) Is Nothing And Not c.HasFormula And Not t.HasFormula Then
            If Not IsError(c.Value) And Not IsError(t.Value) Then
                If CStr(c.Value) <> CStr(t.Value) Then
                    nL = nL + 1
                    If nL <= 8 Then lList = lList & vbLf & "  " & c.Address(False, False) & ": """ & CStr(c.Value) & """ / here """ & CStr(t.Value) & """"
                End If
            End If
        End If
    Next c
    If nL > 0 Then
        If MsgBox(nL & " labels in column A differ between the two INPUT sheets (source / here):" & lList & vbLf & vbLf & _
                  "The layouts may not match. Copy anyway?", vbYesNo + vbExclamation, "DEGER_CAGIRMA") = vbNo Then GoTo Done
    End If

    calcMode = Application.Calculation
    busy = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    For Each c In wsS.Range(COPY_AREA)
        Set t = wsT.Range(c.Address)
        If Not Intersect(t, keep) Is Nothing Then GoTo NextCell
        If t.MergeCells Then
            If t.Address <> t.MergeArea.Cells(1, 1).Address Then GoTo NextCell
        End If
        v = c.Value
        If IsError(v) Then nE = nE + 1: GoTo NextCell
        If t.HasFormula Then
            tv = t.Value
            If Not IsError(tv) Then
                If Not SameValue(tv, v) And Not c.HasFormula Then
                    nF = nF + 1
                    If nF <= 8 Then fList = fList & vbLf & "  " & t.Address(False, False)
                End If
            End If
            GoTo NextCell
        End If
        If Not SameValue(t.Value, v) Then
            WriteValue t, v
            nW = nW + 1
        End If
NextCell:
    Next c

    ' last Tablo / Sekil / Sayfa of 1-2 -> start numbers of this chapter
    For i = 0 To 2
        v = wsS.Range("M2").Offset(0, i).Value
        If Not IsError(v) And Not IsEmpty(v) And IsNumeric(v) Then
            wsT.Range(START_CELL).Offset(0, i).Value = v
        Else
            msg = msg & "Source INPUT!" & wsS.Range("M2").Offset(0, i).Address(False, False) & " is not a number; " & _
                  wsT.Range(START_CELL).Offset(0, i).Address(False, False) & " not changed." & vbLf
        End If
    Next i
    wsT.Range("L1").Value = "'" & path

Done:
    If busy Then
        Application.Calculation = calcMode
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculate
        msg = nW & " cells copied from " & wbS.Name & "." & vbLf & msg
        If nF > 0 Then msg = msg & nF & " cells have a formula here and a different typed value in the source (kept):" & fList & vbLf
        If nE > 0 Then msg = msg & nE & " source cells hold an error value (skipped)." & vbLf
    End If
    If opened Then opened = False: wbS.Close SaveChanges:=False
    If busy Then MsgBox msg, vbInformation, "DEGER_CAGIRMA"
    Exit Sub
Fail:
    msg = Err.Description
    If busy Then
        Application.Calculation = calcMode
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
    On Error Resume Next
    If opened Then wbS.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "DEGER_CAGIRMA stopped: " & msg, vbCritical, "DEGER_CAGIRMA"
End Sub

' File dialog, starting in the folder of the last used file.
Private Function PickFile(lastPath As String) As String
    Dim f As Variant, p As Long
    On Error Resume Next
    p = InStrRev(lastPath, "\")
    If p > 3 Then
        ChDrive Left$(lastPath, 1)
        ChDir Left$(lastPath, p - 1)
    End If
    On Error GoTo 0
    f = Application.GetOpenFilename("Excel (*.xlsm;*.xlsx),*.xlsm;*.xlsx", , "Choose the 1-2 workbook (INT-LDS)")
    If VarType(f) = vbBoolean Then Exit Function
    PickFile = CStr(f)
End Function

' The chosen workbook: reused when already open (opened = False), else opened read-only.
Private Function OpenSource(path As String, opened As Boolean) As Workbook
    Dim wb As Workbook, nm As String
    nm = Mid$(path, InStrRev(path, "\") + 1)
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then Set OpenSource = wb: Exit Function
        If StrComp(wb.Name, nm, vbTextCompare) = 0 Then
            MsgBox "Another file named " & nm & " is open (" & wb.FullName & "). Close it first.", vbExclamation, "DEGER_CAGIRMA"
            Exit Function
        End If
    Next wb
    Set OpenSource = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True)
    opened = True
End Function

Private Function SameValue(a As Variant, b As Variant) As Boolean
    If IsError(a) Or IsError(b) Then Exit Function
    If VarType(a) <> VarType(b) Then Exit Function
    SameValue = (a = b)
End Function

' Text that Excel would turn into a number or formula is stored as text.
Private Sub WriteValue(t As Range, v As Variant)
    If VarType(v) = vbString Then
        If Len(v) > 0 Then
            If IsNumeric(v) Or Left$(v, 1) = "=" Then
                t.Value = "'" & v
                Exit Sub
            End If
        End If
    End If
    t.Value = v
End Sub
