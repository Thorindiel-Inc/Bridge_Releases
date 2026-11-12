Attribute VB_Name = "KAYIP"
' KAYIP - elastic shortening loss of the girder by goal seek on a ..._KDME sheet (AASHTO LRFD 5.9.3.2.3a):
' N341 = assumed loss DfpES (MPa, typed), Z341 = computed loss - assumed loss; the goal seek makes Z341 = 0.
' Guard (owner 2026-11-07): runs on a _KDME sheet only, refuses an error or a non-number in Z341 and a formula in
' N341, remembers N341 and puts it back when the goal seek fails or does not reach 0, and says why. Before, a
' Z341 error made the goal seek write junk into N341.
Option Explicit

Private Const SCRIPT_ID As String = "kayip"
Private Const SCRIPT_VERSION As String = "2026-11-12a"
Private Const SCRIPT_CHANGELOG As String = "Cells read and written found by M_ names (MIDAS_CORE McAnchor), created at the old address when missing; Guard: _KDME sheets only, refuses an error in Z341 or a formula in N341, puts N341 back when the goal seek does not reach 0"
Private Const SCRIPT_NEEDS As String = "midas-core"   ' UPDATE adds this module when the workbook lacks it

Private Const TARGET_CELL As String = "Z341"    ' computed - assumed loss (MPa)
Private Const CHANGING_CELL As String = "N341"  ' assumed loss DfpES (MPa), typed
Private Const TOL_MPA As Double = 0.01          ' accepted |Z341| after the goal seek

' Button on the _KDME sheets (OnAction KAYIP.KAYIP). Silent when solved, a message otherwise.
Public Sub KAYIP()
    Dim msg As String
    msg = KAYIP_RUN()
    If msg <> "" Then MsgBox msg, vbExclamation, "KAYIP"
End Sub

' Scripted entry (tests): sheetName "" = active sheet. Returns "" when solved, else the reason;
' N341 is then as it was before.
Public Function KAYIP_RUN(Optional ByVal sheetName As String = "") As String
    Dim ws As Object, tgt As Range, chg As Range, oldVal As Variant, solved As Boolean, calcWas As Long, wsh As Worksheet

    If sheetName = "" Then
        Set ws = ActiveSheet
    Else
        On Error Resume Next
        Set ws = ActiveWorkbook.Worksheets(sheetName)
        On Error GoTo 0
        If ws Is Nothing Then KAYIP_RUN = "No sheet '" & sheetName & "' in this workbook.": Exit Function
    End If
    If TypeName(ws) <> "Worksheet" Then KAYIP_RUN = "KAYIP runs on a worksheet.": Exit Function
    If Not (UCase$(ws.Name) Like "*_KDME") Then
        KAYIP_RUN = "KAYIP runs on a ..._KDME sheet (girder chapter). The active sheet is '" & ws.Name & "'. Nothing was changed."
        Exit Function
    End If

    ' by the names of this sheet (owner 2026-11-12), created at Z341 / N341 when missing
    Set wsh = ws
    Set tgt = McAnchor(wsh, "M_KAYIP_HEDEF", TARGET_CELL, True)
    Set chg = McAnchor(wsh, "M_KAYIP_DEGISEN", CHANGING_CELL, True)
    If Not tgt.HasFormula Then
        KAYIP_RUN = ws.Name & "!" & tgt.Address(False, False) & " holds no formula, so there is nothing to solve. Nothing was changed."
        Exit Function
    End If
    If chg.HasFormula Then
        KAYIP_RUN = ws.Name & "!" & chg.Address(False, False) & " holds a formula; the goal seek needs a typed value there. Nothing was changed."
        Exit Function
    End If
    If IsError(tgt.Value) Then
        KAYIP_RUN = ws.Name & "!" & tgt.Address(False, False) & " shows " & tgt.Text & ", so the loss cannot be solved. Fill the inputs of this span first. " & chg.Address(False, False) & " was not changed."
        Exit Function
    End If
    If Not IsNumeric(tgt.Value) Or IsEmpty(tgt.Value) Then
        KAYIP_RUN = ws.Name & "!" & tgt.Address(False, False) & " is not a number ('" & tgt.Text & "'). Nothing was changed."
        Exit Function
    End If

    oldVal = chg.Value
    calcWas = Application.Calculation
    If calcWas <> xlCalculationAutomatic Then Application.Calculation = xlCalculationAutomatic

    On Error Resume Next
    solved = tgt.GoalSeek(Goal:=0, ChangingCell:=chg)
    If Err.Number <> 0 Then solved = False
    On Error GoTo 0
    ws.Calculate

    If solved Then
        If IsError(tgt.Value) Then
            solved = False
        ElseIf Not IsNumeric(tgt.Value) Then
            solved = False
        ElseIf Abs(CDbl(tgt.Value)) > TOL_MPA Then
            solved = False
        End If
    End If

    If Not solved Then
        chg.Value = oldVal
        ws.Calculate
        KAYIP_RUN = "The goal seek did not bring " & ws.Name & "!" & tgt.Address(False, False) & " to 0 (within " & TOL_MPA & " MPa). " & _
            chg.Address(False, False) & " was put back to " & CStr(oldVal) & ". Check the inputs of the loss block."
    End If
    If calcWas <> xlCalculationAutomatic Then Application.Calculation = calcWas
End Function
