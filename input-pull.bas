Attribute VB_Name = "DEGER_CAGIRMA"
Option Explicit
' ============================================================================
'  DEGER_CAGIRMA - one INPUT pull macro for every report workbook (2026-10-30)
' ----------------------------------------------------------------------------
'  Button "DEGER CAGIR" (macro DEGER_CAGIRMA.DEGER_CAGIRMA). The same module
'  goes into every workbook; nothing in it is workbook-specific.
'
'  HOW CELLS ARE MATCHED: by defined names starting with G_.
'    The 1-2 workbook (INT-LDS) names every input, e.g. G_KOPRU_ADI,
'    G_HESAP_ACIKLIK (6 cells), G_KIRIS_ARALIK (6 cells). A workbook gives
'    the SAME name to its own cell(s), anywhere on any sheet; the macro copies
'    the values of every G_ name found in both files, cell by cell in reading
'    order (both ranges must hold the same number of cells). Names move with
'    the cells on cut / paste, so rearranging a sheet never breaks the pull.
'    New input: add the name in Formulas > Name Manager (workbook scope).
'
'  TWO FILES ARE ASKED FOR
'    1. inputs  : the 1-2 workbook -> every G_ name here except the two below
'    2. numbers : the PREVIOUS CHAPTER workbook (only if this workbook has
'                 G_ONCEKI_NUMARA) -> its G_SON_NUMARA (last Tablo / Sekil /
'                 Sayfa, 3 cells) is written into G_ONCEKI_NUMARA here.
'                 Cancel = numbers not changed.
'    3. others  : ONLY when some G_ names here were not found in the 1-2
'                 workbook - another chapter workbook that has them (e.g. 3A
'                 for G_PHI3). Fills only the missing names, so 1-2 always
'                 wins; asked again while names are still missing. A name
'                 whose cells here are all formulas is an export of this
'                 workbook (G_PHI3 in 3A) and is never asked for.
'                 Cancel = leave them as they are.
'    The same file may be chosen twice. Each dialog starts in the folder of
'    the file chosen last time (kept in the workbook's document properties).
'
'  NEVER WRITTEN: cells holding a formula here (listed when the source value
'  differs), parts of merged cells other than the first, G_SON_NUMARA here.
'  An EMPTY source cell empties the cell here (the pull mirrors the source).
'  Ranges must match in size; a single row may feed a single column.
' ============================================================================

Private Const SCRIPT_ID As String = "input-pull"
Private Const SCRIPT_VERSION As String = "2026-10-30b"
Private Const SCRIPT_CHANGELOG As String = "Third optional file for G_ names that are not in the 1-2 workbook (e.g. G_PHI3 from 3A)"

Private Const NAME_PREFIX As String = "G_"
Private Const NUM_SOURCE As String = "G_SON_NUMARA"
Private Const NUM_TARGET As String = "G_ONCEKI_NUMARA"
Private Const PROP_INPUT As String = "DEGER_CAGIRMA_INPUT_PATH"
Private Const PROP_NUMBERS As String = "DEGER_CAGIRMA_NUMARA_PATH"
Private Const PROP_OTHER As String = "DEGER_CAGIRMA_DIGER_PATH"
Private Const MSG_TITLE As String = "DEGER_CAGIRMA"

Private nWritten As Long, nSame As Long, nFormula As Long, nError As Long
Private formulaList As String

Public Sub DEGER_CAGIRMA()
    Dim wbIn As Workbook, wbNum As Workbook
    Dim openedIn As Boolean, openedNum As Boolean, busy As Boolean
    Dim pathIn As String, pathNum As String, calcMode As XlCalculation
    Dim nm As Name, sName As String, tR As Range, sR As Range
    Dim nNames As Long, missing As String, nMissing As Long, bad As String, nBad As Long
    Dim msg As String, numMsg As String
    Dim missNames() As String, wbOther As Workbook, openedOther As Boolean, pathOther As String
    Dim otherMsg As String, nFound As Long, i As Long

    On Error GoTo Fail
    nWritten = 0: nSame = 0: nFormula = 0: nError = 0: formulaList = ""

    pathIn = PickFile(GetProp(PROP_INPUT), "1/2  Choose the INPUT source (1-2 INT-LDS workbook)")
    If pathIn = "" Then Exit Sub
    Set wbIn = OpenSource(pathIn, openedIn)
    If wbIn Is Nothing Then GoTo Done

    If Not TargetRange(NUM_TARGET) Is Nothing Then
        pathNum = PickFile(GetProp(PROP_NUMBERS), "2/2  Choose the PREVIOUS CHAPTER workbook (Tablo / Sekil / Sayfa numbers) - Cancel = skip")
        If pathNum <> "" Then
            If StrComp(pathNum, wbIn.FullName, vbTextCompare) = 0 Then
                Set wbNum = wbIn
            Else
                Set wbNum = OpenSource(pathNum, openedNum)
                If wbNum Is Nothing Then GoTo Done
            End If
        End If
    End If

    calcMode = Application.Calculation
    busy = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 1. inputs
    For Each nm In ThisWorkbook.Names
        sName = ShortName(nm.Name)
        If UCase$(Left$(sName, Len(NAME_PREFIX))) = NAME_PREFIX And _
           StrComp(sName, NUM_SOURCE, vbTextCompare) <> 0 And StrComp(sName, NUM_TARGET, vbTextCompare) <> 0 Then
            Set tR = RangeOf(nm, ThisWorkbook)
            Set sR = Nothing
            If Not tR Is Nothing Then Set sR = FindRange(wbIn, sName)
            If tR Is Nothing Then
                nBad = nBad + 1: bad = bad & vbLf & "  " & sName & " here does not point to cells (#REF?)"
            ElseIf sR Is Nothing Then
                If tR.HasFormula = True Then GoTo NextName   ' all formulas here = a value this workbook exports (e.g. G_PHI3 in 3A)
                nMissing = nMissing + 1
                ReDim Preserve missNames(1 To nMissing)
                missNames(nMissing) = sName
            ElseIf Not SameShape(sR, tR) Then
                nBad = nBad + 1: bad = bad & vbLf & "  " & sName & ": " & sR.Address(False, False) & " in the source, " & tR.Address(False, False) & " here (size differs)"
            Else
                CopyCells sR, tR, sName
                nNames = nNames + 1
            End If
        End If
NextName:
    Next nm

    ' 2. Tablo / Sekil / Sayfa numbers of the previous chapter
    If wbNum Is Nothing Then
        If Not TargetRange(NUM_TARGET) Is Nothing Then numMsg = "Numbers: skipped (no previous chapter file chosen)." & vbLf
    Else
        Set tR = TargetRange(NUM_TARGET)
        Set sR = FindRange(wbNum, NUM_SOURCE)
        If sR Is Nothing Then
            numMsg = "Numbers: " & wbNum.Name & " has no " & NUM_SOURCE & " name - not changed." & vbLf
        ElseIf Not SameShape(sR, tR) Then
            numMsg = "Numbers: " & NUM_SOURCE & " has " & sR.Cells.Count & " cells, " & NUM_TARGET & " " & tR.Cells.Count & " - not changed." & vbLf
        Else
            CopyCells sR, tR, NUM_TARGET
            numMsg = "Numbers from " & wbNum.Name & ": " & JoinValues(sR) & vbLf
            SetProp PROP_NUMBERS, pathNum
        End If
    End If
    SetProp PROP_INPUT, pathIn

    ' 3. names not in the 1-2 workbook, from other chapter workbooks
    Do While nMissing > 0
        pathOther = PickFile(GetProp(PROP_OTHER), "3/3  Not in " & wbIn.Name & ": " & ListNames(missNames, nMissing) & _
                             " - choose the workbook that has them (Cancel = leave as they are)")
        If pathOther = "" Then Exit Do
        openedOther = False
        If StrComp(pathOther, wbIn.FullName, vbTextCompare) = 0 Then
            Set wbOther = wbIn
        ElseIf Not wbNum Is Nothing And StrComp(pathOther, pathNum, vbTextCompare) = 0 Then
            Set wbOther = wbNum
        Else
            Set wbOther = OpenSource(pathOther, openedOther)
        End If
        If wbOther Is Nothing Then Exit Do
        nFound = 0
        i = 1
        Do While i <= nMissing
            Set tR = TargetRange(missNames(i))
            Set sR = FindRange(wbOther, missNames(i))
            If sR Is Nothing Then
                i = i + 1
            Else
                If SameShape(sR, tR) Then
                    CopyCells sR, tR, missNames(i)
                    otherMsg = otherMsg & vbLf & "  " & missNames(i) & " from " & wbOther.Name
                Else
                    nBad = nBad + 1: bad = bad & vbLf & "  " & missNames(i) & ": " & sR.Address(False, False) & " in " & wbOther.Name & ", " & tR.Address(False, False) & " here (size differs)"
                End If
                nFound = nFound + 1
                missNames(i) = missNames(nMissing)   ' drop entry i
                nMissing = nMissing - 1
            End If
        Loop
        If openedOther Then openedOther = False: wbOther.Close SaveChanges:=False
        Set wbOther = Nothing
        If nFound = 0 Then
            MsgBox "That workbook has none of the missing names.", vbExclamation, MSG_TITLE
        Else
            SetProp PROP_OTHER, pathOther
        End If
    Loop
    For i = 1 To nMissing
        If i <= 10 Then missing = missing & vbLf & "  " & missNames(i)
    Next i

    msg = nNames & " names copied from " & wbIn.Name & ": " & nWritten & " cells changed, " & nSame & " already equal." & vbLf & numMsg
    If otherMsg <> "" Then msg = msg & "From other workbooks:" & otherMsg & vbLf
    If nMissing > 0 Then msg = msg & nMissing & " names of this workbook are not in the source (left as they are):" & missing & IIf(nMissing > 10, vbLf & "  ...", "") & vbLf
    If nBad > 0 Then msg = msg & nBad & " names could not be copied:" & bad & vbLf
    If nFormula > 0 Then msg = msg & nFormula & " cells hold a formula here and a different value in the source (kept):" & formulaList & vbLf
    If nError > 0 Then msg = msg & nError & " source cells hold an error value (skipped)." & vbLf

Done:
    If busy Then
        Application.Calculation = calcMode
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculate
    End If
    If openedOther Then openedOther = False: wbOther.Close SaveChanges:=False
    If openedNum Then openedNum = False: wbNum.Close SaveChanges:=False
    If openedIn Then openedIn = False: wbIn.Close SaveChanges:=False
    If busy Then MsgBox msg, IIf(nBad + nMissing + nFormula + nError > 0, vbExclamation, vbInformation), MSG_TITLE
    Exit Sub
Fail:
    msg = Err.Description
    On Error Resume Next
    If busy Then
        Application.Calculation = calcMode
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
    If openedOther Then wbOther.Close SaveChanges:=False
    If openedNum Then wbNum.Close SaveChanges:=False
    If openedIn Then wbIn.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "DEGER_CAGIRMA stopped: " & msg, vbCritical, MSG_TITLE
End Sub

' Copies sR -> tR cell by cell in reading order (same cell count checked by the caller).
Private Sub CopyCells(sR As Range, tR As Range, nameText As String)
    Dim sv() As Variant, i As Long, c As Range, v As Variant
    ReDim sv(1 To sR.Cells.Count)
    i = 0
    For Each c In sR.Cells
        i = i + 1
        sv(i) = c.Value
    Next c
    i = 0
    For Each c In tR.Cells
        i = i + 1
        v = sv(i)
        If c.MergeCells Then
            If c.Address <> c.MergeArea.Cells(1, 1).Address Then GoTo NextCell
        End If
        If IsError(v) Then
            nError = nError + 1
        ElseIf c.HasFormula Then
            If Not SameValue(c.Value, v) Then
                nFormula = nFormula + 1
                If nFormula <= 8 Then formulaList = formulaList & vbLf & "  " & nameText & " " & c.Worksheet.Name & "!" & c.Address(False, False)
            End If
        ElseIf SameValue(c.Value, v) Then
            nSame = nSame + 1
        Else
            WriteValue c, v
            nWritten = nWritten + 1
        End If
NextCell:
    Next c
End Sub

Private Function TargetRange(key As String) As Range
    Set TargetRange = FindRange(ThisWorkbook, key)
End Function

' Range of the name key in wb: the workbook-scope name first, else the first sheet-scope one.
Private Function FindRange(wb As Workbook, key As String) As Range
    Dim nm As Name, pass As Long
    For pass = 1 To 2
        For Each nm In wb.Names
            If StrComp(ShortName(nm.Name), key, vbTextCompare) = 0 Then
                If (pass = 1) = (TypeName(nm.Parent) = "Workbook") Then
                    Set FindRange = RangeOf(nm, wb)
                    If Not FindRange Is Nothing Then Exit Function
                End If
            End If
        Next nm
    Next pass
End Function

' Same cell count, and the same shape unless both are a single row or column.
Private Function SameShape(a As Range, b As Range) As Boolean
    If a.Cells.Count <> b.Cells.Count Then Exit Function
    If a.Areas.Count > 1 Or b.Areas.Count > 1 Then Exit Function
    If a.Rows.Count = b.Rows.Count And a.Columns.Count = b.Columns.Count Then SameShape = True: Exit Function
    SameShape = (a.Rows.Count = 1 Or a.Columns.Count = 1) And (b.Rows.Count = 1 Or b.Columns.Count = 1)
End Function

Private Function RangeOf(nm As Name, wb As Workbook) As Range
    Dim r As Range
    On Error Resume Next
    Set r = nm.RefersToRange
    On Error GoTo 0
    If r Is Nothing Then Exit Function
    If Not r.Worksheet.Parent Is wb Then Exit Function
    Set RangeOf = r
End Function

' "INPUT!G_X" (sheet scope) -> "G_X"
Private Function ShortName(fullName As String) As String
    ShortName = Mid$(fullName, InStrRev(fullName, "!") + 1)
End Function

Private Function ListNames(arr() As String, n As Long) As String
    Dim i As Long
    For i = 1 To n
        If i > 5 Then ListNames = ListNames & " ...": Exit Function
        ListNames = ListNames & IIf(i > 1, ", ", "") & arr(i)
    Next i
End Function

Private Function JoinValues(r As Range) As String
    Dim c As Range
    For Each c In r.Cells
        If Not IsError(c.Value) Then JoinValues = JoinValues & IIf(JoinValues = "", "", " / ") & CStr(c.Value)
    Next c
End Function

' File dialog starting in the folder of lastPath.
Private Function PickFile(lastPath As String, dlgTitle As String) As String
    Dim f As Variant, p As Long
    On Error Resume Next
    p = InStrRev(lastPath, "\")
    If p > 3 Then
        ChDrive Left$(lastPath, 1)
        ChDir Left$(lastPath, p - 1)
    End If
    On Error GoTo 0
    f = Application.GetOpenFilename("Excel (*.xlsm;*.xlsx),*.xlsm;*.xlsx", , dlgTitle)
    If VarType(f) = vbBoolean Then Exit Function
    PickFile = CStr(f)
End Function

' The chosen workbook: reused when already open (opened = False), else opened read-only.
Private Function OpenSource(path As String, opened As Boolean) As Workbook
    Dim wb As Workbook, nm As String
    If StrComp(path, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        MsgBox "That is this workbook itself.", vbExclamation, MSG_TITLE
        Exit Function
    End If
    nm = Mid$(path, InStrRev(path, "\") + 1)
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then Set OpenSource = wb: Exit Function
        If StrComp(wb.Name, nm, vbTextCompare) = 0 Then
            MsgBox "Another file named " & nm & " is open (" & wb.FullName & "). Close it first.", vbExclamation, MSG_TITLE
            Exit Function
        End If
    Next wb
    Set OpenSource = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True)
    opened = True
End Function

Private Function GetProp(propName As String) As String
    On Error Resume Next
    GetProp = CStr(ThisWorkbook.CustomDocumentProperties(propName).Value)
End Function

Private Sub SetProp(propName As String, v As String)
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties(propName).Value = v
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.CustomDocumentProperties.Add Name:=propName, LinkToContent:=False, Type:=4, Value:=v
    End If
End Sub

Private Function SameValue(a As Variant, b As Variant) As Boolean
    If IsError(a) Or IsError(b) Then Exit Function
    If VarType(a) <> VarType(b) Then Exit Function
    SameValue = (a = b)
End Function

' Text is always stored as text (Excel would turn "1/2", "12.05", "5%", "=x" into
' numbers, dates or formulas); a cell formatted as Text takes it as it is.
Private Sub WriteValue(t As Range, v As Variant)
    If VarType(v) = vbString Then
        If Len(v) > 0 And t.NumberFormat <> "@" Then
            t.Value = "'" & v
            Exit Sub
        End If
    End If
    t.Value = v
End Sub
