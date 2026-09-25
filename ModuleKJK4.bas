Attribute VB_Name = "Module4"
Option Explicit

'==================================================================
' modMultiAttr - ПУСТОЙ ШАБЛОН. Все группы (036,044,127,050,087,112,
' 032,176) переехали в Module5 (modNestedGroups) на систему сводной
' ячейки. Модуль оставлен как пустая оболочка, чтобы не трогать
' Module3 (Router), который все еще вызывает MultiAttr_OpenFromMain /
' MultiAttr_HandleClick / IsMultiAttrSheet - при CFG_COUNT=0 они
' просто ничего не делают.
'
' Если понадобится добавить НОВЫЙ атрибут по СТАРОЙ схеме (несколько
' обычных столбцов-триггеров вместо одной сводной ячейки) - можно
' сюда вернуть Case-настройку, как было раньше.
'==================================================================

Public Const MULTI_ATTR_CODE_ROW As Long = 8

Private Const CFG_COUNT As Long = 0
Private Const HEADER_ROW As Long = 4
Private Const FIRST_DATA_ROW As Long = 5
Private Const DICT_CODE_ROW_CHILD As Long = 3

Private Type TMultiAttr
    TriggerCodes As String
    SheetName As String
    TitleText As String
    LastCol As String
    CtxCol As String
    ReturnCol As String
    Headers As Variant
    Widths As Variant
    IntCols As String
    MoneyCols As String
    DictCodes As Variant
End Type

Private CurrentProjectID As String
Private CurrentSourceRow As Long

Private Sub LoadCfg(ByVal Index As Long, ByRef Cfg As TMultiAttr)
    ' пусто - все группы теперь в Module5
End Sub

Public Function MultiAttr_OpenFromMain( _
    ByVal ws As Worksheet, _
    ByVal Target As Range) As Boolean

    Dim Code As String
    Dim i As Long
    Dim Cfg As TMultiAttr

    If ws.Name <> MAIN_SHEET_NAME Then Exit Function
    If Target.Row < FIRST_INPUT_ROW Then Exit Function
    If CFG_COUNT = 0 Then Exit Function

    Code = Trim$(CStr(ws.Cells(MULTI_ATTR_CODE_ROW, Target.Column).Value2))
    If Len(Code) = 0 Then Exit Function

    For i = 1 To CFG_COUNT

        LoadCfg i, Cfg

        If InStr(1, ";" & Cfg.TriggerCodes & ";", ";" & Code & ";", vbTextCompare) > 0 Then
            MultiAttr_OpenFromMain = True
            OpenMultiAttr Cfg, ws, Target
            Exit Function
        End If

    Next i

End Function

Public Function MultiAttr_HandleClick( _
    ByVal Sh As Worksheet, _
    ByVal Target As Range) As Boolean

    Dim i As Long
    Dim Cfg As TMultiAttr

    For i = 1 To CFG_COUNT

        LoadCfg i, Cfg

        If StrComp(Sh.Name, Cfg.SheetName, vbTextCompare) = 0 Then

            Select Case Target.Address(False, False)

                Case "A2"
                    AddMultiAttrRecord Cfg
                    MultiAttr_HandleClick = True

                Case "B2"
                    ReturnFromMultiAttr Cfg
                    MultiAttr_HandleClick = True

            End Select

            Exit Function

        End If

    Next i

End Function

Public Function IsMultiAttrSheet(ByVal SheetName As String) As Boolean

    Dim i As Long
    Dim Cfg As TMultiAttr

    For i = 1 To CFG_COUNT
        LoadCfg i, Cfg
        If StrComp(Cfg.SheetName, SheetName, vbTextCompare) = 0 Then
            IsMultiAttrSheet = True
            Exit Function
        End If
    Next i

End Function

Private Sub OpenMultiAttr( _
    ByRef Cfg As TMultiAttr, _
    ByVal wsSource As Worksheet, _
    ByVal Target As Range)

    Dim wsChild As Worksheet
    Dim LastRow As Long
    Dim FirstVisibleRow As Long

    CurrentProjectID = Trim$(CStr(wsSource.Cells(Target.Row, PROJECT_ID_COL).Value2))

    If CurrentProjectID = vbNullString Then
        MsgBox "Сначала заполните KJKEMD_003 — ID проекта в столбце " & _
               PROJECT_ID_COL & ".", vbExclamation
        Exit Sub
    End If

    CurrentSourceRow = Target.Row
    Set wsChild = EnsureMultiAttrSheet(Cfg)

    wsChild.Range(Cfg.CtxCol & "1").Value2 = CurrentProjectID
    wsChild.Range(Cfg.CtxCol & "2").Value2 = CurrentSourceRow
    wsChild.Range("A1").Value2 = Cfg.TitleText & ": " & CurrentProjectID

    ShowAllMultiAttr wsChild

    EnsureFirstMultiAttrRecord Cfg, wsChild
    LastRow = LastMultiAttrRow(wsChild)

    wsChild.Range("A" & HEADER_ROW & ":" & Cfg.LastCol & LastRow).AutoFilter _
        Field:=1, Criteria1:=CurrentProjectID

    FirstVisibleRow = FirstVisibleMultiAttrRow(wsChild, LastRow)

    wsChild.Visible = xlSheetVisible
    wsChild.Activate

    If FirstVisibleRow = 0 Then
        wsChild.Range("A2").Select
    Else
        wsChild.Cells(FirstVisibleRow, "B").Select
    End If

End Sub

Private Sub EnsureFirstMultiAttrRecord(ByRef Cfg As TMultiAttr, ByVal ws As Worksheet)

    Dim LastRow As Long
    Dim NewRow As Long
    Dim ExistingCount As Long

    LastRow = LastMultiAttrRow(ws)

    If LastRow >= FIRST_DATA_ROW Then
        ExistingCount = Application.WorksheetFunction.CountIf( _
            ws.Range("A" & FIRST_DATA_ROW & ":A" & LastRow), CurrentProjectID)
    End If

    If ExistingCount > 0 Then Exit Sub

    NewRow = LastRow + 1
    ws.Cells(NewRow, "A").Value2 = CurrentProjectID

    FormatMultiAttrRow Cfg, ws, NewRow

End Sub

Private Sub AddMultiAttrRecord(ByRef Cfg As TMultiAttr)

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim NewRow As Long

    Set ws = EnsureMultiAttrSheet(Cfg)
    RestoreMultiAttrContext Cfg, ws

    If CurrentProjectID = vbNullString Then
        MsgBox "Не определен ID проекта.", vbExclamation
        Exit Sub
    End If

    ShowAllMultiAttr ws

    LastRow = LastMultiAttrRow(ws)
    NewRow = LastRow + 1

    ws.Cells(NewRow, "A").Value2 = CurrentProjectID

    FormatMultiAttrRow Cfg, ws, NewRow

    ws.Range("A" & HEADER_ROW & ":" & Cfg.LastCol & NewRow).AutoFilter _
        Field:=1, Criteria1:=CurrentProjectID

    ws.Cells(NewRow, "B").Select

End Sub

Private Sub ReturnFromMultiAttr(ByRef Cfg As TMultiAttr)

    Dim wsChild As Worksheet
    Dim wsSource As Worksheet

    Set wsChild = EnsureMultiAttrSheet(Cfg)
    RestoreMultiAttrContext Cfg, wsChild

    ShowAllMultiAttr wsChild

    If Not ValidateMultiAttrRows(Cfg, wsChild) Then Exit Sub

    wsChild.Visible = xlSheetHidden

    Set wsSource = ThisWorkbook.Worksheets(MAIN_SHEET_NAME)
    wsSource.Activate

    If CurrentSourceRow >= FIRST_INPUT_ROW Then
        wsSource.Cells(CurrentSourceRow, Cfg.ReturnCol).Select
    Else
        wsSource.Range(Cfg.ReturnCol & FIRST_INPUT_ROW).Select
    End If

End Sub

Private Function ValidateMultiAttrRows(ByRef Cfg As TMultiAttr, ByVal ws As Worksheet) As Boolean

    Dim LastRow As Long
    Dim r As Long
    Dim c As Long
    Dim LastColNum As Long
    Dim FilledCount As Long
    Dim TotalCount As Long
    Dim FirstEmptyCol As Long

    ValidateMultiAttrRows = True

    LastColNum = ws.Range(Cfg.LastCol & "1").Column
    LastRow = LastMultiAttrRow(ws)

    For r = FIRST_DATA_ROW To LastRow

        FilledCount = 0
        TotalCount = 0
        FirstEmptyCol = 0

        For c = 2 To LastColNum

            TotalCount = TotalCount + 1

            If Len(Trim$(CStr(ws.Cells(r, c).Value2))) > 0 Then
                FilledCount = FilledCount + 1
            ElseIf FirstEmptyCol = 0 Then
                FirstEmptyCol = c
            End If

        Next c

        If FilledCount > 0 And FilledCount < TotalCount Then

            MsgBox "Строка " & r & " на листе """ & ws.Name & """ заполнена не полностью." & _
                   vbCrLf & "Если начали вносить запись - заполните все столбцы, " & _
                   "либо очистите строку целиком.", vbExclamation, "Возврат отменен"

            ws.Cells(r, FirstEmptyCol).Select

            ValidateMultiAttrRows = False
            Exit Function

        End If

    Next r

End Function

Private Sub FormatMultiAttrRow(ByRef Cfg As TMultiAttr, ByVal ws As Worksheet, ByVal RowNumber As Long)

    Dim Col As Variant

    With ws.Range("A" & RowNumber & ":" & Cfg.LastCol & RowNumber)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 190, 200)
        .Interior.Color = RGB(226, 239, 218)
    End With

    For Each Col In Split(Cfg.IntCols, ",")
        If Len(Col) > 0 Then ws.Cells(RowNumber, CStr(Col)).NumberFormat = "0"
    Next Col

    For Each Col In Split(Cfg.MoneyCols, ",")

        If Len(Col) > 0 Then

            ws.Cells(RowNumber, CStr(Col)).NumberFormat = "#,##0.00"

            With ws.Cells(RowNumber, CStr(Col)).Validation
                .Delete
                .Add _
                    Type:=xlValidateDecimal, _
                    AlertStyle:=xlValidAlertStop, _
                    Operator:=xlGreaterEqual, _
                    Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowError = True
                .ErrorTitle = "Некорректная сумма"
                .ErrorMessage = "Введите число, равное или больше нуля."
            End With

        End If

    Next Col

End Sub

Private Function EnsureMultiAttrSheet(ByRef Cfg As TMultiAttr) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Cfg.SheetName)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))

        ws.Name = Cfg.SheetName
        BuildMultiAttrSheet ws, Cfg
        ws.Visible = xlSheetHidden

    End If

    Set EnsureMultiAttrSheet = ws

End Function

Private Sub BuildMultiAttrSheet(ByVal ws As Worksheet, ByRef Cfg As TMultiAttr)

    Dim i As Long
    Dim PrevEvents As Boolean

    PrevEvents = Application.EnableEvents
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    With ws

        .Cells.Clear

        .Range("A1:" & Cfg.LastCol & "1").Merge
        .Range("A1").Value2 = Cfg.TitleText

        .Range("A2").Value2 = "Добавить запись"
        .Range("B2").Value2 = "Вернуться на форму"

        For i = 0 To UBound(Cfg.Headers)
            .Cells(HEADER_ROW, i + 1).Value2 = Cfg.Headers(i)
        Next i

        If Not IsEmpty(Cfg.DictCodes) Then
            For i = 0 To UBound(Cfg.DictCodes)
                .Cells(DICT_CODE_ROW_CHILD, i + 1).Value2 = Cfg.DictCodes(i)
            Next i
            .Rows(DICT_CODE_ROW_CHILD).Hidden = True
        End If

        With .Range("A1:" & Cfg.LastCol & "1")
            .Interior.Color = RGB(31, 78, 121)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 14
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With

        With .Range("A2:B2")
            .Interior.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With

        With .Range("A" & HEADER_ROW & ":" & Cfg.LastCol & HEADER_ROW)
            .Interior.Color = RGB(112, 173, 71)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Borders.LineStyle = xlContinuous
        End With

        .Rows(1).RowHeight = 26
        .Rows(HEADER_ROW).RowHeight = 40

        For i = 0 To UBound(Cfg.Widths)
            .Columns(i + 1).ColumnWidth = Cfg.Widths(i)
        Next i

        .Range("A" & HEADER_ROW & ":" & Cfg.LastCol & (HEADER_ROW + 1)).AutoFilter

        .Columns(Cfg.CtxCol).Hidden = True

    End With

    Application.DisplayAlerts = True
    Application.EnableEvents = PrevEvents

End Sub

Private Sub RestoreMultiAttrContext(ByRef Cfg As TMultiAttr, ByVal ws As Worksheet)
    CurrentProjectID = Trim$(CStr(ws.Range(Cfg.CtxCol & "1").Value2))
    CurrentSourceRow = Val(ws.Range(Cfg.CtxCol & "2").Value2)
End Sub

Private Sub ShowAllMultiAttr(ByVal ws As Worksheet)
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    On Error GoTo 0
End Sub

Private Function LastMultiAttrRow(ByVal ws As Worksheet) As Long
    LastMultiAttrRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If LastMultiAttrRow < HEADER_ROW Then LastMultiAttrRow = HEADER_ROW
End Function

Private Function FirstVisibleMultiAttrRow(ByVal ws As Worksheet, ByVal LastRow As Long) As Long

    Dim RowNumber As Long

    For RowNumber = FIRST_DATA_ROW To LastRow
        If Not ws.Rows(RowNumber).Hidden Then
            If Trim$(CStr(ws.Cells(RowNumber, "A").Value2)) = CurrentProjectID Then
                FirstVisibleMultiAttrRow = RowNumber
                Exit Function
            End If
        End If
    Next RowNumber

End Function
