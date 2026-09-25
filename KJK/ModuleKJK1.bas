Attribute VB_Name = "ModuleKJK1"
Option Explicit

'==================================================================
' modChildTables - служебные листы-формы (подтаблицы проекта)
' БЕЗ ИЗМЕНЕНИЙ в этой итерации (06) - перенесён как есть.
'==================================================================

Private Const CFG_COUNT As Long = 2
Private Const HEADER_ROW As Long = 4
Private Const FIRST_DATA_ROW As Long = 5

Private Type TChild
    SheetName As String
    TriggerCodes As String
    TitleText As String
    LastCol As String
    CtxCol As String
    ReturnCol As String
    Groups As Variant
    Headers As Variant
    Widths As Variant
    IntCols As String
    MoneyCols As String
End Type

Private Sub LoadCfg(ByVal Index As Long, ByRef Cfg As TChild)

    Select Case Index

        Case 1
            Cfg.SheetName = "AKKEMD_024"
            Cfg.TriggerCodes = "AKKEMD_026;AKKEMD_027;AKKEMD_028;" & _
                               "AKKEMD_030;AKKEMD_031;AKKEMD_032"
            Cfg.TitleText = "AKKEMD_024 — Мощность проекта"
            Cfg.LastCol = "H"
            Cfg.CtxCol = "J"
            Cfg.ReturnCol = "U"
            Cfg.Groups = Array( _
                "A3:B3|Связь с проектом", _
                "C3:E3|AKKEMD_025 — Планируемая мощность", _
                "F3:H3|AKKEMD_029 — Фактическая мощность")
            Cfg.Headers = Array( _
                "AKKEMD_003 — ID проекта", "№ записи", _
                "AKKEMD_026 — Наименование продукции", _
                "AKKEMD_027 — Количество", _
                "AKKEMD_028 — Единица измерения", _
                "AKKEMD_030 — Наименование продукции", _
                "AKKEMD_031 — Количество", _
                "AKKEMD_032 — Единица измерения")
            Cfg.Widths = Array(20, 12, 28, 15, 22, 28, 15, 22)
            Cfg.IntCols = "D,G"
            Cfg.MoneyCols = ""

        Case 2
            Cfg.SheetName = "AKKEMD_035"
            Cfg.TriggerCodes = "AKKEMD_036"
            Cfg.TitleText = "AKKEMD_035 — Планируемый объем финансирования"
            Cfg.LastCol = "D"
            Cfg.CtxCol = "F"
            Cfg.ReturnCol = "W"
            Cfg.Groups = Array( _
                "A3:B3|Связь с проектом", _
                "C3:D3|AKKEMD_035 — Сумма в валюте финансирования")
            Cfg.Headers = Array( _
                "AKKEMD_003 — ID проекта", "№ записи", _
                "AKKEMD_036 — Сумма", _
                "AKKEMD_037 — Валюта (SPR017)")
            Cfg.Widths = Array(20, 12, 22, 25)
            Cfg.IntCols = ""
            Cfg.MoneyCols = "C"

    End Select

End Sub

Public Function Child_OpenFromMain( _
    ByVal AttrCode As String, _
    ByVal Target As Range) As Boolean

    Dim i As Long
    Dim Cfg As TChild

    If Len(AttrCode) = 0 Then Exit Function

    For i = 1 To CFG_COUNT

        LoadCfg i, Cfg

        If InStr(1, ";" & Cfg.TriggerCodes & ";", ";" & AttrCode & ";", _
                 vbTextCompare) > 0 Then
            Child_OpenFromMain = True
            OpenChild Cfg, Target
            Exit Function
        End If

    Next i

End Function

Public Function Child_HandleClick( _
    ByVal Sh As Worksheet, _
    ByVal Target As Range) As Boolean

    Dim i As Long
    Dim Cfg As TChild

    For i = 1 To CFG_COUNT

        LoadCfg i, Cfg

        If StrComp(Sh.Name, Cfg.SheetName, vbTextCompare) = 0 Then

            Select Case Target.Address(False, False)

                Case "A2"
                    AddRecord Cfg
                    Child_HandleClick = True

                Case "B2"
                    ReturnToMain Cfg
                    Child_HandleClick = True

            End Select

            Exit Function

        End If

    Next i

End Function

Private Sub OpenChild(ByRef Cfg As TChild, ByVal Target As Range)

    Dim wsSource As Worksheet
    Dim wsChild As Worksheet
    Dim ProjectID As String
    Dim LastRow As Long
    Dim FirstVisibleRow As Long

    If Target.Row < FIRST_INPUT_ROW Then Exit Sub

    Set wsSource = ThisWorkbook.Worksheets(MAIN_SHEET_NAME)

    ProjectID = Trim$(CStr(wsSource.Cells(Target.Row, PROJECT_ID_COL).Value2))

    If ProjectID = vbNullString Then
        MsgBox "Сначала заполните AKKEMD_003 — ID проекта в столбце " & _
               PROJECT_ID_COL & ".", vbExclamation
        Exit Sub
    End If

    Set wsChild = EnsureSheet(Cfg)

    wsChild.Range(Cfg.CtxCol & "1").Value2 = ProjectID
    wsChild.Range(Cfg.CtxCol & "2").Value2 = Target.Row
    wsChild.Range("A1").Value2 = Cfg.TitleText & ": " & ProjectID

    ShowAll wsChild

    EnsureFirstRecord Cfg, wsChild, ProjectID
    LastRow = LastDataRow(wsChild)

    wsChild.Range("A" & HEADER_ROW & ":" & Cfg.LastCol & LastRow).AutoFilter _
        Field:=1, Criteria1:=ProjectID

    FirstVisibleRow = FirstVisibleRowOf(wsChild, LastRow, ProjectID)

    wsChild.Activate

    If FirstVisibleRow = 0 Then
        wsChild.Range("A2").Select
    Else
        wsChild.Cells(FirstVisibleRow, "C").Select
    End If

End Sub

Private Sub EnsureFirstRecord( _
    ByRef Cfg As TChild, _
    ByVal ws As Worksheet, _
    ByVal ProjectID As String)

    Dim LastRow As Long
    Dim NewRow As Long
    Dim ExistingCount As Long

    LastRow = LastDataRow(ws)

    If LastRow >= FIRST_DATA_ROW Then
        ExistingCount = Application.WorksheetFunction.CountIf( _
            ws.Range("A" & FIRST_DATA_ROW & ":A" & LastRow), ProjectID)
    End If

    If ExistingCount > 0 Then Exit Sub

    NewRow = LastRow + 1

    ws.Cells(NewRow, "A").Value2 = ProjectID
    ws.Cells(NewRow, "B").Value2 = 1

    FormatRow Cfg, ws, NewRow

End Sub

Private Sub AddRecord(ByRef Cfg As TChild)

    Dim ws As Worksheet
    Dim ProjectID As String
    Dim LastRow As Long
    Dim NewRow As Long
    Dim RecordNumber As Long

    Set ws = EnsureSheet(Cfg)

    ProjectID = Trim$(CStr(ws.Range(Cfg.CtxCol & "1").Value2))

    If ProjectID = vbNullString Then
        MsgBox "Не определен ID проекта.", vbExclamation
        Exit Sub
    End If

    ShowAll ws

    LastRow = LastDataRow(ws)
    NewRow = LastRow + 1
    RecordNumber = 1

    If LastRow >= FIRST_DATA_ROW Then
        RecordNumber = Application.WorksheetFunction.CountIf( _
            ws.Range("A" & FIRST_DATA_ROW & ":A" & LastRow), ProjectID) + 1
    End If

    ws.Cells(NewRow, "A").Value2 = ProjectID
    ws.Cells(NewRow, "B").Value2 = RecordNumber

    FormatRow Cfg, ws, NewRow

    ws.Range("A" & HEADER_ROW & ":" & Cfg.LastCol & NewRow).AutoFilter _
        Field:=1, Criteria1:=ProjectID

    ws.Cells(NewRow, "C").Select

End Sub

Private Sub ReturnToMain(ByRef Cfg As TChild)

    Dim wsChild As Worksheet
    Dim wsSource As Worksheet
    Dim SourceRow As Long

    Set wsChild = EnsureSheet(Cfg)
    SourceRow = Val(wsChild.Range(Cfg.CtxCol & "2").Value2)

    ShowAll wsChild

    Set wsSource = ThisWorkbook.Worksheets(MAIN_SHEET_NAME)
    wsSource.Activate

    If SourceRow >= FIRST_INPUT_ROW Then
        wsSource.Cells(SourceRow, Cfg.ReturnCol).Select
    Else
        wsSource.Range(Cfg.ReturnCol & FIRST_INPUT_ROW).Select
    End If

End Sub

Private Sub FormatRow( _
    ByRef Cfg As TChild, _
    ByVal ws As Worksheet, _
    ByVal RowNumber As Long)

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

Private Function EnsureSheet(ByRef Cfg As TChild) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Cfg.SheetName)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))

        ws.Name = Cfg.SheetName
        BuildSheet ws, Cfg

    End If

    Set EnsureSheet = ws

End Function

Private Sub BuildSheet(ByVal ws As Worksheet, ByRef Cfg As TChild)

    Dim i As Long
    Dim g As Variant
    Dim Parts As Variant

    Application.DisplayAlerts = False

    With ws

        .Cells.Clear

        .Range("A1:" & Cfg.LastCol & "1").Merge
        .Range("A1").Value2 = Cfg.TitleText

        .Range("A2").Value2 = "Добавить запись"
        .Range("B2").Value2 = "Вернуться на форму"

        For Each g In Cfg.Groups
            Parts = Split(CStr(g), "|")
            .Range(CStr(Parts(0))).Merge
            .Range(CStr(Parts(0))).Cells(1, 1).Value2 = Parts(1)
        Next g

        For i = 0 To UBound(Cfg.Headers)
            .Cells(HEADER_ROW, i + 1).Value2 = Cfg.Headers(i)
        Next i

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

        With .Range("A3:" & Cfg.LastCol & "3")
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With

        With .Range("A4:" & Cfg.LastCol & "4")
            .Interior.Color = RGB(112, 173, 71)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Borders.LineStyle = xlContinuous
        End With

        .Rows(1).RowHeight = 26
        .Rows(3).RowHeight = 24
        .Rows(4).RowHeight = 55

        For i = 0 To UBound(Cfg.Widths)
            .Columns(i + 1).ColumnWidth = Cfg.Widths(i)
        Next i

        .Range("A4:" & Cfg.LastCol & "5").AutoFilter

        .Columns(Cfg.CtxCol).Hidden = True

    End With

    Application.DisplayAlerts = True

End Sub

Private Sub ShowAll(ByVal ws As Worksheet)
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    On Error GoTo 0
End Sub

Private Function LastDataRow(ByVal ws As Worksheet) As Long
    LastDataRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If LastDataRow < HEADER_ROW Then LastDataRow = HEADER_ROW
End Function

Private Function FirstVisibleRowOf( _
    ByVal ws As Worksheet, _
    ByVal LastRow As Long, _
    ByVal ProjectID As String) As Long

    Dim RowNumber As Long

    For RowNumber = FIRST_DATA_ROW To LastRow
        If Not ws.Rows(RowNumber).Hidden Then
            If Trim$(CStr(ws.Cells(RowNumber, "A").Value2)) = ProjectID Then
                FirstVisibleRowOf = RowNumber
                Exit Function
            End If
        End If
    Next RowNumber

End Function
