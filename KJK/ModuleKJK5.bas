Attribute VB_Name = "ModuleKJK5"
Option Explicit

'==================================================================
' modNestedGroups - множественные атрибуты со сводной ячейкой.
' Две настоящие вложенности ("лесенка"):
'   KJKEMD_032 -> KJKEMD_176  (Учредитель -> Учредитель учредителя)
'   KJKEMD_087 -> KJKEMD_112  (Договор -> Реструктуризация)
' Остальные (036, 044, 127, 050) - плоские группы верхнего уровня,
' без вложенности (проверено по листу "КЖК" - все их дочерние поля
' помечены "Простой", вложенных МНОЖЕСТВЕННЫЙ внутри них нет).
'
' Логика та же, что и раньше для 032:
'   - на главном листе одна ячейка-сводка на группу, код группы
'     прописан в строке MULTI_ATTR_CODE_ROW (7)
'   - двойной клик по ней открывает лист группы
'   - если у группы есть вложенная - внутри ее листа есть своя
'     сводная ячейка (ChildSummaryCol), открывающая лист вложенной
'     группы, отфильтрованный по конкретной родительской строке
'     (через уникальный ключ в KeyCol)
'   - при возврате в сводной ячейке родителя проставляется
'     "Заполнен N атрибутами"
'==================================================================

Private Const GROUP_COUNT As Long = 8
Private Const HEADER_ROW As Long = 4
Private Const FIRST_DATA_ROW As Long = 5
Private Const DICT_CODE_ROW_CHILD As Long = 3

Private Type TGroup
    GroupCode As String
    ParentSheetName As String
    ParentCol As String
    ParentKeyCol As String
    SheetName As String
    TitleText As String
    LastCol As String
    ValueCols As String
    KeyCol As String
    CtxCol As String
    Headers As Variant
    Widths As Variant
    IntCols As String
    MoneyCols As String
    DictCodes As Variant
    ChildGroupCode As String
    ChildSummaryCol As String
End Type

Private CurrentKey As String
Private CurrentProjectID As String
Private CurrentSourceRow As Long
Private CurrentParentSheetName As String

'------------------------------------------------------------------
' НАСТРОЙКА ГРУПП
'------------------------------------------------------------------
Private Sub LoadGroupCfg(ByVal Index As Long, ByRef Cfg As TGroup)

    Select Case Index

        Case 1      ' KJKEMD_032 - Учредитель заемщика (верхний уровень)
            Cfg.GroupCode = "KJKEMD_032"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "AE"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "KJKEMD_032"
            Cfg.TitleText = "KJKEMD_032 — Учредитель заемщика"
            Cfg.LastCol = "F"
            Cfg.ValueCols = "B,C,D,E"
            Cfg.KeyCol = "G"
            Cfg.CtxCol = "H"
            Cfg.Headers = Array( _
                "KJKEMD_003 — ID проекта", _
                "KJKEMD_175 — Правовой тип лица (SPR053)", _
                "KJKEMD_033 — ФИО учредителя", _
                "KJKEMD_034 — Страна (SPR039)", _
                "KJKEMD_035 — Доля, %", _
                "KJKEMD_176 — Учредитель учредителя (список)")
            Cfg.Widths = Array(20, 22, 32, 22, 14, 26)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", "SPR053", "", "SPR039", "", "")
            Cfg.ChildGroupCode = "KJKEMD_176"
            Cfg.ChildSummaryCol = "F"

        Case 2      ' KJKEMD_176 - Учредитель учредителя (вложена в 032)
            Cfg.GroupCode = "KJKEMD_176"
            Cfg.ParentSheetName = "KJKEMD_032"
            Cfg.ParentCol = "F"
            Cfg.ParentKeyCol = "G"
            Cfg.SheetName = "KJKEMD_176"
            Cfg.TitleText = "KJKEMD_176 — Учредитель учредителя заемщика"
            Cfg.LastCol = "D"
            Cfg.ValueCols = "B,C,D"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "E"
            Cfg.Headers = Array( _
                "Ключ учредителя (служебное)", _
                "KJKEMD_177 — ФИО", _
                "KJKEMD_178 — Страна (SPR039)", _
                "KJKEMD_179 — Доля, %")
            Cfg.Widths = Array(20, 32, 22, 14)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", "", "SPR039", "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 3      ' KJKEMD_036 - Иностранные инвесторы, ТНК (плоская)
            Cfg.GroupCode = "KJKEMD_036"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "AF"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "KJKEMD_036"
            Cfg.TitleText = "KJKEMD_036 — Наименование иностранных инвесторов, ТНК"
            Cfg.LastCol = "C"
            Cfg.ValueCols = "B,C"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "D"
            Cfg.Headers = Array( _
                "KJKEMD_003 — ID проекта", _
                "KJKEMD_037 — Наименование", _
                "KJKEMD_038 — Страна (SPR039)")
            Cfg.Widths = Array(20, 32, 22)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", "", "SPR039")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 4      ' KJKEMD_044 - Валютное финансирование (плоская)
            Cfg.GroupCode = "KJKEMD_044"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "AJ"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "KJKEMD_044"
            Cfg.TitleText = "KJKEMD_044 — Планируемый объем инвестиций (валютное финансирование)"
            Cfg.LastCol = "C"
            Cfg.ValueCols = "B,C"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "D"
            Cfg.Headers = Array( _
                "KJKEMD_003 — ID проекта", _
                "KJKEMD_045 — Сумма", _
                "KJKEMD_046 — Валюта (SPR017)")
            Cfg.Widths = Array(20, 18, 18)
            Cfg.IntCols = ""
            Cfg.MoneyCols = "B"
            Cfg.DictCodes = Array("", "", "SPR017")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 5      ' KJKEMD_127 - Освоение средств по годам (плоская)
            Cfg.GroupCode = "KJKEMD_127"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "BI"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "KJKEMD_127"
            Cfg.TitleText = "KJKEMD_127 — Освоение средств по годам"
            Cfg.LastCol = "D"
            Cfg.ValueCols = "B,C,D"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "E"
            Cfg.Headers = Array( _
                "KJKEMD_003 — ID проекта", _
                "KJKEMD_128 — Год (ГГГГ)", _
                "KJKEMD_129 — План", _
                "KJKEMD_130 — Факт")
            Cfg.Widths = Array(20, 12, 16, 16)
            Cfg.IntCols = "B"
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", "", "", "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 6      ' KJKEMD_050 - Источник фондирования Холдинга (плоская)
            Cfg.GroupCode = "KJKEMD_050"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "AM"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "KJKEMD_050"
            Cfg.TitleText = "KJKEMD_050 — Источник фондирования Холдинга"
            Cfg.LastCol = "R"
            Cfg.ValueCols = "B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "S"
            Cfg.Headers = Array( _
                "KJKEMD_003 — ID проекта", _
                "KJKEMD_051 — Источник средств (SPR015)", _
                "KJKEMD_052 — Сумма по источнику средств", _
                "KJKEMD_053 — Способ привлечения фондирования (SPR011)", _
                "KJKEMD_055 — Наименование ФЭО", _
                "KJKEMD_056 — № бюджетной программы", _
                "KJKEMD_057 — Дата Приказа ГО", _
                "KJKEMD_058 — № Приказа ГО", _
                "KJKEMD_059 — Наименование ЦГО (SPR046)", _
                "KJKEMD_061 — Статус участия в гос. программах (SPR047)", _
                "KJKEMD_062 — Полное наименование программы", _
                "KJKEMD_063 — Наименование НПА", _
                "KJKEMD_064 — Дата", _
                "KJKEMD_065 — №", _
                "KJKEMD_066 — Тип ставки привлечения (SPR029)", _
                "KJKEMD_067 — Ставка привлечения, % (фиксированная)", _
                "KJKEMD_069 — Индекс", _
                "KJKEMD_070 — Спред")
            Cfg.Widths = Array(20, 24, 18, 26, 22, 18, 16, 16, 22, 26, 26, 26, 14, 10, 22, 22, 14, 14)
            Cfg.IntCols = "N"
            Cfg.MoneyCols = "C"
            Cfg.DictCodes = Array("", "SPR015", "", "SPR011", "", "", "", "", _
                                   "SPR046", "SPR047", "", "", "", "", "SPR029", "", "", "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 7      ' KJKEMD_087 - Договор (верхний уровень, вложена 112)
            Cfg.GroupCode = "KJKEMD_087"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "BF"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "KJKEMD_087"
            Cfg.TitleText = "KJKEMD_087 — Договор"
            Cfg.LastCol = "S"
            Cfg.ValueCols = "B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R"
            Cfg.KeyCol = "T"
            Cfg.CtxCol = "U"
            Cfg.Headers = Array( _
                "KJKEMD_003 — ID проекта", _
                "KJKEMD_094 — Номер СОКЛ/договора/банк.гарантии", _
                "KJKEMD_095 — Дата СОКЛ/договора/банк.гарантии", _
                "KJKEMD_098 — Дата ввода информации по задолженности", _
                "KJKEMD_100 — Остаток: по основному долгу, тенге", _
                "KJKEMD_101 — Остаток: по вознаграждению, тенге", _
                "KJKEMD_103 — Просрочка: по основному долгу, тенге", _
                "KJKEMD_104 — Просрочка: по вознаграждению, тенге", _
                "KJKEMD_105 — Дисконт/премия, корректировка стоимости займа", _
                "KJKEMD_106 — Пеня, штрафы, тенге", _
                "KJKEMD_107 — Количество дней просрочки", _
                "KJKEMD_108 — Стадия (МСФО 9)", _
                "KJKEMD_109 — Провизии, %", _
                "KJKEMD_110 — Категория учета (SPR031)", _
                "KJKEMD_111 — Провизии, тенге", _
                "KJKEMD_119 — Дата заключения договора займа", _
                "KJKEMD_120 — Дата окончания срока действия договора", _
                "KJKEMD_121 — Дни (окончание минус заключение)", _
                "KJKEMD_112 — Реструктуризация (список)")
            Cfg.Widths = Array(20, 22, 16, 20, 18, 18, 18, 18, 22, 16, 14, 16, 12, 22, 16, 18, 18, 14, 24)
            Cfg.IntCols = "K,R"
            Cfg.MoneyCols = "E,F,G,H,J,O"
            Cfg.DictCodes = Array("", "", "", "", "", "", "", "", "", "", "", "", "", _
                                   "SPR031", "", "", "", "", "")
            Cfg.ChildGroupCode = "KJKEMD_112"
            Cfg.ChildSummaryCol = "S"

        Case 8      ' KJKEMD_112 - Реструктуризация (вложена в 087)
            Cfg.GroupCode = "KJKEMD_112"
            Cfg.ParentSheetName = "KJKEMD_087"
            Cfg.ParentCol = "S"
            Cfg.ParentKeyCol = "T"
            Cfg.SheetName = "KJKEMD_112"
            Cfg.TitleText = "KJKEMD_112 — Реструктуризация"
            Cfg.LastCol = "F"
            Cfg.ValueCols = "B,C,D,E,F"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "G"
            Cfg.Headers = Array( _
                "Ключ договора (служебное)", _
                "KJKEMD_113 — Дата реструктуризации", _
                "KJKEMD_114 — Основные условия по реструктуризации", _
                "KJKEMD_115 — Решения УО/ДЗО/поручения по проблемным вопросам", _
                "KJKEMD_116 — Дальнейший план мероприятий по оздоровлению", _
                "KJKEMD_117 — Документ дальнейшего плана мероприятий")
            Cfg.Widths = Array(20, 18, 30, 34, 34, 30)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", "", "", "", "", "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

    End Select

End Sub

'------------------------------------------------------------------
' Маршрутизация (вызывается из modRouter)
'------------------------------------------------------------------
Public Function IsGroupSheet(ByVal SheetName As String) As Boolean

    Dim i As Long
    Dim Cfg As TGroup

    For i = 1 To GROUP_COUNT
        LoadGroupCfg i, Cfg
        If StrComp(Cfg.SheetName, SheetName, vbTextCompare) = 0 Then
            IsGroupSheet = True
            Exit Function
        End If
    Next i

End Function

Public Function Group_HandleClick(ByVal Sh As Worksheet, ByVal Target As Range) As Boolean

    Dim i As Long
    Dim Cfg As TGroup

    For i = 1 To GROUP_COUNT

        LoadGroupCfg i, Cfg

        If StrComp(Sh.Name, Cfg.SheetName, vbTextCompare) = 0 Then

            Select Case Target.Address(False, False)

                Case "A2"
                    AddGroupRecord Cfg
                    Group_HandleClick = True

                Case "B2"
                    ReturnFromGroup Cfg
                    Group_HandleClick = True

            End Select

            Exit Function

        End If

    Next i

End Function

Public Function Group_OpenFromMain(ByVal ws As Worksheet, ByVal Target As Range) As Boolean

    Dim Code As String
    Dim i As Long
    Dim Cfg As TGroup

    If ws.Name <> MAIN_SHEET_NAME Then Exit Function
    If Target.Row < FIRST_INPUT_ROW Then Exit Function

    Code = Trim$(CStr(ws.Cells(MULTI_ATTR_CODE_ROW, Target.Column).Value2))
    If Len(Code) = 0 Then Exit Function

    For i = 1 To GROUP_COUNT

        LoadGroupCfg i, Cfg

        If Len(Cfg.ParentSheetName) = 0 And StrComp(Cfg.GroupCode, Code, vbTextCompare) = 0 Then
            Group_OpenFromMain = True
            OpenGroup Cfg, ws, Target
            Exit Function
        End If

    Next i

End Function

Public Function Group_OpenFromParent(ByVal ws As Worksheet, ByVal Target As Range) As Boolean

    Dim i As Long
    Dim Cfg As TGroup

    If Target.Row < FIRST_DATA_ROW Then Exit Function

    For i = 1 To GROUP_COUNT

        LoadGroupCfg i, Cfg

        If Len(Cfg.ParentSheetName) > 0 Then
            If StrComp(ws.Name, Cfg.ParentSheetName, vbTextCompare) = 0 Then
                If Target.Column = ws.Range(Cfg.ParentCol & "1").Column Then
                    Group_OpenFromParent = True
                    OpenGroup Cfg, ws, Target
                    Exit Function
                End If
            End If
        End If

    Next i

End Function

'------------------------------------------------------------------
' Открыть группу (единая логика и для верхнего уровня, и для вложенной)
'------------------------------------------------------------------
Private Sub OpenGroup(ByRef Cfg As TGroup, ByVal wsSource As Worksheet, ByVal Target As Range)

    Dim wsChild As Worksheet
    Dim KeyValue As String
    Dim LastRow As Long
    Dim FirstVisibleRow As Long

    If Len(Cfg.ParentSheetName) = 0 Then

        CurrentProjectID = Trim$(CStr(wsSource.Cells(Target.Row, PROJECT_ID_COL).Value2))

        If CurrentProjectID = vbNullString Then
            MsgBox "Сначала заполните KJKEMD_003 — ID проекта в столбце " & _
                   PROJECT_ID_COL & ".", vbExclamation
            Exit Sub
        End If

        KeyValue = CurrentProjectID

    Else

        CurrentProjectID = Trim$(CStr(wsSource.Cells(Target.Row, "A").Value2))

        If CurrentProjectID = vbNullString Then
            MsgBox "Сначала заполните ID проекта в этой строке.", vbExclamation
            Exit Sub
        End If

        KeyValue = EnsureRecordKey(wsSource, Target.Row, Cfg.ParentKeyCol)

    End If

    CurrentKey = KeyValue
    CurrentSourceRow = Target.Row
    CurrentParentSheetName = wsSource.Name

    Set wsChild = EnsureGroupSheet(Cfg)

    wsChild.Range(Cfg.CtxCol & "1").Value2 = CurrentKey
    wsChild.Range(Cfg.CtxCol & "2").Value2 = CurrentSourceRow
    wsChild.Range(Cfg.CtxCol & "3").Value2 = CurrentProjectID
    wsChild.Range(Cfg.CtxCol & "4").Value2 = CurrentParentSheetName
    wsChild.Range("A1").Value2 = Cfg.TitleText & ": " & CurrentProjectID

    ShowAllGroup wsChild

    EnsureFirstGroupRecord Cfg, wsChild, CurrentKey

    LastRow = LastGroupRow(wsChild)

    wsChild.Range("A" & HEADER_ROW & ":" & Cfg.LastCol & LastRow).AutoFilter _
        Field:=1, Criteria1:=CurrentKey

    FirstVisibleRow = FirstVisibleGroupRow(wsChild, LastRow, CurrentKey)

    wsChild.Visible = xlSheetVisible
    wsChild.Activate

    If FirstVisibleRow = 0 Then
        wsChild.Range("A2").Select
    Else
        wsChild.Cells(FirstVisibleRow, "B").Select
    End If

End Sub

Private Function EnsureRecordKey( _
    ByVal ws As Worksheet, ByVal RowNum As Long, ByVal KeyColLetter As String) As String

    Dim ExistingKey As String
    Dim Seq As Long

    ExistingKey = Trim$(CStr(ws.Cells(RowNum, KeyColLetter).Value2))

    If Len(ExistingKey) > 0 Then
        EnsureRecordKey = ExistingKey
        Exit Function
    End If

    Seq = Val(ws.Range("BZ1").Value2) + 1
    ws.Range("BZ1").Value2 = Seq

    ExistingKey = "K" & Format(Seq, "000000")

    Application.EnableEvents = False
    ws.Cells(RowNum, KeyColLetter).Value2 = ExistingKey
    Application.EnableEvents = True

    EnsureRecordKey = ExistingKey

End Function

Private Sub EnsureFirstGroupRecord(ByRef Cfg As TGroup, ByVal ws As Worksheet, ByVal KeyValue As String)

    Dim LastRow As Long
    Dim NewRow As Long
    Dim ExistingCount As Long

    LastRow = LastGroupRow(ws)

    If LastRow >= FIRST_DATA_ROW Then
        ExistingCount = Application.WorksheetFunction.CountIf( _
            ws.Range("A" & FIRST_DATA_ROW & ":A" & LastRow), KeyValue)
    End If

    If ExistingCount > 0 Then Exit Sub

    NewRow = LastRow + 1
    ws.Cells(NewRow, "A").Value2 = KeyValue

    FormatGroupRow Cfg, ws, NewRow

End Sub

Private Sub AddGroupRecord(ByRef Cfg As TGroup)

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim NewRow As Long

    Set ws = EnsureGroupSheet(Cfg)
    RestoreGroupContext Cfg, ws

    If CurrentKey = vbNullString Then
        MsgBox "Не определен ключ записи.", vbExclamation
        Exit Sub
    End If

    ShowAllGroup ws

    LastRow = LastGroupRow(ws)
    NewRow = LastRow + 1

    ws.Cells(NewRow, "A").Value2 = CurrentKey

    FormatGroupRow Cfg, ws, NewRow

    ws.Range("A" & HEADER_ROW & ":" & Cfg.LastCol & NewRow).AutoFilter _
        Field:=1, Criteria1:=CurrentKey

    ws.Cells(NewRow, "B").Select

End Sub

Private Sub ReturnFromGroup(ByRef Cfg As TGroup)

    Dim wsChild As Worksheet
    Dim wsParent As Worksheet

    Set wsChild = EnsureGroupSheet(Cfg)
    RestoreGroupContext Cfg, wsChild

    ShowAllGroup wsChild

    If Not ValidateGroupRows(Cfg, wsChild) Then Exit Sub

    wsChild.Visible = xlSheetHidden

    RefreshSummaryForKey Cfg, CurrentKey

    On Error Resume Next
    Set wsParent = ThisWorkbook.Worksheets(CurrentParentSheetName)
    On Error GoTo 0
    If wsParent Is Nothing Then Set wsParent = ThisWorkbook.Worksheets(MAIN_SHEET_NAME)

    wsParent.Activate
    wsParent.Range(Cfg.ParentCol & CurrentSourceRow).Select

End Sub

Private Sub RestoreGroupContext(ByRef Cfg As TGroup, ByVal ws As Worksheet)
    CurrentKey = Trim$(CStr(ws.Range(Cfg.CtxCol & "1").Value2))
    CurrentSourceRow = Val(ws.Range(Cfg.CtxCol & "2").Value2)
    CurrentProjectID = Trim$(CStr(ws.Range(Cfg.CtxCol & "3").Value2))
    CurrentParentSheetName = Trim$(CStr(ws.Range(Cfg.CtxCol & "4").Value2))
End Sub

Private Function ValidateGroupRows(ByRef Cfg As TGroup, ByVal ws As Worksheet) As Boolean

    Dim LastRow As Long
    Dim r As Long
    Dim ColsArr As Variant
    Dim i As Long
    Dim ColLetter As String
    Dim FilledCount As Long
    Dim TotalCount As Long
    Dim FirstEmptyCol As String

    ValidateGroupRows = True

    ColsArr = Split(Cfg.ValueCols, ",")
    LastRow = LastGroupRow(ws)

    For r = FIRST_DATA_ROW To LastRow

        FilledCount = 0
        TotalCount = 0
        FirstEmptyCol = vbNullString

        For i = LBound(ColsArr) To UBound(ColsArr)

            ColLetter = Trim$(CStr(ColsArr(i)))
            If Len(ColLetter) = 0 Then GoTo NextCol

            TotalCount = TotalCount + 1

            If Len(Trim$(CStr(ws.Range(ColLetter & r).Value2))) > 0 Then
                FilledCount = FilledCount + 1
            ElseIf Len(FirstEmptyCol) = 0 Then
                FirstEmptyCol = ColLetter
            End If

NextCol:
        Next i

        If FilledCount > 0 And FilledCount < TotalCount Then

            MsgBox "Строка " & r & " на листе """ & ws.Name & """ заполнена не полностью." & _
                   vbCrLf & "Заполните все поля записи, либо очистите строку целиком.", _
                   vbExclamation, "Возврат отменен"

            ws.Range(FirstEmptyCol & r).Select

            ValidateGroupRows = False
            Exit Function

        End If

    Next r

End Function

Private Sub RefreshSummaryForKey(ByRef Cfg As TGroup, ByVal KeyValue As String)

    Dim wsParent As Worksheet
    Dim FoundCell As Range
    Dim Count As Long

    If Len(Trim$(KeyValue)) = 0 Then Exit Sub

    Count = CountGroupRecords(Cfg.SheetName, KeyValue)

    If Len(Cfg.ParentSheetName) = 0 Then

        Set wsParent = ThisWorkbook.Worksheets(MAIN_SHEET_NAME)
        Set FoundCell = wsParent.Columns(PROJECT_ID_COL).Find( _
            What:=KeyValue, LookAt:=xlWhole, LookIn:=xlValues)

    Else

        On Error Resume Next
        Set wsParent = ThisWorkbook.Worksheets(Cfg.ParentSheetName)
        On Error GoTo 0
        If wsParent Is Nothing Then Exit Sub

        Set FoundCell = wsParent.Columns(Cfg.ParentKeyCol).Find( _
            What:=KeyValue, LookAt:=xlWhole, LookIn:=xlValues)

    End If

    If FoundCell Is Nothing Then Exit Sub

    Application.EnableEvents = False
    wsParent.Range(Cfg.ParentCol & FoundCell.Row).Value2 = SummaryText(Count)
    Application.EnableEvents = True

End Sub

Private Function SummaryText(ByVal Count As Long) As String
    If Count <= 0 Then
        SummaryText = vbNullString
    Else
        SummaryText = "Заполнен " & Count & " атрибутами"
    End If
End Function

Private Function CountGroupRecords(ByVal SheetName As String, ByVal KeyValue As String) As Long

    Dim ws As Worksheet
    Dim LastRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SheetName)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function
    If Len(Trim$(KeyValue)) = 0 Then Exit Function

    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If LastRow < FIRST_DATA_ROW Then Exit Function

    CountGroupRecords = Application.WorksheetFunction.CountIf( _
        ws.Range("A" & FIRST_DATA_ROW & ":A" & LastRow), KeyValue)

End Function

'------------------------------------------------------------------
' Обработка изменения ячейки
'------------------------------------------------------------------
Public Sub NestedGroup_HandleChange(ByVal ws As Worksheet, ByVal Target As Range)

    Dim i As Long
    Dim Cfg As TGroup

    For i = 1 To GROUP_COUNT

        LoadGroupCfg i, Cfg

        If StrComp(ws.Name, Cfg.SheetName, vbTextCompare) = 0 Then
            HandleGroupSheetChange Cfg, ws, Target
            Exit Sub
        End If

    Next i

End Sub

Private Sub HandleGroupSheetChange(ByRef Cfg As TGroup, ByVal ws As Worksheet, ByVal Target As Range)

    Dim ChangedRows As Range
    Dim OneRow As Range
    Dim RowNumber As Long
    Dim LastRow As Long
    Dim ColsArr As Variant
    Dim i As Long
    Dim AnyFilled As Boolean
    Dim ColLetter As String
    Dim KeyForFilter As String
    Dim OldKeyForChild As String

    ColsArr = Split(Cfg.ValueCols, ",")

    On Error Resume Next
    Set ChangedRows = Intersect(Target, _
        ws.Range(ColsArr(LBound(ColsArr)) & FIRST_DATA_ROW & ":" & _
                 ColsArr(UBound(ColsArr)) & 100000))
    On Error GoTo 0

    If ChangedRows Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler
    Application.EnableEvents = False

    KeyForFilter = Trim$(CStr(ws.Range(Cfg.CtxCol & "1").Value2))

    For Each OneRow In ChangedRows.Rows

        RowNumber = OneRow.Row
        AnyFilled = False

        For i = LBound(ColsArr) To UBound(ColsArr)
            ColLetter = Trim$(CStr(ColsArr(i)))
            If Len(Trim$(CStr(ws.Range(ColLetter & RowNumber).Value2))) > 0 Then
                AnyFilled = True
                Exit For
            End If
        Next i

        If Not AnyFilled Then

            OldKeyForChild = vbNullString
            If Len(Cfg.KeyCol) > 0 Then
                OldKeyForChild = Trim$(CStr(ws.Cells(RowNumber, Cfg.KeyCol).Value2))
            End If

            If Len(OldKeyForChild) > 0 Then
                ClearChildGroupRecords Cfg, OldKeyForChild
            End If

            ws.Cells(RowNumber, "A").ClearContents
            If Len(Cfg.KeyCol) > 0 Then ws.Cells(RowNumber, Cfg.KeyCol).ClearContents
            If Len(Cfg.ChildSummaryCol) > 0 Then ws.Cells(RowNumber, Cfg.ChildSummaryCol).ClearContents

        Else

            If Trim$(CStr(ws.Cells(RowNumber, "A").Value2)) = vbNullString Then
                ws.Cells(RowNumber, "A").NumberFormat = "@"
                ws.Cells(RowNumber, "A").Value2 = KeyForFilter
            End If

        End If

    Next OneRow

    LastRow = LastGroupRow(ws)

    ws.Range("A" & HEADER_ROW & ":" & Cfg.LastCol & LastRow).AutoFilter _
        Field:=1, Criteria1:=KeyForFilter

ExitHandler:
    Application.EnableEvents = True
    RefreshSummaryForKey Cfg, KeyForFilter
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    MsgBox "Ошибка обработки изменения на листе " & ws.Name & ": " & Err.Description, vbExclamation

End Sub

Private Sub ClearChildGroupRecords(ByRef Cfg As TGroup, ByVal ChildKeyValue As String)

    Dim wsChild As Worksheet
    Dim ChildCfg As TGroup
    Dim i As Long
    Dim r As Long
    Dim LastRow As Long

    If Len(Cfg.ChildGroupCode) = 0 Then Exit Sub
    If Len(Trim$(ChildKeyValue)) = 0 Then Exit Sub

    For i = 1 To GROUP_COUNT

        LoadGroupCfg i, ChildCfg

        If StrComp(ChildCfg.GroupCode, Cfg.ChildGroupCode, vbTextCompare) = 0 Then

            On Error Resume Next
            Set wsChild = ThisWorkbook.Worksheets(ChildCfg.SheetName)
            On Error GoTo 0
            If wsChild Is Nothing Then Exit Sub

            LastRow = LastGroupRow(wsChild)

            For r = FIRST_DATA_ROW To LastRow
                If Trim$(CStr(wsChild.Cells(r, "A").Value2)) = Trim$(ChildKeyValue) Then
                    wsChild.Range("A" & r & ":" & ChildCfg.LastCol & r).ClearContents
                End If
            Next r

            Exit For

        End If

    Next i

End Sub

'------------------------------------------------------------------
' Построение листа группы (если его еще нет)
'------------------------------------------------------------------
Private Function EnsureGroupSheet(ByRef Cfg As TGroup) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Cfg.SheetName)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))

        ws.Name = Cfg.SheetName
        BuildGroupSheet ws, Cfg
        ws.Visible = xlSheetHidden

    End If

    Set EnsureGroupSheet = ws

End Function

Private Sub BuildGroupSheet(ByVal ws As Worksheet, ByRef Cfg As TGroup)

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

        If Len(Cfg.KeyCol) > 0 Then .Columns(Cfg.KeyCol).Hidden = True
        .Columns(Cfg.CtxCol).Hidden = True

        If Len(Cfg.ChildSummaryCol) > 0 Then
            With .Range(Cfg.ChildSummaryCol & FIRST_DATA_ROW & ":" & Cfg.ChildSummaryCol & 10000)
                .Interior.Color = RGB(242, 242, 242)
                .Font.Italic = True
            End With
        End If

        .Range("BZ1").Value2 = 0   ' счетчик уникальных ключей для дочерних записей

    End With

    ProtectGroupSheet ws, Cfg

    Application.DisplayAlerts = True
    Application.EnableEvents = PrevEvents

End Sub

Public Sub ProtectGroupSheet(ByVal ws As Worksheet, ByRef Cfg As TGroup)

    Const BIG_ROW As Long = 10000

    On Error Resume Next
    If ws.ProtectContents Then ws.Unprotect
    On Error GoTo 0

    ' Строку 1 (объединенная ячейка заголовка A1:LastCol1) НЕ трогаем -
    ' Excel не разрешает менять Locked для части объединенной ячейки,
    ' а столбец целиком (ws.Columns(...)) как раз ее захватывает.
    ws.Range("A2:AZ" & BIG_ROW).Locked = False

    ws.Range("A2:A" & BIG_ROW).Locked = True
    If Len(Cfg.KeyCol) > 0 Then ws.Range(Cfg.KeyCol & "2:" & Cfg.KeyCol & BIG_ROW).Locked = True
    ws.Range(Cfg.CtxCol & "2:" & Cfg.CtxCol & BIG_ROW).Locked = True
    If Len(Cfg.ChildSummaryCol) > 0 Then ws.Range(Cfg.ChildSummaryCol & "2:" & Cfg.ChildSummaryCol & BIG_ROW).Locked = True

    ws.Range("A2:B2").Locked = False

    On Error Resume Next
    ws.Protect DrawingObjects:=False, _
               Contents:=True, _
               Scenarios:=True, _
               UserInterfaceOnly:=True, _
               AllowFiltering:=True, _
               AllowSorting:=True
    On Error GoTo 0

    ws.EnableSelection = xlNoRestrictions

End Sub

Public Sub ReapplyGroupProtection()

    Dim i As Long
    Dim Cfg As TGroup
    Dim ws As Worksheet

    For i = 1 To GROUP_COUNT

        LoadGroupCfg i, Cfg

        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(Cfg.SheetName)
        On Error GoTo 0

        If Not ws Is Nothing Then ProtectGroupSheet ws, Cfg

    Next i

End Sub

'------------------------------------------------------------------
' Вспомогательные
'------------------------------------------------------------------
Private Sub FormatGroupRow(ByRef Cfg As TGroup, ByVal ws As Worksheet, ByVal RowNumber As Long)

    Dim Col As Variant

    With ws.Range("A" & RowNumber & ":" & Cfg.LastCol & RowNumber)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 190, 200)
        .Interior.Color = RGB(226, 239, 218)
    End With

    If Len(Cfg.ChildSummaryCol) > 0 Then
        ws.Cells(RowNumber, Cfg.ChildSummaryCol).Interior.Color = RGB(242, 242, 242)
        ws.Cells(RowNumber, Cfg.ChildSummaryCol).Font.Italic = True
    End If

    For Each Col In Split(Cfg.IntCols, ",")
        If Len(Col) > 0 Then ws.Cells(RowNumber, CStr(Col)).NumberFormat = "0"
    Next Col

    For Each Col In Split(Cfg.MoneyCols, ",")
        If Len(Col) > 0 Then ws.Cells(RowNumber, CStr(Col)).NumberFormat = "#,##0.00"
    Next Col

End Sub

Private Sub ShowAllGroup(ByVal ws As Worksheet)
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    On Error GoTo 0
End Sub

Private Function LastGroupRow(ByVal ws As Worksheet) As Long
    LastGroupRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If LastGroupRow < HEADER_ROW Then LastGroupRow = HEADER_ROW
End Function

Private Function FirstVisibleGroupRow(ByVal ws As Worksheet, ByVal LastRow As Long, ByVal KeyValue As String) As Long

    Dim RowNumber As Long

    For RowNumber = FIRST_DATA_ROW To LastRow
        If Not ws.Rows(RowNumber).Hidden Then
            If Trim$(CStr(ws.Cells(RowNumber, "A").Value2)) = KeyValue Then
                FirstVisibleGroupRow = RowNumber
                Exit Function
            End If
        End If
    Next RowNumber

End Function
