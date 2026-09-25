Attribute VB_Name = "ModuleLSG5"
Option Explicit

'==================================================================
' modNestedGroups - множественные атрибуты со сводной ячейкой.
' Сгенерировано автоматически по листу параметрического представления.
'
' Логика:
'   - на главном листе одна ячейка-сводка на группу (код в строке
'     MULTI_ATTR_CODE_ROW), двойной клик открывает лист группы
'   - если у группы есть вложенная группа - внутри ее листа есть
'     своя сводная ячейка, открывающая вложенный список,
'     отфильтрованный по конкретной родительской строке (через
'     уникальный ключ)
'   - при возврате в сводной ячейке родителя проставляется
'     "Заполнен N атрибутами"
'
' ВАЖНО: IntCols/MoneyCols (числовой/денежный формат) оставлены
' пустыми - расставьте по смыслу полей, где нужно (в отличие от
' КЖК, тут это не мог сделать автоматически с уверенностью).
'==================================================================

Private Const GROUP_COUNT As Long = 13
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

        Case 1      ' LsgEMD_018 - Сведения об основном средстве
            Cfg.GroupCode = "LsgEMD_018"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "CH"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "LsgEMD_018"
            Cfg.TitleText = "LsgEMD_018 — Сведения об основном средстве"
            Cfg.LastCol = "A"
            Cfg.ValueCols = ""
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "B"
            Cfg.Headers = Array("LsgEMD_003 — ID проекта")
            Cfg.Widths = Array(20)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 2      ' LsgEMD_031 - Учредитель заемщика
            Cfg.GroupCode = "LsgEMD_031"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "CI"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "LsgEMD_031"
            Cfg.TitleText = "LsgEMD_031 — Учредитель заемщика"
            Cfg.LastCol = "F"
            Cfg.ValueCols = "B,C,D,E"
            Cfg.KeyCol = "G"
            Cfg.CtxCol = "H"
            Cfg.Headers = Array("LsgEMD_003 — ID проекта", _
                    "LsgEMD_157 — Правовой тип лица (SPR053)", _
                    "LsgEMD_032 — Физ.лицо, которому принадлежит более 1% долей участия в уставном капитале или размещенных акций предприятия завителя проекта (полное ФИО)", _
                    "LsgEMD_033 — Страна (указать страну по каждому учредителю отдельно) (SPR039)", _
                    "LsgEMD_034 — Доля, %", _
                    "LsgEMD_158 — Физ.лицо - Учредитель учредителя заемщика (список)")
            Cfg.Widths = Array(20, 19, 40, 38, 16, 26)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "SPR053", _
                    "", _
                    "SPR039", _
                    "", _
                    "")
            Cfg.ChildGroupCode = "LsgEMD_158"
            Cfg.ChildSummaryCol = "F"

        Case 3      ' LsgEMD_158 - Физ.лицо - Учредитель учредителя заемщика (вложена в LsgEMD_031)
            Cfg.GroupCode = "LsgEMD_158"
            Cfg.ParentSheetName = "LsgEMD_031"
            Cfg.ParentCol = "F"
            Cfg.ParentKeyCol = "G"
            Cfg.SheetName = "LsgEMD_158"
            Cfg.TitleText = "LsgEMD_158 — Физ.лицо - Учредитель учредителя заемщика"
            Cfg.LastCol = "D"
            Cfg.ValueCols = "B,C,D"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "E"
            Cfg.Headers = Array("Ключ (служебное)", _
                    "LsgEMD_159 — Физ.лицо, которому принадлежит более 1% долей участия в уставном капитале или размещенных акций предприятия завителя проекта (полное ФИО", _
                    "LsgEMD_160 — Страна (указать страну по каждому учредителю отдельно) (SPR039)", _
                    "LsgEMD_161 — Доля, %")
            Cfg.Widths = Array(20, 40, 38, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "SPR039", _
                    "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 4      ' LsgEMD_035 - Наименование иностранных инвесторов, ТНК, участвующих в проекте
            Cfg.GroupCode = "LsgEMD_035"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "AE"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "LsgEMD_035"
            Cfg.TitleText = "LsgEMD_035 — Наименование иностранных инвесторов, ТНК, участвующих в проекте"
            Cfg.LastCol = "C"
            Cfg.ValueCols = "B,C"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "D"
            Cfg.Headers = Array("LsgEMD_003 — ID проекта", _
                    "LsgEMD_036 — Наименование", _
                    "LsgEMD_037 — Страна (SPR039)")
            Cfg.Widths = Array(20, 16, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "SPR039")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 5      ' LsgEMD_043 - Планируемый объем финансирования (в случае валютного финансирования), валюта
            Cfg.GroupCode = "LsgEMD_043"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "AI"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "LsgEMD_043"
            Cfg.TitleText = "LsgEMD_043 — Планируемый объем финансирования (в случае валютного финансирования), валюта"
            Cfg.LastCol = "C"
            Cfg.ValueCols = "B,C"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "D"
            Cfg.Headers = Array("LsgEMD_003 — ID проекта", _
                    "LsgEMD_044 — Сумма", _
                    "LsgEMD_045 — валюта (SPR017)")
            Cfg.Widths = Array(20, 16, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "SPR017")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 6      ' LsgEMD_049 - Источник фондирования Холдинга
            Cfg.GroupCode = "LsgEMD_049"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "CJ"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "LsgEMD_049"
            Cfg.TitleText = "LsgEMD_049 — Источник фондирования Холдинга"
            Cfg.LastCol = "R"
            Cfg.ValueCols = "B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "S"
            Cfg.Headers = Array("LsgEMD_003 — ID проекта", _
                    "LsgEMD_050 — Собственные средства ДО/ДЗО (SPR015)", _
                    "LsgEMD_051 — Сумма по источнику средств", _
                    "LsgEMD_052 — Пополнение УК за счет бюджетных средств (SPR011)", _
                    "LsgEMD_054 — Наименование ФЭО", _
                    "LsgEMD_055 — № бюджетной программы", _
                    "LsgEMD_056 — Дата Приказа ГО", _
                    "LsgEMD_057 — № Приказа ГО", _
                    "LsgEMD_058 — Наименование ЦГО (список) (SPR046 )", _
                    "LsgEMD_060 — Статус участия в гос. Программах (SPR047)", _
                    "LsgEMD_061 — Полное наименование программы", _
                    "LsgEMD_062 — Наменование НПА, которым утвержден документ", _
                    "LsgEMD_063 — Дата", _
                    "LsgEMD_064 — №", _
                    "LsgEMD_065 — Тип ставки привлечения, (фиксированная/плавающая) (SPR029)", _
                    "LsgEMD_066 — Ставка привлечения, % (при фиксированной)", _
                    "LsgEMD_068 — Индекс", _
                    "LsgEMD_069 — Спред")
            Cfg.Widths = Array(20, 24, 19, 30, 16, 17, 16, 16, 24, 27, 21, 28, 16, 16, 35, 27, 16, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "SPR015", _
                    "", _
                    "SPR011", _
                    "", _
                    "", _
                    "", _
                    "", _
                    "SPR046 ", _
                    "SPR047", _
                    "", _
                    "", _
                    "", _
                    "", _
                    "SPR029", _
                    "", _
                    "", _
                    "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 7      ' LsgEMD_086 - Договор
            Cfg.GroupCode = "LsgEMD_086"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "BU"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "LsgEMD_086"
            Cfg.TitleText = "LsgEMD_086 — Договор"
            Cfg.LastCol = "R"
            Cfg.ValueCols = "B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q"
            Cfg.KeyCol = "S"
            Cfg.CtxCol = "T"
            Cfg.Headers = Array("LsgEMD_003 — ID проекта", _
                    "LsgEMD_093 — Номер СОКЛ/договора/догова банковской гарантии", _
                    "LsgEMD_094 — Дата СОКЛ/договора/догова банковской гарантии", _
                    "LsgEMD_097 — Дата ввода информации по задолженности", _
                    "LsgEMD_099 — по основному долгу, тенге", _
                    "LsgEMD_100 — по вознаграждению, тенге", _
                    "LsgEMD_102 — по основному долгу, тенге", _
                    "LsgEMD_103 — по вознаграждению, тенге", _
                    "LsgEMD_104 — Дисконт (-), премия (+), отрицательная (-), положительная (+) корректировка стоимости займа", _
                    "LsgEMD_105 — Пеня, штрафы, тенге", _
                    "LsgEMD_106 — Количество дней просрочки", _
                    "LsgEMD_107 — Стадия (МСФО 9)", _
                    "LsgEMD_109 — Категория учета (амортизированная/справедливая) (SPR031)", _
                    "LsgEMD_110 — Провизии, тенге", _
                    "LsgEMD_117 — Дата заключения договора лизинга", _
                    "LsgEMD_118 — Дата окончания срока действия договора лизинга согласно графику погашения", _
                    "LsgEMD_119 — дни (дата окончания минус дата заключения)", _
                    "LsgEMD_111 — Реструктуризация (список)")
            Cfg.Widths = Array(20, 29, 29, 25, 19, 18, 19, 18, 40, 16, 19, 16, 34, 16, 22, 40, 27, 26)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "", _
                    "", _
                    "", _
                    "", _
                    "", _
                    "", _
                    "", _
                    "", _
                    "", _
                    "", _
                    "SPR031", _
                    "", _
                    "", _
                    "", _
                    "", _
                    "")
            Cfg.ChildGroupCode = "LsgEMD_111"
            Cfg.ChildSummaryCol = "R"

        Case 8      ' LsgEMD_111 - Реструктуризация (вложена в LsgEMD_086)
            Cfg.GroupCode = "LsgEMD_111"
            Cfg.ParentSheetName = "LsgEMD_086"
            Cfg.ParentCol = "R"
            Cfg.ParentKeyCol = "S"
            Cfg.SheetName = "LsgEMD_111"
            Cfg.TitleText = "LsgEMD_111 — Реструктуризация"
            Cfg.LastCol = "F"
            Cfg.ValueCols = "B,C,D,E,F"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "G"
            Cfg.Headers = Array("Ключ (служебное)", _
                    "LsgEMD_112 — Дата реструктуризации", _
                    "LsgEMD_113 — Основные условия по реструктуризации", _
                    "LsgEMD_114 — Краткое описание решений Уполномоченного органа ДО/ДЗО/поручения Правительства или ГО относительно проблемных вопросов (при наличии)", _
                    "LsgEMD_115 — Дальнейший План мероприятий/предложения/видение ДО/ДЗО по оздоровлению проекта (прикрепить соответствующий документ)", _
                    "LsgEMD_156 — Документ дальнейшего Плана мероприятий/предложения/видение ДО/ДЗО по оздоровлению проекта")
            Cfg.Widths = Array(20, 17, 24, 40, 40, 40)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "", _
                    "", _
                    "", _
                    "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 9      ' LsgEMD_125 - Освоение средств по годам
            Cfg.GroupCode = "LsgEMD_125"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "CK"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "LsgEMD_125"
            Cfg.TitleText = "LsgEMD_125 — Освоение средств по годам"
            Cfg.LastCol = "D"
            Cfg.ValueCols = "B,C,D"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "E"
            Cfg.Headers = Array("LsgEMD_003 — ID проекта", _
                    "LsgEMD_126 — Год (ГГГГ)", _
                    "LsgEMD_127 — План", _
                    "LsgEMD_128 — Факт")
            Cfg.Widths = Array(20, 16, 16, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "", _
                    "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 10      ' LsgEMD_131 - Текущее количество рабочих мест, занятых на предприятии, ед.
            Cfg.GroupCode = "LsgEMD_131"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "CL"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "LsgEMD_131"
            Cfg.TitleText = "LsgEMD_131 — Текущее количество рабочих мест, занятых на предприятии, ед."
            Cfg.LastCol = "B"
            Cfg.ValueCols = "B"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "C"
            Cfg.Headers = Array("LsgEMD_003 — ID проекта", _
                    "LsgEMD_132 — Год отчетного периода")
            Cfg.Widths = Array(20, 17)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 11      ' LsgEMD_135 - Количество созданных рабочих мест, ед.
            Cfg.GroupCode = "LsgEMD_135"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "CM"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "LsgEMD_135"
            Cfg.TitleText = "LsgEMD_135 — Количество созданных рабочих мест, ед."
            Cfg.LastCol = "B"
            Cfg.ValueCols = "B"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "C"
            Cfg.Headers = Array("LsgEMD_003 — ID проекта", _
                    "LsgEMD_136 — Год отчетного периода")
            Cfg.Widths = Array(20, 17)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 12      ' LsgEMD_141 - Уплаченные налоги и другие обязательные платежи в бюджет, тенге
            Cfg.GroupCode = "LsgEMD_141"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "CN"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "LsgEMD_141"
            Cfg.TitleText = "LsgEMD_141 — Уплаченные налоги и другие обязательные платежи в бюджет, тенге"
            Cfg.LastCol = "F"
            Cfg.ValueCols = "B,C,D,E,F"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "G"
            Cfg.Headers = Array("LsgEMD_003 — ID проекта", _
                    "LsgEMD_142 — Год отчетного периода", _
                    "LsgEMD_143 — план", _
                    "LsgEMD_144 — факт", _
                    "LsgEMD_145 — план (кумулятивно за все года)", _
                    "LsgEMD_146 — факт (кумулятивно за все года)")
            Cfg.Widths = Array(20, 17, 16, 16, 21, 21)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "", _
                    "", _
                    "", _
                    "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 13      ' LsgEMD_147 - Встречные обязательства
            Cfg.GroupCode = "LsgEMD_147"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "CO"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "LsgEMD_147"
            Cfg.TitleText = "LsgEMD_147 — Встречные обязательства"
            Cfg.LastCol = "D"
            Cfg.ValueCols = "B,C,D"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "E"
            Cfg.Headers = Array("LsgEMD_003 — ID проекта", _
                    "LsgEMD_148 — Год отчетного периода", _
                    "LsgEMD_149 — план", _
                    "LsgEMD_150 — факт")
            Cfg.Widths = Array(20, 17, 16, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "", _
                    "")
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
            MsgBox "Сначала заполните ID проекта в столбце " & _
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

        .Range("BZ1").Value2 = 0

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

    With ws.Range("A" & RowNumber & ":" & Cfg.LastCol & RowNumber)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 190, 200)
        .Interior.Color = RGB(226, 239, 218)
    End With

    If Len(Cfg.ChildSummaryCol) > 0 Then
        ws.Cells(RowNumber, Cfg.ChildSummaryCol).Interior.Color = RGB(242, 242, 242)
        ws.Cells(RowNumber, Cfg.ChildSummaryCol).Font.Italic = True
    End If

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
