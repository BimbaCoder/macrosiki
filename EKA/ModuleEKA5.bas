Attribute VB_Name = "ModuleEKA5"
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

Private Const GROUP_COUNT As Long = 15
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

        Case 1      ' EKAEMD_039 - Учредитель заемщика
            Cfg.GroupCode = "EKAEMD_039"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "AL"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "EKAEMD_039"
            Cfg.TitleText = "EKAEMD_039 — Учредитель заемщика"
            Cfg.LastCol = "F"
            Cfg.ValueCols = "B,C,D,E"
            Cfg.KeyCol = "G"
            Cfg.CtxCol = "H"
            Cfg.Headers = Array("EKAEMD_003 — ID проекта", _
                    "EKAEMD_040 — Правовой тип лица (SPR053)", _
                    "EKAEMD_041 — Физ.лицо/Юр.лицо, которому принадлежит более 1% долей участия в уставном капитале или размещенных акций предприятия завителя проекта (полное ФИО)", _
                    "EKAEMD_042 — Страна (указать страну по каждому учредителю отдельно) (SPR039)", _
                    "EKAEMD_043 — Доля, %", _
                    "EKAEMD_044 — Физ.лицо - Учредитель учредителя заемщика (список)")
            Cfg.Widths = Array(20, 19, 40, 38, 16, 26)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "SPR053", _
                    "", _
                    "SPR039", _
                    "", _
                    "")
            Cfg.ChildGroupCode = "EKAEMD_044"
            Cfg.ChildSummaryCol = "F"

        Case 2      ' EKAEMD_044 - Физ.лицо - Учредитель учредителя заемщика (вложена в EKAEMD_039)
            Cfg.GroupCode = "EKAEMD_044"
            Cfg.ParentSheetName = "EKAEMD_039"
            Cfg.ParentCol = "F"
            Cfg.ParentKeyCol = "G"
            Cfg.SheetName = "EKAEMD_044"
            Cfg.TitleText = "EKAEMD_044 — Физ.лицо - Учредитель учредителя заемщика"
            Cfg.LastCol = "D"
            Cfg.ValueCols = "B,C,D"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "E"
            Cfg.Headers = Array("Ключ (служебное)", _
                    "EKAEMD_045 — Физ.лицо, которому принадлежит более 1% долей участия в уставном капитале или размещенных акций предприятия завителя проекта (полное ФИО", _
                    "EKAEMD_046 — Страна (указать страну по каждому учредителю отдельно) (SPR039)", _
                    "EKAEMD_047 — Доля, %")
            Cfg.Widths = Array(20, 40, 38, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "SPR039", _
                    "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 3      ' EKAEMD_048 - Наименование иностранных инвесторов, ТНК, участвующих в проекте
            Cfg.GroupCode = "EKAEMD_048"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "CG"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "EKAEMD_048"
            Cfg.TitleText = "EKAEMD_048 — Наименование иностранных инвесторов, ТНК, участвующих в проекте"
            Cfg.LastCol = "C"
            Cfg.ValueCols = "B,C"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "D"
            Cfg.Headers = Array("EKAEMD_003 — ID проекта", _
                    "EKAEMD_049 — Наименование", _
                    "EKAEMD_050 — Страна (SPR039)")
            Cfg.Widths = Array(20, 16, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "SPR039")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 4      ' EKAEMD_057 - Планируемый объем финансирования (в случае валютного финансирования), валюта
            Cfg.GroupCode = "EKAEMD_057"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "AQ"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "EKAEMD_057"
            Cfg.TitleText = "EKAEMD_057 — Планируемый объем финансирования (в случае валютного финансирования), валюта"
            Cfg.LastCol = "C"
            Cfg.ValueCols = "B,C"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "D"
            Cfg.Headers = Array("EKAEMD_003 — ID проекта", _
                    "EKAEMD_058 — сумма", _
                    "EKAEMD_059 — валюта (SPR017)")
            Cfg.Widths = Array(20, 16, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "SPR017")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 5      ' EKAEMD_063 - Источник фондирования Холдинга
            Cfg.GroupCode = "EKAEMD_063"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "AT"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "EKAEMD_063"
            Cfg.TitleText = "EKAEMD_063 — Источник фондирования Холдинга"
            Cfg.LastCol = "R"
            Cfg.ValueCols = "B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "S"
            Cfg.Headers = Array("EKAEMD_003 — ID проекта", _
                    "EKAEMD_064 — Источник средств, (выбрать из списка один или несколько) (SPR015)", _
                    "EKAEMD_065 — Сумма по источнику средств", _
                    "EKAEMD_066 — Способ привлечения фондирования, (выбрать из списка) (SPR011)", _
                    "EKAEMD_068 — Наименование ФЭО", _
                    "EKAEMD_069 — № бюджетной программы", _
                    "EKAEMD_070 — Дата Приказа ГО", _
                    "EKAEMD_071 — № Приказа ГО", _
                    "EKAEMD_072 — Наименование ЦГО (список) (SPR046)", _
                    "EKAEMD_074 — Статус участия в гос. Программе (SPR047)", _
                    "EKAEMD_075 — Полное наименование программы", _
                    "EKAEMD_076 — Наменование НПА, которым утвержден документ", _
                    "EKAEMD_077 — Дата", _
                    "EKAEMD_078 — №", _
                    "EKAEMD_079 — Тип ставки привлечения, (SPR029)", _
                    "EKAEMD_080 — Ставка привлечения, % (при фиксированной)", _
                    "EKAEMD_082 — Индекс", _
                    "EKAEMD_083 — Спред")
            Cfg.Widths = Array(20, 39, 19, 37, 16, 17, 16, 16, 23, 26, 21, 28, 16, 16, 22, 27, 16, 16)
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
                    "SPR046", _
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

        Case 6      ' EKAEMD_104 - Договор
            Cfg.GroupCode = "EKAEMD_104"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "BQ"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "EKAEMD_104"
            Cfg.TitleText = "EKAEMD_104 — Договор"
            Cfg.LastCol = "T"
            Cfg.ValueCols = "B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S"
            Cfg.KeyCol = "U"
            Cfg.CtxCol = "V"
            Cfg.Headers = Array("EKAEMD_003 — ID проекта", _
                    "EKAEMD_111 — Номер договора гарантирования/страхования/входящего перестрахования/ДУБВ", _
                    "EKAEMD_112 — Дата договора гарантирования/страхования/входящего перестрахования/ДУБВ", _
                    "EKAEMD_120 — Сумма возмещения по регрессному требованию, в тенге", _
                    "EKAEMD_122 — Дата ввода информации по задолженности", _
                    "EKAEMD_124 — по обязательствам, тенге", _
                    "EKAEMD_125 — по вознаграждению/премии, тенге", _
                    "EKAEMD_127 — по обязательствам, тенге", _
                    "EKAEMD_128 — по вознаграждению/премии, тенге", _
                    "EKAEMD_129 — Дисконт (-), премия (+), отрицательная (-), положительная (+) корректировка стоимости займа", _
                    "EKAEMD_130 — Пеня, штрафы, тенге", _
                    "EKAEMD_131 — Количество дней просрочки", _
                    "EKAEMD_132 — Стадия (МСФО 9)", _
                    "EKAEMD_133 — Провизии, %", _
                    "EKAEMD_134 — Категория учета (амортизированная/справедливая) (SPR031)", _
                    "EKAEMD_135 — Провизии, тенге", _
                    "EKAEMD_143 — Дата заключения договора", _
                    "EKAEMD_144 — Дата окончания срока действия договора согласно графику погашения", _
                    "EKAEMD_145 — дни (дата окончания минус дата заключения)", _
                    "EKAEMD_113 — Страховое событие (список)")
            Cfg.Widths = Array(20, 40, 40, 32, 25, 18, 22, 18, 22, 40, 16, 19, 16, 16, 34, 16, 18, 39, 27, 26)
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
                    "", _
                    "", _
                    "SPR031", _
                    "", _
                    "", _
                    "", _
                    "", _
                    "")
            Cfg.ChildGroupCode = "EKAEMD_113"
            Cfg.ChildSummaryCol = "T"

        Case 7      ' EKAEMD_113 - Страховое событие (вложена в EKAEMD_104)
            Cfg.GroupCode = "EKAEMD_113"
            Cfg.ParentSheetName = "EKAEMD_104"
            Cfg.ParentCol = "T"
            Cfg.ParentKeyCol = "U"
            Cfg.SheetName = "EKAEMD_113"
            Cfg.TitleText = "EKAEMD_113 — Страховое событие"
            Cfg.LastCol = "F"
            Cfg.ValueCols = "B,C,D,E,F"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "G"
            Cfg.Headers = Array("Ключ (служебное)", _
                    "EKAEMD_114 — Дата наступления страхового события и (или) страхового случая", _
                    "EKAEMD_115 — Дата сообщения страховщику о наступлении страхового события и (или) страхового случая", _
                    "EKAEMD_116 — Статус (SPR051)", _
                    "EKAEMD_117 — Сумма выплаты, в тенге", _
                    "EKAEMD_118 — Дата начисления суммы страховой выплаты")
            Cfg.Widths = Array(20, 37, 40, 16, 17, 26)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "", _
                    "SPR051", _
                    "", _
                    "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 8      ' EKAEMD_151 - Освоение средств по годам
            Cfg.GroupCode = "EKAEMD_151"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "BV"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "EKAEMD_151"
            Cfg.TitleText = "EKAEMD_151 — Освоение средств по годам"
            Cfg.LastCol = "D"
            Cfg.ValueCols = "B,C,D"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "E"
            Cfg.Headers = Array("EKAEMD_003 — ID проекта", _
                    "EKAEMD_152 — Год (ГГГГ)", _
                    "EKAEMD_153 — План", _
                    "EKAEMD_154 — Факт")
            Cfg.Widths = Array(20, 16, 16, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "", _
                    "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 9      ' EKAEMD_157 - в стоимостном выражении, тенге
            Cfg.GroupCode = "EKAEMD_157"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "BW"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "EKAEMD_157"
            Cfg.TitleText = "EKAEMD_157 — в стоимостном выражении, тенге"
            Cfg.LastCol = "D"
            Cfg.ValueCols = "B,C,D"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "E"
            Cfg.Headers = Array("EKAEMD_003 — ID проекта", _
                    "EKAEMD_158 — Год отчетного периода", _
                    "EKAEMD_159 — план", _
                    "EKAEMD_160 — факт")
            Cfg.Widths = Array(20, 17, 16, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "", _
                    "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 10      ' EKAEMD_163 - в натуральном выражении (в единицах измерения проектной мощности в натуральном выражении)
            Cfg.GroupCode = "EKAEMD_163"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "BX"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "EKAEMD_163"
            Cfg.TitleText = "EKAEMD_163 — в натуральном выражении (в единицах измерения проектной мощности в натуральном выражении)"
            Cfg.LastCol = "D"
            Cfg.ValueCols = "B,C,D"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "E"
            Cfg.Headers = Array("EKAEMD_003 — ID проекта", _
                    "EKAEMD_164 — Год отчетного периода", _
                    "EKAEMD_165 — план", _
                    "EKAEMD_166 — факт")
            Cfg.Widths = Array(20, 17, 16, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "", _
                    "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 11      ' EKAEMD_169 - Выручка согласно Плану развития ДО/ДЗО, тенге
            Cfg.GroupCode = "EKAEMD_169"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "BY"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "EKAEMD_169"
            Cfg.TitleText = "EKAEMD_169 — Выручка согласно Плану развития ДО/ДЗО, тенге"
            Cfg.LastCol = "D"
            Cfg.ValueCols = "B,C,D"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "E"
            Cfg.Headers = Array("EKAEMD_003 — ID проекта", _
                    "EKAEMD_170 — Год отчетного периода", _
                    "EKAEMD_171 — план", _
                    "EKAEMD_172 — факт")
            Cfg.Widths = Array(20, 17, 16, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "", _
                    "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 12      ' EKAEMD_176 - Текущее количество рабочих мест, занятых на предприятии, ед.
            Cfg.GroupCode = "EKAEMD_176"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "BZ"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "EKAEMD_176"
            Cfg.TitleText = "EKAEMD_176 — Текущее количество рабочих мест, занятых на предприятии, ед."
            Cfg.LastCol = "D"
            Cfg.ValueCols = "B,C,D"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "E"
            Cfg.Headers = Array("EKAEMD_003 — ID проекта", _
                    "EKAEMD_177 — Год отчетного периода", _
                    "EKAEMD_178 — план", _
                    "EKAEMD_179 — факт")
            Cfg.Widths = Array(20, 17, 16, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "", _
                    "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 13      ' EKAEMD_180 - Количество созданных рабочих мест, ед.
            Cfg.GroupCode = "EKAEMD_180"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "CA"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "EKAEMD_180"
            Cfg.TitleText = "EKAEMD_180 — Количество созданных рабочих мест, ед."
            Cfg.LastCol = "E"
            Cfg.ValueCols = "B,C,D,E"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "F"
            Cfg.Headers = Array("EKAEMD_003 — ID проекта", _
                    "EKAEMD_181 — Год отчетного периода", _
                    "EKAEMD_182 — план", _
                    "EKAEMD_183 — факт", _
                    "EKAEMD_185 — факт (кумулятивно за все года)")
            Cfg.Widths = Array(20, 17, 16, 16, 21)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "", _
                    "", _
                    "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 14      ' EKAEMD_186 - Уплаченные налоги и другие обязательные платежи в бюджет, тенге
            Cfg.GroupCode = "EKAEMD_186"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "CB"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "EKAEMD_186"
            Cfg.TitleText = "EKAEMD_186 — Уплаченные налоги и другие обязательные платежи в бюджет, тенге"
            Cfg.LastCol = "E"
            Cfg.ValueCols = "B,C,D,E"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "F"
            Cfg.Headers = Array("EKAEMD_003 — ID проекта", _
                    "EKAEMD_187 — Год отчетного периода", _
                    "EKAEMD_188 — план", _
                    "EKAEMD_189 — факт", _
                    "EKAEMD_191 — факт (кумулятивно за все года)")
            Cfg.Widths = Array(20, 17, 16, 16, 21)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "", _
                    "", _
                    "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 15      ' EKAEMD_192 - Встречные обязательства
            Cfg.GroupCode = "EKAEMD_192"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "CC"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "EKAEMD_192"
            Cfg.TitleText = "EKAEMD_192 — Встречные обязательства"
            Cfg.LastCol = "E"
            Cfg.ValueCols = "B,C,D,E"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "F"
            Cfg.Headers = Array("EKAEMD_003 — ID проекта", _
                    "EKAEMD_193 — Год отчетного периода", _
                    "EKAEMD_194 — план", _
                    "EKAEMD_195 — факт", _
                    "EKAEMD_196 — статус исполнения")
            Cfg.Widths = Array(20, 17, 16, 16, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
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
