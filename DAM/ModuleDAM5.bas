Attribute VB_Name = "ModuleDAM5"
Option Explicit

'==================================================================
' modNestedGroups - множественные атрибуты со сводной ячейкой.
'
' ОБЩИЙ МОДУЛЬ НА ВЕСЬ ПРОЕКТ: файл одинаковый во всех книгах
' (КЖК, ДАМУ, Лизинг), отличается только строкой Attribute VB_Name.
' Внутри - белые списки групп всех систем; нужный выбирается по
' MAIN_SHEET_NAME из Module3 (см. SystemCode). Правки логики и
' белых списков делать ЗДЕСЬ и копировать файл во все книги.
' АКК СХ и ЭКА пока на своих (старых) Module5 - в этот модуль
' не переведены.
'
' Белый список группы = GroupCode + поля Headers/ValueCols.
' Правило построения (по файлу Spravochka в корне проекта):
'   - в список полей группы попадают ЛИСТЬЯ поддерева, т.е. коды,
'     у которых в Spravochka нет своих дочерних (пометка
'     "Составной" в реестре тут не важна - например LsgEMD_108,
'     LsgEMD_133/134, LsgEMD_137-140 составные, но полей под
'     ними нет, значит это обычные поля);
'   - промежуточные узлы с дочерними (KJKEMD_054, KJKEMD_096 ...)
'     в столбцы не попадают - только их листья;
'   - вложенный МНОЖЕСТВЕННЫЙ атрибут (KJKEMD_176, KJKEMD_112,
'     DamuEMD_108, LsgEMD_158, LsgEMD_111) - отдельная группа
'     со своим листом, у родителя под него сводный столбец
'     ChildSummaryCol, а его листья в родителя НЕ входят.
'
' Вложенности ("лесенка"):
'   КЖК:    KJKEMD_032 -> KJKEMD_176,  KJKEMD_087 -> KJKEMD_112
'   ДАМУ:   DamuEMD_138 -> DamuEMD_108
'   Лизинг: LsgEMD_031 -> LsgEMD_158,  LsgEMD_086 -> LsgEMD_111
'
' Логика:
'   - на главном листе одна ячейка-сводка на группу (ParentCol);
'     двойной клик по ней открывает лист группы. Группа узнается
'     по коду в строке MULTI_ATTR_CODE_ROW ИЛИ по самому столбцу
'     ParentCol (у Лизинга 131/135/141/147 в этих столбцах стоит
'     код первого поля, а не код группы)
'   - если у группы есть вложенная - внутри ее листа есть своя
'     сводная ячейка (ChildSummaryCol), открывающая лист вложенной
'     группы, отфильтрованный по конкретной родительской строке
'     (через уникальный ключ в KeyCol)
'   - при возврате в сводной ячейке родителя проставляется
'     "Заполнен N атрибутами"
'==================================================================

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
' ВЫБОР СИСТЕМЫ (по главному листу книги)
'------------------------------------------------------------------
Private Function SystemCode() As String

    Select Case MAIN_SHEET_NAME
        Case "КЖК_ЭФ":    SystemCode = "KJK"
        Case "ДАМУ ЭФ":   SystemCode = "DAM"
        Case "ЭФ_Лизинг": SystemCode = "LSG"
        Case Else:        SystemCode = vbNullString
    End Select

End Function

Private Function GroupCount() As Long

    Select Case SystemCode()
        Case "KJK": GroupCount = 8
        Case "DAM": GroupCount = 11
        Case "LSG": GroupCount = 11
        Case Else:  GroupCount = 0
    End Select

End Function

Private Sub LoadGroupCfg(ByVal Index As Long, ByRef Cfg As TGroup)

    Dim Blank As TGroup

    Cfg = Blank

    Select Case SystemCode()
        Case "KJK": LoadGroupCfg_KJK Index, Cfg
        Case "DAM": LoadGroupCfg_DAM Index, Cfg
        Case "LSG": LoadGroupCfg_LSG Index, Cfg
    End Select

End Sub

'------------------------------------------------------------------
' БЕЛЫЙ СПИСОК: КЖК
'------------------------------------------------------------------
Private Sub LoadGroupCfg_KJK(ByVal Index As Long, ByRef Cfg As TGroup)

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
' БЕЛЫЙ СПИСОК: ДАМУ
'------------------------------------------------------------------
Private Sub LoadGroupCfg_DAM(ByVal Index As Long, ByRef Cfg As TGroup)

    Select Case Index

        Case 1      ' DamuEMD_029 - Планируемый объем инвестиций в проект (в случае валютного финансирования), валюта
            Cfg.GroupCode = "DamuEMD_029"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "Z"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "DamuEMD_029"
            Cfg.TitleText = "DamuEMD_029 — Планируемый объем инвестиций в проект (в случае валютного финансирования), валюта"
            Cfg.LastCol = "C"
            Cfg.ValueCols = "B,C"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "D"
            Cfg.Headers = Array("DamuEMD_004 — ID проекта", _
                    "DamuEMD_030 — сумма", _
                    "DamuEMD_031 — валюта (SPR017)")
            Cfg.Widths = Array(20, 16, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "SPR017")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 2      ' DamuEMD_035 - Источник фондирования Холдинга
            Cfg.GroupCode = "DamuEMD_035"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "AC"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "DamuEMD_035"
            Cfg.TitleText = "DamuEMD_035 — Источник фондирования Холдинга"
            Cfg.LastCol = "R"
            Cfg.ValueCols = "B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "S"
            Cfg.Headers = Array("DamuEMD_004 — ID проекта", _
                    "DamuEMD_036 — Источник средств, (выбрать из списка один или несколько) (SPR015)", _
                    "DamuEMD_037 — Сумма по источнику средств", _
                    "DamuEMD_038 — Способ привлечения фондирования, (выбрать из списка) (SPR011)", _
                    "DamuEMD_040 — Наименование ФЭО", _
                    "DamuEMD_041 — № бюджетной программы", _
                    "DamuEMD_042 — Дата Приказа ГО", _
                    "DamuEMD_043 — № Приказа ГО", _
                    "DamuEMD_044 — Наименование ЦГО (список) (SPR046)", _
                    "DamuEMD_046 — Признак участия в гос. Программах (SPR047)", _
                    "DamuEMD_047 — Полное наименование программы", _
                    "DamuEMD_048 — Наменование НПА, которым утвержден документ", _
                    "DamuEMD_049 — Дата", _
                    "DamuEMD_050 — №", _
                    "DamuEMD_051 — Тип ставки привлечения (SPR029)", _
                    "DamuEMD_052 — Ставка привлечения, % (при фиксированной)", _
                    "DamuEMD_054 — Индекс", _
                    "DamuEMD_055 — Спред")
            Cfg.Widths = Array(20, 39, 20, 37, 16, 17, 16, 16, 24, 28, 21, 28, 16, 16, 22, 27, 16, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", "SPR015", "", "SPR011", "", "", "", "", "SPR046", _
                                   "SPR047", "", "", "", "", "SPR029", "", "", "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 3      ' DamuEMD_138 - Договор
            Cfg.GroupCode = "DamuEMD_138"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "AV"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "DamuEMD_138"
            Cfg.TitleText = "DamuEMD_138 — Договор"
            Cfg.LastCol = "S"
            Cfg.ValueCols = "B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R"
            Cfg.KeyCol = "T"
            Cfg.CtxCol = "U"
            Cfg.Headers = Array("DamuEMD_004 — ID проекта", _
                    "DamuEMD_139 — Номер СОКЛ/договора/догова банковской гарантии", _
                    "DamuEMD_082 — Дата СОКЛ/договора/догова банковской гарантии", _
                    "DamuEMD_094 — Дата ввода информации по задолженности", _
                    "DamuEMD_096 — по основному долгу, тенге", _
                    "DamuEMD_097 — по вознаграждению, тенге", _
                    "DamuEMD_099 — по основному долгу, тенге", _
                    "DamuEMD_100 — по вознаграждению, тенге", _
                    "DamuEMD_101 — Дисконт (-), премия (+), отрицательная (-), положительная (+) корректировка стоимости займа", _
                    "DamuEMD_102 — Пеня, штрафы, тенге", _
                    "DamuEMD_103 — Количество дней просрочки", _
                    "DamuEMD_104 — Стадия (МСФО 9)", _
                    "DamuEMD_105 — Провизии, %", _
                    "DamuEMD_106 — Категория учета (SPR031)", _
                    "DamuEMD_107 — Провизии, тенге", _
                    "DamuEMD_072 — Дата заключения договора займа", _
                    "DamuEMD_073 — Дата окончания срока действия договора займа согласно графику погашения", _
                    "DamuEMD_074 — дни (дата окончания минус дата заключения)", _
                    "DamuEMD_108 — Реструктуризация (список)")
            Cfg.Widths = Array(20, 30, 29, 26, 19, 19, 19, 19, 40, 16, 19, 16, 16, 19, 16, 22, 40, 28, 26)
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
                    "SPR031", _
                    "", _
                    "", _
                    "", _
                    "", _
                    "")
            Cfg.ChildGroupCode = "DamuEMD_108"
            Cfg.ChildSummaryCol = "S"

        Case 4      ' DamuEMD_108 - Реструктуризация (вложена в DamuEMD_138)
            Cfg.GroupCode = "DamuEMD_108"
            Cfg.ParentSheetName = "DamuEMD_138"
            Cfg.ParentCol = "S"
            Cfg.ParentKeyCol = "T"
            Cfg.SheetName = "DamuEMD_108"
            Cfg.TitleText = "DamuEMD_108 — Реструктуризация"
            Cfg.LastCol = "F"
            Cfg.ValueCols = "B,C,D,E,F"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "G"
            Cfg.Headers = Array("Ключ (служебное)", _
                    "DamuEMD_109 — Дата реструктуризации", _
                    "DamuEMD_110 — Основные условия по реструктуризации", _
                    "DamuEMD_111 — Краткое описание решений Уполномоченного органа ДО/ДЗО/поручения Правительства или ГО относительно проблемных вопросов (при наличии)", _
                    "DamuEMD_112 — Дальнейший План мероприятий/предложения/видение ДО/ДЗО по оздоровлению проекта (прикрепить соответствующий документ)", _
                    "DamuEMD_140 — Документ дальнейшего Плана мероприятий/предложения/видение ДО/ДЗО по оздоровлению проекта")
            Cfg.Widths = Array(20, 17, 25, 40, 40, 40)
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

        Case 5      ' DamuEMD_088 - Освоение средств по годам
            Cfg.GroupCode = "DamuEMD_088"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "BA"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "DamuEMD_088"
            Cfg.TitleText = "DamuEMD_088 — Освоение средств по годам"
            Cfg.LastCol = "D"
            Cfg.ValueCols = "B,C,D"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "E"
            Cfg.Headers = Array("DamuEMD_004 — ID проекта", _
                    "DamuEMD_089 — Год (ГГГГ)", _
                    "DamuEMD_090 — План", _
                    "DamuEMD_091 — Факт")
            Cfg.Widths = Array(20, 16, 16, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "", _
                    "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 6      ' DamuEMD_114 - Выручка согласно Плану развития ДО/ДЗО, тенге
            Cfg.GroupCode = "DamuEMD_114"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "BB"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "DamuEMD_114"
            Cfg.TitleText = "DamuEMD_114 — Выручка согласно Плану развития ДО/ДЗО, тенге"
            Cfg.LastCol = "F"
            Cfg.ValueCols = "B,C,D,E,F"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "G"
            Cfg.Headers = Array("DamuEMD_004 — ID проекта", _
                    "DamuEMD_149 — Год отчетного периода", _
                    "DamuEMD_115 — план (отчетный период)", _
                    "DamuEMD_116 — факт", _
                    "DamuEMD_150 — план (кумулятивно за все года)", _
                    "DamuEMD_117 — факт (кумулятивно за все года)")
            Cfg.Widths = Array(20, 17, 18, 16, 22, 22)
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

        Case 7      ' DamuEMD_119 - Текущее количество рабочих мест, занятых на предприятии, ед.
            Cfg.GroupCode = "DamuEMD_119"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "BC"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "DamuEMD_119"
            Cfg.TitleText = "DamuEMD_119 — Текущее количество рабочих мест, занятых на предприятии, ед."
            Cfg.LastCol = "D"
            Cfg.ValueCols = "B,C,D"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "E"
            Cfg.Headers = Array("DamuEMD_004 — ID проекта", _
                    "DamuEMD_141 — Год отчетного периода", _
                    "DamuEMD_120 — план", _
                    "DamuEMD_121 — факт")
            Cfg.Widths = Array(20, 17, 16, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", _
                    "", _
                    "", _
                    "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 8      ' DamuEMD_122 - Количество созданных рабочих мест, ед.
            Cfg.GroupCode = "DamuEMD_122"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "BD"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "DamuEMD_122"
            Cfg.TitleText = "DamuEMD_122 — Количество созданных рабочих мест, ед."
            Cfg.LastCol = "F"
            Cfg.ValueCols = "B,C,D,E,F"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "G"
            Cfg.Headers = Array("DamuEMD_004 — ID проекта", _
                    "DamuEMD_142 — Год отчетного периода", _
                    "DamuEMD_151 — план", _
                    "DamuEMD_123 — факт", _
                    "DamuEMD_152 — План (кумулятивно за все года)", _
                    "DamuEMD_124 — факт (кумулятивно за все года)")
            Cfg.Widths = Array(20, 17, 16, 16, 22, 22)
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

        Case 9      ' DamuEMD_125 - Фонд оплаты труда, тенге
            Cfg.GroupCode = "DamuEMD_125"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "BE"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "DamuEMD_125"
            Cfg.TitleText = "DamuEMD_125 — Фонд оплаты труда, тенге"
            Cfg.LastCol = "F"
            Cfg.ValueCols = "B,C,D,E,F"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "G"
            Cfg.Headers = Array("DamuEMD_004 — ID проекта", _
                    "DamuEMD_143 — Год отчетного периода", _
                    "DamuEMD_153 — план", _
                    "DamuEMD_126 — факт", _
                    "DamuEMD_154 — План (кумулятивно за все года)", _
                    "DamuEMD_144 — факт (кумулятивно за все года)")
            Cfg.Widths = Array(20, 17, 16, 16, 22, 22)
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

        Case 10      ' DamuEMD_127 - Уплаченные налоги и другие обязательные платежи в бюджет, тенге
            Cfg.GroupCode = "DamuEMD_127"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "BF"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "DamuEMD_127"
            Cfg.TitleText = "DamuEMD_127 — Уплаченные налоги и другие обязательные платежи в бюджет, тенге"
            Cfg.LastCol = "F"
            Cfg.ValueCols = "B,C,D,E,F"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "G"
            Cfg.Headers = Array("DamuEMD_004 — ID проекта", _
                    "DamuEMD_145 — Год отчетного периода", _
                    "DamuEMD_155 — План", _
                    "DamuEMD_128 — факт", _
                    "DamuEMD_156 — план (кумулятивно за все года)", _
                    "DamuEMD_129 — факт (кумулятивно за все года)")
            Cfg.Widths = Array(20, 17, 16, 16, 22, 22)
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

        Case 11      ' DamuEMD_130 - Встречные обязательства
            Cfg.GroupCode = "DamuEMD_130"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "BG"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "DamuEMD_130"
            Cfg.TitleText = "DamuEMD_130 — Встречные обязательства"
            Cfg.LastCol = "F"
            Cfg.ValueCols = "B,C,D,E,F"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "G"
            Cfg.Headers = Array("DamuEMD_004 — ID проекта", _
                    "DamuEMD_146 — Год отчетного периода", _
                    "DamuEMD_131 — план", _
                    "DamuEMD_132 — факт", _
                    "DamuEMD_147 — план (кумулятивно за все года)", _
                    "DamuEMD_148 — факт (кумулятивно за все года)")
            Cfg.Widths = Array(20, 17, 16, 16, 22, 22)
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

    End Select

End Sub

'------------------------------------------------------------------
' БЕЛЫЙ СПИСОК: Лизинг
' (LsgEMD_018 и LsgEMD_049 на главном листе разложены отдельными
' столбцами S:W и AL:BB - остаются одинарными, групп под них нет)
'------------------------------------------------------------------
Private Sub LoadGroupCfg_LSG(ByVal Index As Long, ByRef Cfg As TGroup)

    Select Case Index

        Case 1      ' LsgEMD_031 - Учредитель заемщика (вложена 158)
            Cfg.GroupCode = "LsgEMD_031"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "AD"
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

        Case 2      ' LsgEMD_158 - Физ.лицо - Учредитель учредителя заемщика (вложена в LsgEMD_031)
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

        Case 3      ' LsgEMD_035 - Наименование иностранных инвесторов, ТНК, участвующих в проекте
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

        Case 4      ' LsgEMD_043 - Планируемый объем финансирования (в случае валютного финансирования), валюта
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

        Case 5      ' LsgEMD_086 - Договор (вложена 111)
            Cfg.GroupCode = "LsgEMD_086"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "BU"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "LsgEMD_086"
            Cfg.TitleText = "LsgEMD_086 — Договор"
            Cfg.LastCol = "S"
            Cfg.ValueCols = "B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R"
            Cfg.KeyCol = "T"
            Cfg.CtxCol = "U"
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
                    "LsgEMD_108 — Провизии, %", _
                    "LsgEMD_109 — Категория учета (амортизированная/справедливая) (SPR031)", _
                    "LsgEMD_110 — Провизии, тенге", _
                    "LsgEMD_117 — Дата заключения договора лизинга", _
                    "LsgEMD_118 — Дата окончания срока действия договора лизинга согласно графику погашения", _
                    "LsgEMD_119 — дни (дата окончания минус дата заключения)", _
                    "LsgEMD_111 — Реструктуризация (список)")
            Cfg.Widths = Array(20, 29, 29, 25, 19, 18, 19, 18, 40, 16, 19, 16, 14, 34, 16, 22, 40, 27, 26)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", "", "", "", "", "", "", "", "", "", "", "", "", _
                                   "SPR031", "", "", "", "", "")
            Cfg.ChildGroupCode = "LsgEMD_111"
            Cfg.ChildSummaryCol = "S"

        Case 6      ' LsgEMD_111 - Реструктуризация (вложена в LsgEMD_086)
            Cfg.GroupCode = "LsgEMD_111"
            Cfg.ParentSheetName = "LsgEMD_086"
            Cfg.ParentCol = "S"
            Cfg.ParentKeyCol = "T"
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
            Cfg.DictCodes = Array("", "", "", "", "", "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 7      ' LsgEMD_125 - Освоение средств по годам
            Cfg.GroupCode = "LsgEMD_125"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "BZ"
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
            Cfg.DictCodes = Array("", "", "", "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 8      ' LsgEMD_131 - Текущее количество рабочих мест (сводка в столбце CA, там код LsgEMD_132)
            Cfg.GroupCode = "LsgEMD_131"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "CA"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "LsgEMD_131"
            Cfg.TitleText = "LsgEMD_131 — Текущее количество рабочих мест, занятых на предприятии, ед."
            Cfg.LastCol = "D"
            Cfg.ValueCols = "B,C,D"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "E"
            Cfg.Headers = Array("LsgEMD_003 — ID проекта", _
                    "LsgEMD_132 — Год отчетного периода", _
                    "LsgEMD_133 — план", _
                    "LsgEMD_134 — факт")
            Cfg.Widths = Array(20, 17, 16, 16)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", "", "", "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 9      ' LsgEMD_135 - Количество созданных рабочих мест (сводка в столбце CB, там код LsgEMD_136)
            Cfg.GroupCode = "LsgEMD_135"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "CB"
            Cfg.ParentKeyCol = vbNullString
            Cfg.SheetName = "LsgEMD_135"
            Cfg.TitleText = "LsgEMD_135 — Количество созданных рабочих мест, ед."
            Cfg.LastCol = "F"
            Cfg.ValueCols = "B,C,D,E,F"
            Cfg.KeyCol = vbNullString
            Cfg.CtxCol = "G"
            Cfg.Headers = Array("LsgEMD_003 — ID проекта", _
                    "LsgEMD_136 — Год отчетного периода", _
                    "LsgEMD_137 — план", _
                    "LsgEMD_138 — факт", _
                    "LsgEMD_139 — план (кумулятивно за все года)", _
                    "LsgEMD_140 — факт (кумулятивно за все года)")
            Cfg.Widths = Array(20, 17, 16, 16, 21, 21)
            Cfg.IntCols = ""
            Cfg.MoneyCols = ""
            Cfg.DictCodes = Array("", "", "", "", "", "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 10     ' LsgEMD_141 - Уплаченные налоги (сводка в столбце CC, там код LsgEMD_142)
            Cfg.GroupCode = "LsgEMD_141"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "CC"
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
            Cfg.DictCodes = Array("", "", "", "", "", "")
            Cfg.ChildGroupCode = vbNullString
            Cfg.ChildSummaryCol = vbNullString

        Case 11     ' LsgEMD_147 - Встречные обязательства (сводка в столбце CD, там код LsgEMD_148)
            Cfg.GroupCode = "LsgEMD_147"
            Cfg.ParentSheetName = vbNullString
            Cfg.ParentCol = "CD"
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
            Cfg.DictCodes = Array("", "", "", "")
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

    For i = 1 To GroupCount()
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

    For i = 1 To GroupCount()

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

    ' Группа узнается по столбцу-сводке ParentCol или по коду группы
    ' в строке кодов (регистр не важен: на листе Лизинга "LSGEMD_031")
    For i = 1 To GroupCount()

        LoadGroupCfg i, Cfg

        If Len(Cfg.ParentSheetName) = 0 Then
            If Target.Column = ws.Range(Cfg.ParentCol & "1").Column Or _
               StrComp(Cfg.GroupCode, Code, vbTextCompare) = 0 Then
                Group_OpenFromMain = True
                OpenGroup Cfg, ws, Target
                Exit Function
            End If
        End If

    Next i

End Function

Public Function Group_OpenFromParent(ByVal ws As Worksheet, ByVal Target As Range) As Boolean

    Dim i As Long
    Dim Cfg As TGroup

    If Target.Row < FIRST_DATA_ROW Then Exit Function

    For i = 1 To GroupCount()

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

    For i = 1 To GroupCount()

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

    For i = 1 To GroupCount()

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

    For i = 1 To GroupCount()

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
