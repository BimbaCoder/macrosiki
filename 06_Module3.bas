Attribute VB_Name = "Module3"
Option Explicit

'==================================================================
' modRouter - единая точка входа для двойного клика и для изменений
' ячеек (нужно для автостирания записи листов вложенных групп).
'==================================================================

Public Const MAIN_SHEET_NAME As String = "КЖК_ЭФ"
Public Const ATTR_CODE_ROW As Long = 8
Public Const DICT_CODE_ROW As Long = 9
Public Const FIRST_INPUT_ROW As Long = 10
Public Const PROJECT_ID_COL As String = "D"
Public Const FIRST_DATA_ROW_CHILD As Long = 5

'------------------------------------------------------------------
' Двойной клик. Порядок проверок:
'   1) служебные листы AKKEMD_024 / AKKEMD_035 (кнопки A2 / B2)
'   1.4) кнопки на листах вложенных групп (032 / 176)
'   1.45) переход по вложенной группе (176 внутри 032)
'   1.5) кнопки на листах старых множественных атрибутов (036,044,127,050)
'   1.55) справочники внутри листов вложенных групп и старых множ. атрибутов
'   2) основной лист: сначала вложенные группы (032), потом старые
'      множественные атрибуты, потом "дочерние формы", потом справочники
'   3) файл справочников 40_*.xlsx: выбор значения
'------------------------------------------------------------------
Public Sub Router_DoubleClick( _
    ByVal Sh As Object, _
    ByVal Target As Range, _
    ByRef Cancel As Boolean)

    Dim ws As Worksheet
    Dim Cell As Range
    Dim AttrCode As String

    On Error GoTo ErrorHandler

    If TypeName(Sh) <> "Worksheet" Then Exit Sub

    Set ws = Sh
    Set Cell = Target.Cells(1, 1)

    If ws.Parent Is ThisWorkbook Then

        ' 1) кнопки на служебных листах
        If Child_HandleClick(ws, Cell) Then
            Cancel = True
            Exit Sub
        End If

        ' 1.4) кнопки на листах вложенных групп (032/176)
        If Group_HandleClick(ws, Cell) Then
            Cancel = True
            Exit Sub
        End If

        ' 1.45) переход по вложенной группе (176 внутри 032)
        If Group_OpenFromParent(ws, Cell) Then
            Cancel = True
            Exit Sub
        End If

        ' 1.5) кнопки на листах старых множественных атрибутов
        If MultiAttr_HandleClick(ws, Cell) Then
            Cancel = True
            Exit Sub
        End If

        ' 1.55) справочники внутри листов-рыб (новых и старых)
        If IsGroupSheet(ws.Name) Or IsMultiAttrSheet(ws.Name) Then
            If Cell.Row >= FIRST_DATA_ROW_CHILD Then
                If Dict_OpenFromRow(ws, Cell, 3) Then Cancel = True
            End If
            Exit Sub
        End If

        ' 2) основной лист
        If ws.Name <> MAIN_SHEET_NAME Then Exit Sub
        If Cell.Row < FIRST_INPUT_ROW Then Exit Sub

        ' 2.05) новая система вложенных групп (032)
        If Group_OpenFromMain(ws, Cell) Then
            Cancel = True
            Exit Sub
        End If

        ' 2.1) старые множественные атрибуты (036,044,127,050)
        If MultiAttr_OpenFromMain(ws, Cell) Then
            Cancel = True
            Exit Sub
        End If

        AttrCode = Trim$(CStr(ws.Cells(ATTR_CODE_ROW, Cell.Column).Value2))

        If Child_OpenFromMain(AttrCode, Cell) Then
            Cancel = True
            Exit Sub
        End If

        If Dict_OpenFromMain(ws, Cell) Then Cancel = True

    ElseIf Dict_IsDictionaryBook(ws.Parent) Then

        ' 3) выбор значения в справочнике
        Dict_SelectFromDictionary ws, Cell, Cancel

    End If

    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    MsgBox "Ошибка: " & Err.Description, vbExclamation

End Sub

'------------------------------------------------------------------
' Изменение ячейки (нужно только для листов вложенных групп -
' автостирание записи, если очищены все содержательные поля,
' и автоподстановка ключа, если начали заполнять пустую строку)
'------------------------------------------------------------------
Public Sub Router_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    On Error GoTo ErrorHandler

    If TypeName(Sh) <> "Worksheet" Then Exit Sub

    NestedGroup_HandleChange Sh, Target

    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    MsgBox "Ошибка обработки изменения: " & Err.Description, vbExclamation

End Sub

' Подключить события вручную (если они "слетели" после ошибки/End)
Public Sub Router_Init()
    ThisWorkbook.InitEvents
    MsgBox "Обработка двойного клика и изменений включена.", vbInformation
End Sub

'------------------------------------------------------------------
' Обязательные поля листа "КЖК_ЭФ" (лист "КЖК", столбец R = "ДА")
' Собрано через двумерный массив построчным присвоением, а не одним
' большим Array(...) литералом - у VBA лимит 24 продолжения строки
' на один оператор, и 29 элементов через "_" в него не влезали.
'------------------------------------------------------------------
Private Function RequiredColumns() As Variant

    Dim Result(1 To 29, 1 To 1) As Variant
    Dim i As Long

    i = 1
    Result(i, 1) = Array("A", "KJKEMD_001 — Дата ввода"): i = i + 1
    Result(i, 1) = Array("B", "KJKEMD_002 — Оператор ввода"): i = i + 1
    Result(i, 1) = Array("D", "KJKEMD_003 — ID проекта"): i = i + 1
    Result(i, 1) = Array("F", "KJKEMD_004 — Учетный статус проекта"): i = i + 1
    Result(i, 1) = Array("G", "KJKEMD_005 — № проекта"): i = i + 1
    Result(i, 1) = Array("H", "KJKEMD_006 — Класс проекта"): i = i + 1
    Result(i, 1) = Array("I", "KJKEMD_007 — Базовая категория проекта"): i = i + 1
    Result(i, 1) = Array("J", "KJKEMD_009 — Наименование проекта"): i = i + 1
    Result(i, 1) = Array("K", "KJKEMD_010 — Код проекта"): i = i + 1
    Result(i, 1) = Array("L", "KJKEMD_011 — Уполномоченный госорган"): i = i + 1
    Result(i, 1) = Array("M", "KJKEMD_012 — Область реализации проекта"): i = i + 1
    Result(i, 1) = Array("N", "KJKEMD_014 — Город"): i = i + 1
    Result(i, 1) = Array("O", "KJKEMD_015 — Район"): i = i + 1
    Result(i, 1) = Array("P", "KJKEMD_016 — Вид экономической деятельности (ОКЭД, 2 знака)"): i = i + 1
    Result(i, 1) = Array("Q", "KJKEMD_017 — ОКЭД по проекту (5 знаков)"): i = i + 1
    Result(i, 1) = Array("R", "KJKEMD_018 — Тип проекта согласно решения УО"): i = i + 1
    Result(i, 1) = Array("S", "KJKEMD_019 — Кол-во домов, ед."): i = i + 1
    Result(i, 1) = Array("T", "KJKEMD_020 — Категория жилья"): i = i + 1
    Result(i, 1) = Array("U", "KJKEMD_021 — Площадь, кв.м."): i = i + 1
    Result(i, 1) = Array("V", "KJKEMD_022 — Количество квартир, ед."): i = i + 1
    Result(i, 1) = Array("W", "KJKEMD_023 — Классификация проекта"): i = i + 1
    Result(i, 1) = Array("X", "KJKEMD_024 — Дата ввода жилья"): i = i + 1
    Result(i, 1) = Array("Y", "KJKEMD_026 — Наименование заемщика"): i = i + 1
    Result(i, 1) = Array("Z", "KJKEMD_027 — Резидент/Нерезидент"): i = i + 1
    Result(i, 1) = Array("AA", "KJKEMD_028 — БИН заемщика проекта"): i = i + 1
    Result(i, 1) = Array("AB", "KJKEMD_029 — Код страны для нерезидента"): i = i + 1
    Result(i, 1) = Array("AC", "KJKEMD_030 — Организационно-правовая форма заемщика"): i = i + 1
    Result(i, 1) = Array("AD", "KJKEMD_031 — Размерность предприятия-заемщика"): i = i + 1
    Result(i, 1) = Array("BB", "KJKEMD_089 — Статус заявки/проекта")

    RequiredColumns = Result

End Function

Public Function ValidateRequiredFields() As Boolean

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long
    Dim cols As Variant
    Dim i As Long
    Dim ColLetter As String
    Dim ProjectID As String
    Dim Missing As String
    Dim FirstMissingCol As String

    ValidateRequiredFields = True

    Set ws = ThisWorkbook.Worksheets(MAIN_SHEET_NAME)
    cols = RequiredColumns()

    LastRow = ws.Cells(ws.Rows.Count, PROJECT_ID_COL).End(xlUp).Row
    If LastRow < FIRST_INPUT_ROW Then Exit Function

    For r = FIRST_INPUT_ROW To LastRow

        ProjectID = Trim$(CStr(ws.Cells(r, PROJECT_ID_COL).Value2))

        If Len(ProjectID) > 0 Then

            Missing = vbNullString
            FirstMissingCol = vbNullString

            For i = LBound(cols, 1) To UBound(cols, 1)
                ColLetter = CStr(cols(i, 1)(0))

                If Len(Trim$(CStr(ws.Range(ColLetter & r).Value2))) = 0 Then
                    Missing = Missing & "  - " & CStr(cols(i, 1)(1)) & vbCrLf
                    If Len(FirstMissingCol) = 0 Then FirstMissingCol = ColLetter
                End If
            Next i

            If Len(Missing) > 0 Then

                MsgBox "Проект """ & ProjectID & """ (строка " & r & _
                       ") - не заполнены обязательные поля:" & vbCrLf & Missing, _
                       vbExclamation, "Сохранение отменено"

                ws.Activate
                ws.Range(FirstMissingCol & r).Select

                ValidateRequiredFields = False
                Exit Function

            End If

        End If

    Next r

End Function

' Диагностика
Public Sub TestDictionaryState()

    Dim FileName As String

    FileName = Dir$(ThisWorkbook.Path & Application.PathSeparator & DICTIONARY_MASK)

    MsgBox _
        "ACTIVE BOOK: " & ActiveWorkbook.Name & vbCrLf & _
        "THIS BOOK: " & ThisWorkbook.Name & vbCrLf & _
        "SHEET: " & ActiveSheet.Name & vbCrLf & _
        "CELL: " & ActiveCell.Address(False, False) & vbCrLf & _
        "ATTRIBUTE: " & CStr(ActiveSheet.Cells( _
            ATTR_CODE_ROW, ActiveCell.Column).Value2) & vbCrLf & _
        "SPR CODE (row " & DICT_CODE_ROW & "): " & CStr(ActiveSheet.Cells( _
            DICT_CODE_ROW, ActiveCell.Column).Value2) & vbCrLf & _
        "PATH: " & ThisWorkbook.Path & vbCrLf & _
        "FILE: " & FileName

End Sub
