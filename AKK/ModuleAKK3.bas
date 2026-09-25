Attribute VB_Name = "ModuleAKK3"
Option Explicit

'==================================================================
' modRouter - единая точка входа для двойного клика и для изменений
' ячеек (нужно для автостирания записи листов вложенных групп).
' Сгенерировано по аналогии с файлом КЖК_ЭФ.
'==================================================================

Public Const MAIN_SHEET_NAME As String = "АКК СХ_ЭФ"
Public Const ATTR_CODE_ROW As Long = 8
Public Const DICT_CODE_ROW As Long = 9
Public Const FIRST_INPUT_ROW As Long = 10
Public Const PROJECT_ID_COL As String = "C"
Public Const FIRST_DATA_ROW_CHILD As Long = 5

'------------------------------------------------------------------
' Двойной клик. Порядок проверок:
'   1) служебные листы (кнопки A2 / B2) - дочерние формы (Module1)
'   1.4) кнопки на листах вложенных групп (Module5)
'   1.45) переход по вложенной группе (сводная ячейка внутри листа группы)
'   1.5) кнопки на листах старых множественных атрибутов (Module4, пусто)
'   1.55) справочники внутри листов-рыб (новых и старых)
'   2) основной лист: сначала вложенные группы, потом старые
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

        If Child_HandleClick(ws, Cell) Then
            Cancel = True
            Exit Sub
        End If

        If Group_HandleClick(ws, Cell) Then
            Cancel = True
            Exit Sub
        End If

        If Group_OpenFromParent(ws, Cell) Then
            Cancel = True
            Exit Sub
        End If

        If MultiAttr_HandleClick(ws, Cell) Then
            Cancel = True
            Exit Sub
        End If

        If IsGroupSheet(ws.Name) Or IsMultiAttrSheet(ws.Name) Then
            If Cell.Row >= FIRST_DATA_ROW_CHILD Then
                If Dict_OpenFromRow(ws, Cell, 3) Then Cancel = True
            End If
            Exit Sub
        End If

        If ws.Name <> MAIN_SHEET_NAME Then Exit Sub
        If Cell.Row < FIRST_INPUT_ROW Then Exit Sub

        If Group_OpenFromMain(ws, Cell) Then
            Cancel = True
            Exit Sub
        End If

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

        Dict_SelectFromDictionary ws, Cell, Cancel

    End If

    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    MsgBox "Ошибка: " & Err.Description, vbExclamation

End Sub

Public Sub Router_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    On Error GoTo ErrorHandler

    If TypeName(Sh) <> "Worksheet" Then Exit Sub

    NestedGroup_HandleChange Sh, Target

    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    MsgBox "Ошибка обработки изменения: " & Err.Description, vbExclamation

End Sub

Public Sub Router_Init()
    ThisWorkbook.InitEvents
    MsgBox "Обработка двойного клика и изменений включена.", vbInformation
End Sub
