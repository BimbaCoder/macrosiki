Attribute VB_Name = "ModuleAKK4"
Option Explicit

'==================================================================
' modMultiAttr - листы СТАРЫХ (плоских, без сводной ячейки)
' множественных атрибутов. Пустая оболочка - все найденные группы
' в этом файле переехали сразу в Module5 (система сводной ячейки).
' Оставлено, чтобы Module3 (Router) компилировался без изменений,
' если понадобится - можно добавить Case по аналогии со старыми
' файлами (КЖК Module4, версия до перехода на Module5).
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

Private Sub LoadCfg(ByVal Index As Long, ByRef Cfg As TMultiAttr)
    ' пусто
End Sub

Public Function MultiAttr_OpenFromMain( _
    ByVal ws As Worksheet, _
    ByVal Target As Range) As Boolean
    ' пусто - все группы в Module5
End Function

Public Function MultiAttr_HandleClick( _
    ByVal Sh As Worksheet, _
    ByVal Target As Range) As Boolean
    ' пусто - все группы в Module5
End Function

Public Function IsMultiAttrSheet(ByVal SheetName As String) As Boolean
    ' пусто - все группы в Module5
End Function
