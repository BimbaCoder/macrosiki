Attribute VB_Name = "ModuleLSG1"
Option Explicit

'==================================================================
' modChildTables - служебные листы-формы (подтаблицы проекта).
' Пустая оболочка - в этом файле пока нет отдельных "дочерних форм"
' (как AKKEMD_024/035 в тест1) вне системы множественных атрибутов.
' Если понадобится - добавьте Case в LoadCfg по аналогии с тест1
' и увеличьте CFG_COUNT.
'==================================================================

Private Const CFG_COUNT As Long = 0
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
    ' пусто
End Sub

Public Function Child_OpenFromMain( _
    ByVal AttrCode As String, _
    ByVal Target As Range) As Boolean
    Dim i As Long: Dim Cfg As TChild
    If Len(AttrCode) = 0 Then Exit Function
    If CFG_COUNT = 0 Then Exit Function
    For i = 1 To CFG_COUNT
        LoadCfg i, Cfg
        If InStr(1, ";" & Cfg.TriggerCodes & ";", ";" & AttrCode & ";", vbTextCompare) > 0 Then
            Child_OpenFromMain = True
            Exit Function
        End If
    Next i
End Function

Public Function Child_HandleClick( _
    ByVal Sh As Worksheet, _
    ByVal Target As Range) As Boolean
    Dim i As Long: Dim Cfg As TChild
    For i = 1 To CFG_COUNT
        LoadCfg i, Cfg
        If StrComp(Sh.Name, Cfg.SheetName, vbTextCompare) = 0 Then Exit Function
    Next i
End Function
