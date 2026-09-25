Option Explicit

Private WithEvents App As Application

Private Sub Workbook_Open()

    Set App = Application

    ' сброс старого механизма (Application.OnDoubleClick)
    Application.OnDoubleClick = vbNullString

    ' UserInterfaceOnly-защита листов групп не сохраняется между
    ' сеансами Excel - переставляем ее при каждом открытии
    ReapplyGroupProtection

End Sub

Public Sub InitEvents()
    Set App = Application
    Application.OnDoubleClick = vbNullString
    ReapplyGroupProtection
End Sub

Private Sub App_SheetBeforeDoubleClick( _
    ByVal Sh As Object, _
    ByVal Target As Range, _
    ByRef Cancel As Boolean)

    Router_DoubleClick Sh, Target, Cancel

End Sub

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    Router_SheetChange Sh, Target

End Sub

' Workbook_BeforeSave (проверка обязательных полей перед сохранением)
' ВРЕМЕННО УБРАНА по просьбе - сохранять можно с незаполненными
' обязательными полями. Функция ValidateRequiredFields в Module3
' никуда не делась - когда понадобится снова включить проверку,
' достаточно вернуть сюда:
'
' Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'     If Not ValidateRequiredFields() Then Cancel = True
' End Sub
