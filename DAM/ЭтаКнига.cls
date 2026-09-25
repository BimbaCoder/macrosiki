Option Explicit

Private WithEvents App As Application

Private Sub Workbook_Open()

    Set App = Application
    Application.OnDoubleClick = vbNullString
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
