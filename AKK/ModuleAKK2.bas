Attribute VB_Name = "ModuleAKK2"
Option Explicit

'==================================================================
' modDictionaries - выбор значений из справочников (файл 40_*.xlsx)
' БЕЗ ИЗМЕНЕНИЙ в этой итерации (06) - перенесён как есть, включая
' фикс SPR039 (страны - берем название, не буквенный код) и
' Dict_OpenFromRow (обобщенная версия для листов-рыб).
'==================================================================

Public Const DICTIONARY_MASK As String = "40_*.xlsx"

Private Const PART_SEPARATOR As String = " / "
Private Const FILTER_BY_CURRENT_VALUE As Boolean = True

Private SelectedCell As Range
Private SelectedDictionarySheet As String
Private SelectedResultCols As Variant
Private SelectedFilterCol As String

Private Sub GetDictConfig( _
    ByVal DictCode As String, _
    ByVal AttrCode As String, _
    ByRef ResultCols As Variant, _
    ByRef FilterCol As String)

    Select Case DictCode

        Case "SPR001"
            ResultCols = Array("C", "D", "E")
            FilterCol = "D"

        Case "SPR004"
            Select Case AttrCode
                Case "AKKEMD_019"
                    ResultCols = Array("E")
                    FilterCol = "E"
                Case "AKKEMD_020"
                    ResultCols = Array("D")
                    FilterCol = "D"
                Case Else
                    ResultCols = Array("C", "D", "E")
                    FilterCol = "E"
            End Select

        Case "SPR022"
            ResultCols = Array("F", "G", "O", "P")
            FilterCol = "O"

        Case "SPR017"
            ResultCols = Array("C", "D")
            FilterCol = "D"

        Case "SPR039"                             ' Страны - брать название, не буквенный код
            ResultCols = Array("F")
            FilterCol = "F"

        Case "SPR049"
            ResultCols = Array("C", "D")
            FilterCol = "C"

        Case "SPR050"
            ResultCols = Array("C", "D", "E")
            FilterCol = "E"

        Case Else
            ResultCols = Array("C")
            FilterCol = "C"

    End Select

End Sub

Public Function Dict_OpenFromMain( _
    ByVal Sh As Worksheet, _
    ByVal Target As Range) As Boolean

    Dict_OpenFromMain = Dict_OpenFromRow(Sh, Target, DICT_CODE_ROW)

End Function

Public Function Dict_OpenFromRow( _
    ByVal Sh As Worksheet, _
    ByVal Target As Range, _
    ByVal CodeRow As Long) As Boolean

    Dim DictCode As String
    Dim AttrCode As String

    DictCode = UCase$(CellText(Sh.Cells(CodeRow, Target.Column)))
    If Not (DictCode Like "SPR###") Then Exit Function

    Dict_OpenFromRow = True

    If CodeRow = DICT_CODE_ROW Then
        AttrCode = CellText(Sh.Cells(ATTR_CODE_ROW, Target.Column))
    Else
        AttrCode = vbNullString
    End If

    OpenDictionary Target, DictCode, AttrCode

End Function

Private Sub OpenDictionary( _
    ByVal Target As Range, _
    ByVal DictCode As String, _
    ByVal AttrCode As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim NumRow As Long
    Dim LastRow As Long
    Dim LastCol As Long
    Dim ResultCols As Variant
    Dim FilterCol As String
    Dim FilterColNum As Long
    Dim FilterIndex As Long
    Dim Parts As Variant
    Dim FilterValue As String
    Dim FoundCell As Range

    Set wb = GetDictionaryWorkbook()
    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    Set ws = wb.Worksheets(DictCode)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "В файле """ & wb.Name & """ нет листа " & DictCode & ".", _
               vbExclamation
        Exit Sub
    End If

    ShowOnlySprSheet wb, ws

    If Not GetLayout(ws, NumRow, LastRow, LastCol) Then
        MsgBox "Не удалось определить структуру листа " & DictCode & _
               " (нет строки нумерации столбцов 1, 2, 3... или нет данных).", _
               vbExclamation
        Exit Sub
    End If

    GetDictConfig DictCode, AttrCode, ResultCols, FilterCol
    FilterColNum = ws.Range(FilterCol & "1").Column
    FilterIndex = IndexOfColumn(ResultCols, FilterCol)
    If FilterColNum > LastCol Then LastCol = FilterColNum

    Set SelectedCell = Target
    SelectedDictionarySheet = ws.Name
    SelectedResultCols = ResultCols
    SelectedFilterCol = FilterCol

    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    On Error GoTo 0

    If FILTER_BY_CURRENT_VALUE And FilterIndex >= 0 Then

        If Len(CStr(Target.Value2)) > 0 Then

            Parts = Split(CStr(Target.Value2), PART_SEPARATOR)

            If UBound(Parts) >= FilterIndex Then

                FilterValue = Trim$(CStr(Parts(FilterIndex)))

                If Len(FilterValue) > 0 Then

                    ws.Range(ws.Cells(NumRow, 1), ws.Cells(LastRow, LastCol)) _
                        .AutoFilter Field:=FilterColNum, Criteria1:=FilterValue

                    On Error Resume Next
                    Set FoundCell = ws.Range( _
                        ws.Cells(NumRow + 1, FilterColNum), _
                        ws.Cells(LastRow, FilterColNum)) _
                        .SpecialCells(xlCellTypeVisible).Cells(1, 1)
                    On Error GoTo 0

                    If FoundCell Is Nothing Then
                        On Error Resume Next
                        If ws.FilterMode Then ws.ShowAllData
                        On Error GoTo 0
                    End If

                End If

            End If

        End If

    End If

    wb.Activate
    ws.Activate

    If FoundCell Is Nothing Then
        Application.GoTo ws.Cells(NumRow + 1, FilterColNum), True
    Else
        Application.GoTo FoundCell, True
    End If

End Sub

Public Sub Dict_SelectFromDictionary( _
    ByVal Sh As Worksheet, _
    ByVal Target As Range, _
    ByRef Cancel As Boolean)

    Dim NumRow As Long
    Dim LastRow As Long
    Dim LastCol As Long
    Dim RowNumber As Long
    Dim FilterColNum As Long
    Dim i As Long
    Dim ResultValue As String
    Dim FilterValue As String
    Dim Check As String

    If SelectedCell Is Nothing Then Exit Sub

    On Error Resume Next
    Check = SelectedCell.Address
    If Err.Number <> 0 Then
        Err.Clear
        Set SelectedCell = Nothing
        Exit Sub
    End If
    On Error GoTo 0

    If StrComp(Sh.Name, SelectedDictionarySheet, vbTextCompare) <> 0 Then Exit Sub
    If Not GetLayout(Sh, NumRow, LastRow, LastCol) Then Exit Sub

    RowNumber = Target.Row
    If RowNumber <= NumRow Then Exit Sub
    If RowNumber > LastRow Then Exit Sub

    Cancel = True

    For i = LBound(SelectedResultCols) To UBound(SelectedResultCols)

        If i > LBound(SelectedResultCols) Then
            ResultValue = ResultValue & PART_SEPARATOR
        End If

        ResultValue = ResultValue & _
            CellText(Sh.Cells(RowNumber, CStr(SelectedResultCols(i))))

    Next i

    FilterColNum = Sh.Range(SelectedFilterCol & "1").Column
    FilterValue = CellText(Sh.Cells(RowNumber, FilterColNum))
    If FilterColNum > LastCol Then LastCol = FilterColNum

    Application.EnableEvents = False
    SelectedCell.Value2 = ResultValue
    Application.EnableEvents = True

    On Error Resume Next
    If Sh.FilterMode Then Sh.ShowAllData
    On Error GoTo 0

    If Len(FilterValue) > 0 Then
        Sh.Range(Sh.Cells(NumRow, 1), Sh.Cells(LastRow, LastCol)) _
            .AutoFilter Field:=FilterColNum, Criteria1:=FilterValue
    End If

    ThisWorkbook.Activate
    SelectedCell.Parent.Activate
    SelectedCell.Select

    Set SelectedCell = Nothing

End Sub

Public Function Dict_IsDictionaryBook(ByVal wb As Workbook) As Boolean
    Dict_IsDictionaryBook = (Not (wb Is ThisWorkbook)) And _
                            (LCase$(wb.Name) Like LCase$(DICTIONARY_MASK))
End Function

Private Function GetDictionaryWorkbook() As Workbook

    Dim wb As Workbook
    Dim FileName As String
    Dim FilePath As String

    For Each wb In Application.Workbooks
        If Dict_IsDictionaryBook(wb) Then
            Set GetDictionaryWorkbook = wb
            Exit Function
        End If
    Next wb

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Сначала сохраните текущий файл, чтобы найти папку со справочниками.", _
               vbExclamation
        Exit Function
    End If

    FileName = Dir$(ThisWorkbook.Path & Application.PathSeparator & DICTIONARY_MASK)

    If FileName = vbNullString Then
        MsgBox "Файл справочников " & DICTIONARY_MASK & " не найден в папке:" & _
               vbCrLf & ThisWorkbook.Path, vbExclamation
        Exit Function
    End If

    FilePath = ThisWorkbook.Path & Application.PathSeparator & FileName

    Set GetDictionaryWorkbook = Application.Workbooks.Open( _
        FileName:=FilePath, _
        UpdateLinks:=0, _
        ReadOnly:=True, _
        AddToMru:=False)

    ProtectDictionarySheets GetDictionaryWorkbook

End Function

Private Function GetLayout( _
    ByVal ws As Worksheet, _
    ByRef NumRow As Long, _
    ByRef LastRow As Long, _
    ByRef LastCol As Long) As Boolean

    Dim r As Long

    NumRow = 0
    LastRow = 0
    LastCol = 0

    For r = 1 To 20
        If CellText(ws.Cells(r, 1)) = "1" And _
           CellText(ws.Cells(r, 2)) = "2" And _
           CellText(ws.Cells(r, 3)) = "3" Then
            NumRow = r
            Exit For
        End If
    Next r

    If NumRow = 0 Then Exit Function

    LastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    LastCol = ws.Cells(NumRow, ws.Columns.Count).End(xlToLeft).Column

    GetLayout = (LastRow > NumRow)

End Function

Private Function IndexOfColumn( _
    ByVal Cols As Variant, _
    ByVal Letter As String) As Long

    Dim i As Long

    IndexOfColumn = -1

    For i = LBound(Cols) To UBound(Cols)
        If StrComp(CStr(Cols(i)), Letter, vbTextCompare) = 0 Then
            IndexOfColumn = i - LBound(Cols)
            Exit Function
        End If
    Next i

End Function

Private Function CellText(ByVal Cell As Range) As String

    Dim v As Variant
    Dim s As String

    v = Cell.Value2
    If IsError(v) Then Exit Function

    s = CStr(v)
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, ChrW(160), " ")

    CellText = Trim$(s)

End Function

Private Sub ShowOnlySprSheet(ByVal wb As Workbook, ByVal wsKeep As Worksheet)

    Dim ws As Worksheet

    For Each ws In wb.Worksheets

        If ws.Name Like "SPR###" Then

            If StrComp(ws.Name, wsKeep.Name, vbTextCompare) = 0 Then
                If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
            Else
                If ws.Visible <> xlSheetHidden Then ws.Visible = xlSheetHidden
            End If

        End If

    Next ws

End Sub

Private Sub ProtectDictionarySheets(ByVal wb As Workbook)

    Dim ws As Worksheet

    For Each ws In wb.Worksheets

        On Error Resume Next
        ws.Protect DrawingObjects:=True, _
                   Contents:=True, _
                   Scenarios:=True, _
                   UserInterFaceOnly:=True, _
                   AllowFiltering:=True
        On Error GoTo 0

    Next ws

End Sub
