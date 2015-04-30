'
' NI SE MO NO CSV Editor
' Hidehiro Kozawa
'

Option Explicit

Sub NiSeMoNoCSVEditor()

'
    Dim AERCount As Integer
    Dim SheetExistFlag As Boolean
    SheetExistFlag = False
    AERCount = ActiveSheet.Protection.AllowEditRanges.Count
    
    For i = 1 To AERCount
        If ActiveSheet.Protection.AllowEditRanges(i).Title = "NiSeCSV" Then
            SheetExistFlag = True
            
            GoTo AfterAERFindLoop
        End If
    Next
AfterAERFindLoop:

    Cells.Select
    If SheetExistFlag = False Then
        Selection.NumberFormatLocal = "@"
        ActiveSheet.Protection.AllowEditRanges.Add Title:="NiSeCSV", Range:=Cells
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If
End Sub

Sub workbook_open()

    NiSeMoNoCSVEditor

End Sub

Sub SaveCSV()

    Dim FileName As Variant

    FileName = _
        Application.GetSaveAsFilename( _
             InitialFileName:=ThisWorkbook.Name _
           , FileFilter:="CSVファイル(*.csv),*.csv" _
           , FilterIndex:=1 _
           , Title:="保存先の指定" _
           )
End Sub
