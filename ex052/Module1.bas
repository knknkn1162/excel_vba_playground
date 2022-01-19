Option Explicit

Sub main()
    Dim ws As Worksheet
    Dim arr() As String
    Redim arr(Worksheets.Count)
    Dim i As Integer: i = 0
    For Each ws In Worksheets
        If ws.Visible <> xlSheetVisible Then GoTo Continue
        If ws.Name like "*印刷*" Then
            arr(i) = ws.Name
            i = i + 1
        End If
Continue:
    Next
    ReDim Preserve arr(WorksheetFunction.max(i-1,0))
    Sheets(arr).Select
    ActiveWindow.SelectedSheets.PrintOut Preview:=True
End Sub
