Option Explicit

Sub main()
    Dim ws As Worksheet
    ' push as string
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
    If i <> 0 Then
        ReDim Preserve arr(i-1)
        ThisWorkbook.Sheets(arr).PrintOut Preview:=True
    End If
End Sub
