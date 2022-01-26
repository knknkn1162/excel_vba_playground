Option Explicit

Sub main()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim origRng As Range: Set origRng = ActiveCell
    ' テーブル範囲に選択セルがある場合
    ' シートのオートフィルター操作が正しく行われないので範囲外のセルをselect
    ws.Cells(ws.Rows.Count, ws.Columns.Count).Select
    ' シートのフィルターは1つ
    If ws.AutoFilterMode Then ws.AutoFilter.ShowAllData

    ' テーブルのフィルターはテーブルごとに存在
    Dim tbl As ListObject
    For Each tbl In ws.ListObjects
        If Not tbl.AutoFilter Is Nothing Then
            tbl.AutoFilter.ShowAllData
        End If
    Next
    origRng.Select
End Sub
