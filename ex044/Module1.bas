Option Explicit

Function CreateWorksheet() As Worksheet
    Dim fname As String
    fname = "status_" & Format(Now(), "hhmmss")
    Dim res_ws As Worksheet
    Set res_ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    With res_ws
        .Name = fname
        .Range("A1").resize(,5) = _
            Array("テーブル名", "シート名", "セル範囲", "リスト行数", "リスト列数")
    End With
    Set CreateWorksheet = res_ws
End Function

Sub main()
    Dim ws As Worksheet, tws As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Set tws = CreateWorksheet()
    Dim pos As Integer: pos = 2
    For Each ws In Worksheets
        If ws.Name = tws.Name Then
            GoTo Continue
        End If
        Dim lst As ListObject
        For Each lst In ws.ListObjects
            Dim rng As Range: Set rng = lst.DataBodyRange
            tws.Cells(pos,1).Resize(,5) = _
                Array(lst.Name, ws.Name, _
                    rng.Address, _
                    rng.Rows.Count, _
                    rng.Columns.Count _
                )
            pos = pos + 1
        Next
Continue:
    Next
    tws.UsedRange.EntireColumn.AutoFit
End Sub
