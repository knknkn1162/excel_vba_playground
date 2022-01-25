Option Explicit

Sub main()
    Dim mws As Worksheet: Set mws = Worksheets("マスタ")
    Dim rng As Range: Set rng = mws.Range("A1").CurrentRegion.Columns("A")
    Dim dws As Worksheet: Set dws = Worksheets("data")
    Dim i As Integer
    For i = 2 To dws.Cells(Rows.Count,1).End(xlUp).Row
        Dim pos As Integer
        On Error Resume Next
        pos = WorksheetFunction.match(dws.Cells(i,1), rng, 0)
        If Err.Number <> 0 Then
            dws.Cells(i,1).Font.Color = vbRed
            Err.Clear
            GoTo Continue
        End If
        dws.Cells(i,1).Phonetic.Text = rng.Rows(pos).Phonetic.Text
Continue:
    Next

End Sub
