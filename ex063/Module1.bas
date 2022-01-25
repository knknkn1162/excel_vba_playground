Option Explicit

Sub main()
    Dim ws As Worksheet
    Dim tws As Worksheet
    Dim ws1 As Worksheet: Set ws1 = Worksheets(1)
    Dim header As Range: Set header = ws1.Range("A1").CurrentRegion.Rows(1)
    Set tws = Worksheets.Add(Before:=Worksheets(1))
    With tws
        .Name = "summary"
        .Range("A1").Resize(1,header.Columns.Count) = header.Value()
    End With
    Dim pos As Integer: pos = 2
    For each ws In Worksheets
        If Not ws.Name like "####年##月" Then GoTo Continue
        Dim rng As Range
        Set rng = ws.Range("A1").CurrentRegion
        Set rng = Intersect(rng, rng.Offset(1))
        rng.Copy Destination:=tws.Cells(pos,1)
        pos = pos + rng.Rows.Count
Continue:
    Next
End Sub
