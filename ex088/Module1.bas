Option Explicit

Sub rankABC(rng As Range, b As Range, sum As Long)
    Dim ws As Worksheet: Set ws = rng.Worksheet
    rng.Sort key1:=b, Order1:=xlDescending, Header:=xlYes
    Dim i As Integer
    Dim cum As Long: cum = 0
    Dim rnk As String
    For i = 2 To rng.Rows.Count
        cum = cum + ws.Cells(i, b.Column)
        Select Case (cum / sum)
            Case Is <= 0.5
                rnk = "A"
            Case Is <= 0.9
                rnk = "B"
            Case Else
                rnk = "C"
        End Select
        ws.Cells(i, b.Column+2) = rnk
    Next
End Sub

Sub main()
    Dim ws As Worksheet: Set ws = Worksheets("data")
    Dim mws As Worksheet: Set mws = Worksheets("商品マスタ")
    Dim tws As Worksheet: Set tws = Worksheets("クロスABC")

    Dim rng As Range: Set rng = ws.Range("A1").CurrentRegion
    Set rng = Intersect(rng, rng.Offset(1))
    Dim r As Range
    Dim pos As Integer: pos = 2
    ' Fill A~E
    For each r In rng.Columns(1).Cells
        Dim code As String: code = r.Value()
        Dim idx As Integer
        idx = WorksheetFunction.match(code, mws.Range("A1").CurrentRegion.Columns(1), 0)
        tws.Cells(pos, 1).Resize(1,5) = Array(code, mws.Cells(idx,2), r.Offset(,1), _
            mws.Cells(idx, 3), mws.Cells(idx,4))
        pos = pos + 1
    Next

    ' Fill F,G,H(calculation
    Set rng = tws.Range("A1").CurrentRegion
    Set rng = Intersect(rng, rng.Offset(1))
    rng.Columns("F") = "=C2*D2"
    rng.Columns("G") = "=C2*E2"
    rng.Columns("H") = "=G2-F2"

    ' Fill J and I(sort)
    Set rng = tws.Range("A1").CurrentRegion
    Call rankABC(rng, tws.Range("H1"), WorksheetFunction.Sum(rng.Columns("H")))
    Call rankABC(rng, tws.Range("G1"), WorksheetFunction.Sum(rng.Columns("G")))
End Sub
