Option Explicit

Sub main3()
    Dim rng As Range
    Dim ws As Worksheet: Set ws = ActiveSheet
    Set rng = Range("A1").CurrentRegion
    Set rng = Intersect(rng, rng.Offset(1))
    Dim cols As Integer: cols = rng.Columns.Count
    ws.AutoFilterMode = False
    ' calc condition
    rng.Columns(cols+1) =  "=OR(COUNTIF(B2:F2, ""<50"") = 0, SUM(B2:F2) >= 350)"
    Range("A1").AutoFilter Field:=(cols+1), Criteria1:=True
    rng.Columns(7).SpecialCells(xlCellTypeVisible) = "合格"
    ws.AutoFilterMode = False
    rng.Columns(cols+1).ClearContents
End Sub
