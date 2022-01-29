Option Explicit

Sub main2()
    Dim rng As Range: Set rng = Range("A1").CurrentRegion
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim srng As Range

    ws.AutoFilterMode = False
    ws.Range("A1").AutoFilter Field:=1, Criteria1:="<>*-*"
    Set srng = Intersect(rng.Offset(1,3), rng.SpecialCells(xlCellTypeVisible))
    srng.FormulaR1C1 = "=RC[-2]*RC[-1]"
    ws.AutoFilterMode = False
End Sub
