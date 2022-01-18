Option Explicit

Sub main()
    Dim i As Integer
    For i = 2 To Cells(Rows.Count, 1).end(xlUp).Row
        If Instr(Cells(i,1), "-") = 0 Then
            ' formula
            Cells(i, 4).FormulaR1C1 = "=RC[-2]*RC[-1]"
        End If
    Next
End Sub

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
