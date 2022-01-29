Option Explicit

Sub main2()
    Dim ws As Worksheet
    Set ws = WorkSheets(1)
    ws.AutoFilterMode = false
    Dim rng As Range
    With ws.Range("A1").CurrentRegion
        .AutoFilter field:=3, Criteria1:=""
        .AutoFilter field:=4, Criteria1:="*削除*", Operator:=xlOr, Criteria2:="*不要*"
        Set rng = Intersect(.Offset(1), .SpecialCells(xlCellTypeVisible))
        If Not rng Is Nothing Then rng.EntireRow.Delete
    End With
    ws.AutoFilterMode = false
End Sub
