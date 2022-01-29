Option Explicit

Sub main3()
    Dim master_ws As WorkSheet, dst_ws As WorkSheet
    Set master_ws = Worksheets("社員")
    Set dst_ws = Worksheets("部・課マスタ")
    dst_ws.Cells.Clear
    master_ws.Range("A1").CurrentRegion.Columns("C:F").Copy _
        Destination:=dst_ws.Range("A1")

    With dst_ws
        .Range("A1").CurrentRegion.RemoveDuplicates _
        Columns:=Array(1,2,3,4), Header:=xlYes
        .Range("A1").CurrentRegion.Sort _
            key1:= .Range("A1"), order1:=xlAscending, _
            key2:= .Range("B1"), order2:=xlAscending, _
            Header:=xlYes
    End With
End Sub
