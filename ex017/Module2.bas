Option Explicit

Sub main2()
    Dim master_ws As WorkSheet, dst_ws As WorkSheet
    Set master_ws = Worksheets("社員")
    Set dst_ws = Worksheets("部・課マスタ")
    dst_ws.Cells.Clear
    master_ws.Columns("C:F").AdvancedFilter Action:=xlFilterCopy, _
        CopyToRange:=dst_ws.Range("A1"), _
        Unique:=True

    With dst_ws
        .Range("A1").CurrentRegion.Sort _
            key1:= .Range("A1"), order1:=xlAscending, _
            key2:= .Range("B1"), order2:=xlAscending, _
            Header:=xlYes
    End With
    
End Sub
