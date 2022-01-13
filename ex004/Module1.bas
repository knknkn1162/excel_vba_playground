Option Explicit

Sub main()
    With Range("A1").CurrentRegion.Offset(1,1)
        ' .Value = .Value
        .Resize(.Rows.Count-2, .Columns.Count-2).clearContents
    End With
End Sub

Sub ans()
    On Error Resume Next
    With Range("A1").CurrentRegion.Offset(1,1)
        .SpecialCells(xlCellTypeConstants).ClearContents
    End With
End Sub
