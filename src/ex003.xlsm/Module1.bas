Option Explicit

Sub main()
    'Range("A1").currentRegion.offset(1,1).clearContents
    With Range("A1").CurrentRegion
        Intersect(.Cells, .Offset(1,1)).ClearContents
    End With
End Sub
