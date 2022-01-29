Option Explicit

Sub main()
    With Range("A1").CurrentRegion
        Intersect(.Cells, .Offset(1,1)).ClearContents
    End With
End Sub
