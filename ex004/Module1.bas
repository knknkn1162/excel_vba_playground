Option Explicit

Sub main()
    With Range("A1").CurrentRegion.Offset(1,1)
        On Error Resume Next
        ' SpecialCellsは該当せるが存在しない場合はエラーとなる
        .SpecialCells(xlCellTypeConstants).ClearContents
        Err.Clear
    End With
End Sub
