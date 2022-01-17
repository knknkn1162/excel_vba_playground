Option Explicit

Sub main()
    Dim ws As Worksheet
    Set ws = Activesheet
    With ws.ListObjects(1)
        .ListColumns.Add Position:=4
        .HeaderRowRange(4) = "合計列1"
        .DataBodyRange.Columns(4) = "=[@列1]+[@列2]+[@列3]"

        .ListColumns.Add
        Dim cols As Integer: cols = .HeaderRowRange.Count
        .HeaderRowRange(cols) = "合計列2"
        .DataBodyRange.Columns(cols)="=[@列4]+[@列5]"
    End With
End Sub

Sub main2()
    Dim ws As Worksheet: Set ws = Activesheet

    With ws.Range("B2").ListObject
        With .ListColumns.Add(.ListColumns("列3").Index+1)
            .Name = "合計列"
            .DataBodyRange.Value = "=sum([@[列1]:[列3]])"
        End With

        With .ListColumns.Add
            .Name = "合計列2"
            .DataBodyRange.Value = "=sum([@[列4]:[列5]])"
        End With
    End With
End Sub
