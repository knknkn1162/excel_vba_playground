Option Explicit

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
