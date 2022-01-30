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
