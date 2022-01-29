Option Explicit

Sub main()
    Dim pos As Integer
    Dim master_ws As WorkSheet, dst_ws As WorkSheet
    Set master_ws = Worksheets("社員")
    Set dst_ws = Worksheets("部・課マスタ")
    Dim rng As Range
    With master_ws.Range("A1").CurrentRegion
        Set rng = Intersect(.Cells, .Offset(, 2))
    End With
    ' Copy
    rng.Copy Destination:= dst_ws.Range("A1")

    With dst_ws.Range("A1").CurrentRegion
        Set rng = Intersect(.Cells, .Offset(1))
    End With
    ' sort
    rng = WorkSheetFunction.sort(rng, 2)

    Dim prev As String
    prev = "xxxx"
    ' uniq
    Dim i As Integer
    For i = Cells(Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If Cells(i, 2) = prev Then
            Rows(i).Delete
        Else
            prev = Cells(i,2)
        End If
    Next
End Sub
