Option Explicit

Sub main()
    Dim i As Integer
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        Dim rng As Range
        Set rng = Range(Cells(i, 2), Cells(i, 6))
        If WorksheetFunction.Countif(rng, "<50") > 0 Then
            Cells(i, 7) = ""
        ElseIf WorksheetFunction.Sum(rng) < 350 Then
            Cells(i, 7) = ""
        else
            Cells(i, 7) = "合格"
        End If
    Next
End Sub

Sub main2()
    Dim rng As Range
    Set rng = Range("A1").CurrentRegion
    Set rng = Intersect(rng, rng.Offset(1,1))

    Dim r As Range
    For each r In rng.rows
        Dim arg1 As Range
        Set arg1 = r.resize(,5)
        With WorksheetFunction
            If .Sum(arg1) >= 300 And _
                .CountIf(arg1, "<50") = 0 Then
                r.Columns(6) = "合格"
            End If
        End With
    Next
End Sub
