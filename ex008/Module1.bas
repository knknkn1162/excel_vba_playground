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
