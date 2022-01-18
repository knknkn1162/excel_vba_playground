Option Explicit

Sub main()
    Dim srng As Range
    Set srng = Range("B2").CurrentRegion
    ' specify body of Column:B
    Set srng = Intersect(srng, srng.Offset(1).Resize(,1))
    srng.Columns(3).NumberFormatLocal = "\#,##0"
    Dim r As Range
    For Each r In srng
        If r.Value = "" Or r.Offset(,1).Value = "" Then
            GoTo Continue
        End If
        r.Offset(,2) = r.Value * r.Offset(,1).Value
Continue:
    Next
End Sub

Sub main2()
    Dim i As Integer
    Dim srng As Range: Set srng = Range("B2").CurrentRegion
    srng.Columns(3).NumberFormatLocal = "\#,##0"
    Dim rng3 As Range: Set rng3 = Intersect(srng.Offset(1,2),srng)
    ' calc all
    rng3.FormulaR1C1 = "=RC[-2] * RC[-1]"
    ' clear value if not regular
    Intersect(srng.Columns(3), _
        srng.SpecialCells(xlCellTypeBlanks).EntireRow _
    ).ClearContents
End Sub
