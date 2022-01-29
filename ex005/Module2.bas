Option Explicit

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
