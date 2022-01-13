Option Explicit
Sub main()
    Dim rng As Range
    Set rng = Range("A1").CurrentRegion
    Set rng = Intersect(rng, rng.Offset(1))
    With rng
        .Columns("D") = "=IFERROR(VLOOKUP(B2,マスタ!A:C,2,FALSE),"""")"
        .Columns("E") = "=IFERROR(VLOOKUP(B2,マスタ!A:C,3,FALSE),"""")"
        .Columns("F") = "=C2*E2"
        .Columns("D:F").Value = .Columns("D:F").Value
    End With
End Sub

