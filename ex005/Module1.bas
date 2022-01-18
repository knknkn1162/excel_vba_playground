Option Explicit

Sub main()
    Dim i As Integer, sr As Integer, sc As Integer
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
