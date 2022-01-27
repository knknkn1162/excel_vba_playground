Option Explicit
Function IsOverlapped(top1 As Long, bottom1 As Long, top2 As Long, bottom2 As Long) As Boolean
    IsOverlapped  = top1 < bottom2 And top2 < bottom1
End Function


Function CountOverLapped(rng As Range, isDeleted As Boolean)
    Dim sp As Shape
    Dim ret As Long: ret = 0
    Dim ws As Worksheet: Set ws = ActiveSheet
    For Each sp In ws.Shapes
        If sp.Type <> msoPicture Then GoTo Continue
        If Not IsOverlapped(sp.Top, sp.Top+sp.Height, rng.Top, rng.Top+rng.Height) Then GoTo Continue
        If Not IsOverlapped(sp.Left, sp.Left+sp.Width, rng.Left, rng.Left+rng.Width) Then GoTo Continue
        ret = ret + sp.Height * sp.Width
        If isDeleted Then sp.Delete
Continue:
    Next
    CountOverLapped = ret
End Function

Sub main()
    Range("A3") = CountOverLapped(Range("A1"), True)
    Range("A4") = CountOverLapped(Range("B3:F10"), False)
    Range("A5") = CountOverLapped(Range("B5:C7"), True)
    Range("A6") = CountOverLapped(Range("E4"), True)
    Range("A7") = CountOverLapped(Range("B3:F10"), True)
End Sub
