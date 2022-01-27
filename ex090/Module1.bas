Option Explicit
Function IsOverlapped(top1 As Double, bottom1 As Double, top2 As Double, bottom2 As Double) As Boolean
    IsOverlapped  = top1 < bottom2 And top2 < bottom1
End Function


Function CountOverLapped(rng As Range, isDeleted As Boolean)
    Dim sp As Shape
    Dim ret As Long: ret = 0
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Range
    For Each sp In ws.Shapes
        If sp.Type <> msoPicture Then GoTo Continue1
        For Each r In rng.Areas
            If Not IsOverlapped( _
                sp.Top, sp.Top+sp.Height, r.Top, r.Top+r.Height) Then GoTo Continue2
            If Not IsOverlapped( _
                sp.Left, sp.Left+sp.Width, r.Left, r.Left+r.Width) Then GoTo Continue2
            If isDeleted Then sp.Delete
            ret = ret + 1
            Exit For
Continue2:
        Next
Continue1:
    Next
    CountOverLapped = ret
End Function

Sub main()
    Range("A3") = CountOverLapped(Range("A1"), True)
    Range("A4") = CountOverLapped(Range("B3:F10"), False)
    Range("A5") = CountOverLapped(Range("B3,C6,E8"), False)
    Range("A6") = CountOverLapped(Range("B5:C7"), True)
    Range("A7") = CountOverLapped(Range("E4"), True)
    Range("A8") = CountOverLapped(Range("B3:F10"), True)
End Sub
