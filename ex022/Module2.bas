Option Explicit

Sub main2()
    Dim i As Integer
    For i = 1 To 500
        Select Case True
            Case i Mod 15 = 0
                Cells(i, 4) = "FizzBuzz"
            Case i Mod 3 = 0
                Cells(i, 3) = "Buzz"
            Case i Mod 5 = 0
                Cells(i, 2) = "Buzz"
            Case Else
                Cells(i,1) = i
        End Select
    Next
End Sub
