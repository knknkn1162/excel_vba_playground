Option Explicit

Function GetRGB(v As Long) As String
    Dim myR As Long: myR = v Mod 256
    Dim myG As Long: myG = Int(v / 256) Mod 256
    Dim myB As Long: myB = Int(v / 256 / 256)
    GetRGB = "#" & Right("0" & Hex(myR), 2) & _
        Right("0" & Hex(myG), 2) & _
        Right("0" & Hex(myB), 2)
End Function

Function PrintRGB(rng As Range, tp As Integer) As Variant
    Dim ret() As String
    ReDim ret(1 To rng.Rows.Count, 1 To rng.Columns.Count)
    Dim i As Integer, j As Integer
    For i = 1 To rng.Rows.Count
        For j = 1 To rng.Columns.Count
            Dim r As Range: Set r = rng.Item(i,j)
            Dim v As Long
            Select Case tp
                Case 1
                    v = r.Interior.Color
                Case 2
                    v = r.Font.Color
            End Select
            ret(i, j) = GetRGB(v)
        Next
    Next
    PrintRGB = ret
End Function

Sub main()
    Range("A5") = PrintRGB(Range("A1:B3"), 1)

    Range("A8:B10") = PrintRGB(Range("A1:B3"), 1)
    Range("A11:B13") = PrintRGB(Range("A1:B3"), 2)
End Sub
