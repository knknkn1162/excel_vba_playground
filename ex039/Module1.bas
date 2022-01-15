Option Explicit

Sub main()
    Dim pos As Integer
    pos = 1
    Dim i As Integer, j As Integer
    i = 1: j = 1
    Dim row1 As Integer, row2 As Integer
    row1 = Cells(Rows.Count,1).End(xlUp).Row
    row2 = Cells(Rows.Count,2).End(xlUp).Row
    const IntegerMax As Integer = 32767
    ' guard
    Cells(row1+1, 1) = IntegerMax
    Cells(row2+1, 2) = IntegerMax
    ' shakutori method
    For i = i To row1+1
        For j = j To row2+1
            If Cells(i,1) <= Cells(j,2) Then
                If Cells(i,1) = Cells(j,2) Then
                    j = j+1
                End If
                Exit For
            End If
            Cells(pos, 3) = Cells(j, 2)
            pos = pos + 1
        Next
        if Cells(i, 1) <> IntegerMax Then
            Cells(pos, 3) = Cells(i,1)
            pos = pos + 1
        End If
    Next
    Cells(row1+1,1) = ""
    Cells(row2+1,2)= ""
End Sub
