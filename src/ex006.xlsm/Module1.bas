Option Explicit

Sub main()
    Dim i As Integer
    For i = 2 To Cells(Rows.Count, 1).end(xlUp).Row
        If Instr(Cells(i,1), "-") = 0 Then
            ' formula
            'Cells(i, 4) = "=" & Cells(i,2).Address(False, False) & "*" & Cells(i,3).Address(False, False)
            Cells(i, 4).FormulaR1C1 = "=RC[-2]*RC[-1]"
        End If
    Next
End Sub
