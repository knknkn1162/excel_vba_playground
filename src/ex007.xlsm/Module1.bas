Option Explicit

Sub main()
    Dim i As Integer
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        Dim str As String
        str = Cells(i, 1)
        str = Replace(str, ".", "/")
        If isDate(str) Then
            Dim d As Date
            d = DateValue(str)
            Cells(i, 2) = Format(DateSerial(Year(d), Month(d)+1, 0), "'mmdd")
        Else
            Cells(i, 2) = ""
        End If
    Next
End Sub
