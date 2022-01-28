Option Explicit

Function CompareArr(v1 As Variant, v2 As Variant) As Boolean
    CompareArr = False
    Dim i As Integer, j As Integer
    For i = LBound(v1) To UBound(v1)
        For j = LBound(v2) To UBound(v2)
            If v1(i) = "" Or v2(j) = "" Then GoTo Skip1
            If v1(i) = v2(j) Then CompareArr = True
Skip1:
        Next
    Next
End Function

Function IsMove(oldr As Range, newr As Range) As Boolean
    IsMove = True
    If oldr.Row = newr.Row Or oldr.Column = newr.Column Then IsMove = False
    If CompareArr( _
        Array(oldr.Offset(1,0).Value(), oldr.Offset(-1,0).Value(), oldr.Offset(0,1).Value(), oldr.Offset(0, -1)), _
        Array(newr.Offset(1,0).Value(), newr.Offset(-1,0).Value(), newr.Offset(0,1).Value(), newr.Offset(0, -1)) _
    ) Then IsMove = False
End Function

Sub main()
    Dim ows As Worksheet: Set ows = Worksheets("座席表（現）")
    Dim nws As Worksheet: Set nws = Worksheets("座席表（新）")

    Dim orng As Range: Set orng = ows.Range("B5:G10")
    Dim nrng As Range: Set nrng = nws.Range("B5:G10")

    Dim i As Integer, j As Integer
    For i = 1 To orng.Rows.Count
        For j = 1 To orng.Columns.Count
            Dim or1 As Range: Set or1 = orng.Cells(i,j)
            Dim nr1 As Range: Set nr1 = nrng.Find(What:=or1.Value())
            If Not IsMove(or1, nr1) Then
                nr1.Interior.Color = vbYellow
            End If
        Next
    Next

End Sub
