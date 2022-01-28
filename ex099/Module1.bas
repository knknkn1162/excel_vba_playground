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
    'Msgbox Join(v1, ",") & " vs " & Join(v2, ",") & "=>" & CompareArr
End Function

Function IsMove(oldr As Range, newr As Range) As Boolean
    IsMove = True
    'If oldr.Row = newr.Row Or oldr.Column = newr.Column Then IsMove = False
    If CompareArr( _
        Array(oldr.Offset(1,0).Value(), oldr.Offset(-1,0).Value(), oldr.Offset(0,1).Value(), oldr.Offset(0, -1)), _
        Array(newr.Offset(1,0).Value(), newr.Offset(-1,0).Value(), newr.Offset(0,1).Value(), newr.Offset(0, -1)) _
    ) Then IsMove = False
End Function

Function Shuffle(rows As Integer, columns As Integer) As Variant
    Dim ret() As Integer
    Dim flags() As Boolean
    Dim i As Integer, j As Integer, k As Integer
    Dim n1 As Integer, n2 As Integer
    Dim retry As Integer
    Do
        Redim ret(1 To rows*columns)
        Redim flags(1 to rows*columns)
        For i = 1 To rows
            For j = 1 To columns
                retry = 0
                Do
                    retry = retry + 1
                    n1 = WorksheetFunction.RandBetween(1, rows)
                    n2 = WorksheetFunction.RandBetween(1, columns)
                    if retry >= 200 GoTo Continue
                Loop While n1 = i Or n2 = j Or flags((n1-1)*columns+n2)
                flags((n1-1)*columns + n2) = True
                ret((i-1)*columns + j) = (n1-1)*columns + n2
            Next
        Next
        Exit Do
    Continue:
    Loop
    Shuffle = ret
End Function

Sub main()
    Dim ows As Worksheet: Set ows = Worksheets("座席表（現）")
    Dim nws As Worksheet: Set nws = Worksheets("座席表（新）")

    Dim orng As Range: Set orng = ows.Range("B5:G10")
    Dim nrng As Range: Set nrng = nws.Range("B5:G10")

    Dim sz As Integer: sz = orng.Rows.Count * orng.Columns.COunt
    Dim people() As String
    ReDim people(1 To sz)
    Dim i As Integer, j As Integer
    Dim pos As Integer
    For i = 1 To orng.Rows.Count
        For j = 1 To orng.Columns.Count
            pos = (i-1) * orng.Columns.Count + j
            people(pos) = orng.Cells(i,j)
        Next
    Next

    Dim retry As Integer
    Do While True
        nrng.Clear
        Dim ids() As Integer
        ids = Shuffle(orng.Rows.Count, orng.Columns.Count)
        Dim violatenum As Integer: violatenum = 0
        For i = 1 To orng.Rows.Count
            For j = 1 To orng.Columns.Count
                pos = (i-1) * orng.Columns.Count + j
                nrng.Cells(i, j) = people(ids(pos))
            Next
        Next
        ' check
        For i = 1 To orng.Rows.Count
            For j = 1 To orng.Columns.Count
                Dim or1 As Range: Set or1 = orng.Cells(i,j)
                Dim nr1 As Range: Set nr1 = nrng.Find(What:=or1.Value())
                If Not IsMove(or1, nr1) Then
                    violatenum = violatenum+1
                End If
            Next
        Next
        Debug.Print violatenum
        If violatenum = 0 Then Exit Do
        retry = retry + 1
    Loop
    Debug.Print "total: " & retry
End Sub
