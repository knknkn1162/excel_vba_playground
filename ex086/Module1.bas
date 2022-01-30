Option Explicit

Sub CreateFormat(st As Range, arr As Variant)
    Dim sz As Integer: sz = UBound(arr) - LBound(arr) + 1
    st.Offset(,1).Resize(1, sz) = arr
    st.Offset(1).Resize(sz) = WorksheetFunction.transpose(arr)
    With st.Resize(sz+1, sz+1)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    Dim i As Integer
    For i = 1 To sz
        With st.Offset(i,i)
            .Borders(xlDiagonalDown).LineStyle = xlContinuous
            .ClearContents
        End With
    Next
End Sub

Sub main()
    Dim sz As Integer: sz = Worksheets.Count
    Dim arr() As String
    Redim arr(1 To sz)
    Dim i As Integer
    For i = 1 To sz
        arr(i) = Worksheets(i).Name
    Next

    Dim tws As Worksheet
    WorkSheets.Add Before:=Worksheets(1)
    Set tws = ActiveSheet
    tws.Name = "相関表"
    Dim trng As Range: Set trng = tws.Range("B2")
    Call CreateFormat(trng, arr)
End Sub
