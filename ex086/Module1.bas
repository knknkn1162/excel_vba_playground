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

Function ParseDecendents(str As String, ByRef arr As Variant) As Variant
    Dim i As Integer
    Dim ret() As Boolean
    ReDim ret(LBound(arr) To UBound(arr))
    For i = LBound(arr) To UBound(arr)
        ret(i) = False
        Dim pat As String: pat = "*'" & arr(i) & "'!*"
        If str Like pat Then ret(i) = True
    Next
    ParseDecendents = ret
End Function

Sub UpdateArr(ByRef bs As Variant, tmp As Variant)
    Dim i As Integer
    For i = LBound(bs) To UBound(bs)
        bs(i) = tmp(i) Or bs(i)
    Next
End Sub

Function BoolToox(bs As Variant) As Variant
    Dim ret() As Variant
    ReDim ret(LBound(bs) To UBound(bs))
    Dim i As Integer
    For i = LBound(bs) To UBound(bs)
        ret(i) = IIF(bs(i), "〇", "")
    Next
    BoolToox = ret
End Function

Sub main()
    Dim sz As Integer: sz = Worksheets.Count
    Dim origs() As String
    Dim arr() As String
    Redim orig(1 To sz)
    Redim arr(1 To sz)
    Dim i As Integer, j As Integer
    For i = 1 To sz
        orig(i) = Worksheets(i).Name
        arr(i) = orig(i) & vbTab
        Worksheets(i).Name = arr(i)
    Next

    Dim tws As Worksheet
    WorkSheets.Add Before:=Worksheets(1)
    Set tws = ActiveSheet
    tws.Name = "相関表"
    Dim trng As Range: Set trng = tws.Range("B2")

    Dim r As Range
    For i = 1 To sz
        Dim ws As Worksheet: Set ws = Worksheets(arr(i))
        Dim rng As Range
        Dim bs() As Boolean
        ReDim bs(1 To sz)
        On Error Resume Next
        Set rng = ws.Cells.SpecialCells(XlCellTypeFormulas)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo Continue
        End If
        On Error GoTo 0
        For Each r In rng
            Dim tmp() As Boolean
            tmp = parseDecendents(r.Formula, arr)
            Call UpdateArr(bs, tmp)
        Next
Continue:
        trng.Offset(i,1).Resize(1,sz).Value() = BoolToox(bs)
    Next
    For i = 1 To sz
        Worksheets(arr(i)).Name = orig(i)
    Next
    Call CreateFormat(trng, orig)
End Sub
