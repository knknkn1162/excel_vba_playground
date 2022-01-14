Option Explicit

Sub main()
    Dim arr(1 To 999) As Integer
    Dim rng As Range
    Dim cols As Integer
    Set rng = Range("A1").CurrentRegion
    cols = rng.Columns.Count
    Dim str As String
    Dim i As Integer
    Dim pos As String
    For i = 1 To cols
        str = Cells(1, i)
        pos = InstrRev(str, "(")
        Dim idx As Integer
        idx = Val(Mid(str, pos+1))
        arr(idx) = i
    Next
    pos = 1
    For i = 1 To 999
        If arr(i) = 0 Then
            GoTo Continue
        End If
        Cells.Columns(pos+cols).Value = rng.Columns(arr(i)).Value
        pos = pos + 1
Continue:
    Next
    rng.EntireColumn.Delete
End Sub
