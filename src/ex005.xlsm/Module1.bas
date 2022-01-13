Option Explicit

Sub main()
    Dim i As Integer, sr As Integer, sc As Integer
    Dim srng As Range
    Set srng = Range("B2")
    sr = srng.row + 1
    sc = srng.Column
    Columns(sc+2).NumberFormatLocal = "\#,##0"
    For i = sr To srng.currentRegion.Rows.Count + sr - 2
        if Cells(i, sc) = "" Or Cells(i, sc+1) = "" Then
            GoTo Continue
        End If
        Cells(i, sc+2) = Cells(i, sc) * Cells(i, sc+1)
Continue:
    Next
End Sub
