Option Explicit

Function ReplaceSeriesFormula(formula As String, pos As Integer) As String
    Dim arr() As String
    ' get the first cell in ranges
    arr = Split(formula, ",")

    Dim str As String: str = arr(UBound(arr)-pos)
    Dim r As Range: Set r = Range(str).Cells(1,1)
    Dim ret As String: ret = formula
    ret = Replace(ret, _
        str, _
        Range(r, r.End(xlDown)).Address(External:=True) _
    )
    ReplaceSeriesFormula = ret
End Function

Sub main()
    Dim ws As Worksheet
    Dim i As Integer, j As Integer
    For Each ws In Worksheets
        For i = 1 To ws.ChartObjects.Count
            Dim ser As Series
            For Each ser In ws.ChartObjects(i).Chart.SeriesCollection
                ser.Formula = ReplaceSeriesFormula(ser.Formula, 1)
                ser.Formula = ReplaceSeriesFormula(ser.Formula, 2)
            Next
        Next
    Next
End Sub
