Option Explicit

Function ReplaceSeriesFormula(formula As String, pos As Integer) As String
    Dim arr() As String
    ' get the first cell in ranges
    arr = Split(formula, ",")

    Dim idx As Integer: idx = UBound(arr)-pos
    Dim str As String: str = arr(idx)
    Dim r As Range: Set r = Range(str).Cells(1,1)
    ' replace arg(idx) with new one
    arr(idx) = Range(r, r.End(xlDown)).Address(External:=True)
    ReplaceSeriesFormula = Join(arr, ",")
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
