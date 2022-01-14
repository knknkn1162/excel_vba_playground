Option Explicit

Sub main()
    Dim rng As Range
    Set rng = Range("A1").CurrentRegion.Columns("B")
    Dim max As Long, min As Long
    max = WorksheetFunction.match(WorksheetFunction.max(rng), rng, 0) - 1
    min = WorksheetFunction.match(WorksheetFunction.min(rng), rng, 0) - 1

    Dim chartObj As ChartObject
    Set chartObj = ActiveSheet.ChartObjects(1)

    ' Reset
    With chartObj.Chart.SeriesCollection(1)
        .Interior.Color = RGB(68,114,196)
        .ApplyDataLabels xlDataLabelsShowNone
    End With

    With chartObj.Chart.SeriesCollection(1).Points(min)
        .Interior.Color = vbRed
        .ApplyDataLabels
    End With

    With chartObj.Chart.SeriesCollection(1).Points(max)
        .Interior.Color = vbGreen
        .ApplyDataLabels
    End With
End Sub

Sub main2()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim maxVal As Integer, minVal As Integer
    Dim i As Integer
    With ws.ChartObjects(1).Chart.SeriesCollection(1)
        maxVal = WorksheetFunction.max(.Values)
        minVal = WorksheetFunction.min(.Values)
        .Interior.Color = RGB(68, 114, 196)
　　　　.ApplyDataLabels xlDataLabelsShowNone
        For i = 1 To .Points.Count
            Dim arr() As Variant
            arr = .Values
            Select Case arr(i)
                Case maxVal
                    .Points(i).Interior.Color = vbGreen
                    .Points(i).ApplyDataLabels
                Case minVal
                    .Points(i).Interior.Color = vbRed
                    .Points(i).ApplyDataLabels
            End Select
        Next
    End With
End Sub
