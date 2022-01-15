Option Explicit

Sub main()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim maxVal As Integer, minVal As Integer
    Dim i As Integer
    Dim cht As Chart
    Set cht = ws.ChartObjects(1).Chart

    cht.SetSourceData Source:=Range("A1").CurrentRegion
    With cht.SeriesCollection(1)
        ' We should consider the posibility that
        ' two or more max candidates appear
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
