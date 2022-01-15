Option Explicit

Sub main()
    Dim rng As Range
    Set rng = Range("A1").CurrentRegion

    Dim ews(1) As Worksheet
    Dim pos(1) As Integer
    Set ews(0) = Worksheets("平日"): pos(0) = 2
    Set ews(1) = Worksheets("土日祝"): pos(1) = 2
    ews(0).Cells(1,1).Value = rng.Rows(1).Value
    ews(1).Cells(1,1).Value = rng.Rows(1).Value
    
    Dim filter_ws As Worksheet
    Set filter_ws = Worksheets("祝日")
    Dim i As Integer
    For i = 2 To rng.Rows.Count
        Dim idx As Integer
        idx = 0
        Select Case Weekday(Cells(i,1))
            Case vbSunday,vbSaturday
                idx = 1
        End Select
        Dim n As Integer
        n = 0
        On Error Resume Next
        n = WorksheetFunction.match(Cells(i,1), _
            filter_ws.Range("A1").CurrentRegion.Columns("A"), _
            0)
        Err.Clear
        ' not match
        If n <> 0 Then
            idx = 1
        End If
        rng.Rows(i).Copy Destination:=ews(idx).Cells(pos(idx),1)
        pos(idx) = pos(idx) + 1
    Next
End Sub

' use autofilter
Sub main2()
    Dim rng As Range
    Set rng = Range("A1").CurrentRegion

    Dim ews(1) As Worksheet
    Dim pos(1) As Integer
    Set ews(0) = Worksheets("平日"): pos(0) = 2
    Set ews(1) = Worksheets("土日祝"): pos(1) = 2
    Dim ws As Worksheet
    Set ws = Worksheets("売上")

    ' filter
    Dim filter_col As Integer
    filter_col = rng.Columns.Count+1
    rng.Columns(filter_col) = "=IF(OR(WEEKDAY(A1,2)>=6, COUNTIF(祝日!A:A,売上!A1)>0),1,0)"
    ws.AutoFilterMode = false
    Dim i As Integer
    For i = 0 To 1
        ws.Range("A1").AutoFilter Field:=filter_col, Criteria1:=i
        rng.SpecialCells(xlCellTypeVisible).Copy Destination:=ews(i).Range("A1")
    Next
    ws.AutoFilterMode = false
    rng.Columns(filter_col).ClearContents

End Sub
