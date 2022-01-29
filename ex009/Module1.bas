Option Explicit

Sub main()
    Dim ws As Worksheet
    Set ws = Worksheets("成績表")
    Application.DisplayAlerts = False
    On Error Resume Next
    WorkSheets("合格者").Delete
    Err.Clear
    On Error GoTo 0
    Application.DisplayAlerts = True

    Dim dst_ws As Worksheet
    Set dst_ws = Worksheets.Add(After:=ws)
    dst_ws.Name = "合格者"

    ws.AutoFilterMode = False
    With ws.Range("A1").CurrentRegion
        .AutoFilter Field:=7, Criteria1:="合格"
        .Columns(1).Copy dst_ws.Range("A1")
    End With
    ws.AutoFilterMode = False
End Sub
