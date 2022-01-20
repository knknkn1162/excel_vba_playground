Option Explicit

Sub main()
    Dim data_dir As String
    Dim sheetName As String
    sheetName = "2020年12月"
    Dim this_ws As Worksheet
    Set this_ws = Worksheets(sheetName)
    Dim pos As Long
    pos = this_ws.Range("A1").CurrentRegion.Rows.Count+1
    data_dir = ThisWorkbook.Path & "/" & "ex040_data"
    Dim fname As String
    fname = Dir(data_dir & "/*.xls")
    Dim wb As Workbook
    Dim ws As Worksheet

    Application.ScreenUpdating = false
    Do While fname <> ""
        On Error Resume Next
        Set wb = Workbooks.Open(data_dir & "/" & fname)
        If Err.Number <> 0 Then
            Debug.Print("could not open " & fname)
            GoTo Continue
        End If

        On Error Resume Next
        Set ws = wb.Worksheets(sheetName)
        If Err.Number <> 0 Then
            GoTo CloseWorkbook
        End If
        Dim rng As Range
        Set rng = ws.Range("A1").CurrentRegion
        Set rng = Intersect(rng, rng.Offset(1))
        rng.Copy Destination:=this_ws.Cells(pos, 1)
        pos = pos + rng.Rows.Count
CloseWorkbook:
        wb.Close saveChanges:=False
Continue:
        fname = Dir()
    Loop
    Application.ScreenUpdating = true
End Sub
