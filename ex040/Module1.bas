Option Explicit

Function setAppConfig(conf As Boolean) As Boolean
    Dim orig As Boolean
　　With Application
        orig = .ScreenUpdating
　　　　.ScreenUpdating = conf
　　End With
    setAppConfig = orig
End Function

Function OpenWorksheet(wb As Workbook, n As String) As Worksheet
    On Error Resume Next
    Set OpenWorksheet = wb.Worksheets(n)
    On Error GoTo 0
End Function

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
    Do While fname <> ""
        Dim orig As Boolean: orig = setAppConfig(False)
        On Error Resume Next
        Dim wb As Workbook: Set wb = Workbooks.Open(data_dir & "/" & fname)
        On Error GoTo 0
        If wb Is Nothing Then
            Debug.Print("could not open " & fname)
            GoTo Continue
        End If

        Dim ws As Worksheet: Set ws = OpenWorksheet(wb, sheetName)
        If ws Is Nothing Then
            Debug.Print("not found " & SheetName)
            GoTo CloseWorkbook
        End If
        Dim rng As Range
        Set rng = ws.Range("A1").CurrentRegion
        Set rng = Intersect(rng, rng.Offset(1))
        rng.Copy Destination:=this_ws.Cells(pos, 1)
        pos = pos + rng.Rows.Count
CloseWorkbook:
        wb.Close saveChanges:=False
        Call setAppConfig(orig)
Continue:
        fname = Dir()
    Loop
End Sub
