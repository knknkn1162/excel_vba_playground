Option Explicit

Sub main()
    Dim root As String, str1 As String, str2 As String
    Dim cnt1 As Integer, cnt2 As Integer
    str1 = "Book_20201101.xlsx"
    str2 = "Book_20201102.xlsx"
    Dim cur_ws As Worksheet
    Set cur_ws = Worksheets(1)

    root = ThisWorkbook.Path
    Dim ws As Worksheet
    Dim i As Integer
    i = 1
    Application.ScreenUpdating = False
    With Workbooks.Open(root & "/ex023/" & str1)
        cnt1 = .Worksheets.Count
        For each ws In .Worksheets
            cur_ws.Cells(i, 1) = ws.Name
            i = i + 1
        Next
        .Close SaveChanges:=False
        Application.ScreenUpdating = True
    End With

    i = 1
    Application.ScreenUpdating = False
    With Workbooks.Open(root & "/ex023/" & str2)
        cnt2 = .Worksheets.Count
        For each ws In .Worksheets
            cur_ws.Cells(i, 2) = ws.Name
            i = i + 1
        Next
        .Close SaveChanges:=False
        Application.ScreenUpdating = True
    End With
    If cnt1 <> cnt2 Then
        Msgbox "不一致"
        Exit Sub
    End If

    Dim rng As Range
    Set rng = cur_ws.Range("A1").CurrentRegion
    rng.Columns("A").sort key1:=Range("A1") 
    rng.Columns("B").sort key1:=Range("B1")
    Dim flagStr As String
    flagStr = "一致"
    For i = 1 To cnt1
        If Cells(i, 1) <> Cells(i, 2) Then flagStr = "不一致"
    Next
    
    Msgbox flagStr & " cnt: " & cnt1
    rng.ClearContents

End Sub

Sub SetAppConfig(ByVal b As Boolean)
    application.screenupdating = b
End Sub

Sub main2()
    Dim root As String: root = ThisWorkbook.Path
    Dim str1 As String: str1 = "Book_20201101.xlsx"
    Dim str2 As String: str2 = "Book_20201102.xlsx"

    Call SetAppConfig(False)
    Set wb2 = Workbooks.Open(root & "/ex023/" & str2)
    Set wb1 = Workbooks.Open(root & "/ex023/" & str1)

    ' If count is not match, not equal
    If wb1.Sheets.Count <> wb2.Sheets.Count Then
        GoTo PrintNotEq
    End If
    ' Assume that two workbook has the same number of sheets
    Dim sht As Worksheet: For sht In wb2.Sheets
        On Error Resume Next
        Dim sht2 As Object: Set sht2 = wb1.Sheets(sht.Name)
        Err.Clear
        If sht2 Is Nothing Then GoTo PrintNotEq
    Next

    Call SetAppConfig(True)
PrintEq:
    Msgbox("一致")
    Exit Sub
PrintNotEq:
    Msgbox "不一致"
End Sub
