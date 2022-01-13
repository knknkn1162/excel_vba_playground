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
