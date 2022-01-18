Option Explicit

Sub SetAppConfig(ByVal b As Boolean)
    application.screenupdating = b
End Sub

Sub main()
    Dim root As String: root = ThisWorkbook.Path
    Dim str1 As String: str1 = "Book_20201101.xlsx"
    Dim str2 As String: str2 = "Book_20201102.xlsx"

    Call SetAppConfig(False)
    Dim wb1 As Workbook: Set wb1 = Workbooks.Open(root & "/ex023/" & str1)
    Dim wb2 As Workbook: Set wb2 = Workbooks.Open(root & "/ex023/" & str2)

    ' If count is not match, not equal
    If wb1.Sheets.Count <> wb2.Sheets.Count Then
        GoTo PrintNotEq
    End If
    ' Assume that two workbook has the same number of sheets
    Dim sht As Object: For Each sht In wb2.Sheets
        On Error Resume Next
        Dim sht2 As Object: Set sht2 = wb1.Sheets(sht.Name)
        Err.Clear
        If sht2 Is Nothing Then GoTo PrintNotEq
    Next

PrintEq:
    Msgbox "一致"
    GoTo Release
PrintNotEq:
    Msgbox "不一致"
Release:
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=False
    Call SetAppConfig(True)
End Sub
