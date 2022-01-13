Option Explicit

Sub main()
    Dim wb As Workbook
    Dim cur_wb As Workbook
    Set cur_wb = ThisWorkbook
    Dim txtPath As String
    txtPath = cur_wb.Path & "/" & "log_" & Format(Now(), "yyyymmddhhmmss") & ".txt"
    Dim fnumber As Integer
    fnumber = FreeFile
    Open txtPath For Output As #fnumber
    For each wb In Workbooks
        ' wb.Save
        Print #fnumber, cur_wb.Path & "/" & wb.Name
    Next
    Close #fnumber
    Application.Quit
End Sub
