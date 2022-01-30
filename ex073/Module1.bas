Option Explicit

Public Sub Test()
    Msgbox "成功"
End Sub
Sub main()
    Dim wb0 As Workbook: Set wb0 = ThisWorkbook
    Dim wb As Workbook: Set wb = Workbooks.Add 
    Dim br As Range: Set br = Range("A1")
    ' ref: https://excel-ubara.com/excelvba1/EXCELVBA436.html
    With wb.Worksheets(1).Buttons.Add(br.Left, br.Top, br.Width, br.Height)
        .onAction = "Test"
        .Caption = "テスト"
    End With
    Dim orig As Boolean: orig = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wb.SaveAs wb0.Path & "/out.xlsx"
    wb.Close
    Application.DisplayAlerts = orig
End Sub
