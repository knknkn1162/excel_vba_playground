Option Explicit

Sub main()
    Dim ws As Worksheet
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim root As String: root = wb.Path
    Dim dir As String: dir = root & "/ex059_wb"
    On Error Resume Next
    MkDir(dir)
    Err.Clear
    Dim arr(2) As String
    Dim i As Integer, j As Integer
    Dim baseDate As Date: baseDate = #2020/04/01#
    For i = 1 To 4
        Dim d As Date: d = DateAdd("m", (i-1)*3, baseDate)
        For j = 0 To 2
            arr(j) = Format(DateAdd("m", j, d), "yyyy年mm月")
        Next
        wb.Sheets(arr).copy
        ActiveWorkbook.SaveAs _
            FielName:=dir & "/" & i & "Q.xlsx", _
            FileFormat:=xlOpenXMLWorkbook
        ActiveWorkbook.Close saveChanges:=False
    Next
End Sub
