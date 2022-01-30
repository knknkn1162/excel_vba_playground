Option Explicit

Const LastAuthor As Integer = 6

Sub main()
    Dim ws As WorkSheet: Set ws = ActiveSheet
    Dim root As String: root = ThisWorkbook.Path
    Dim fname As String
    fname = Dir(root & "/*.xls")
    Dim pos As Integer: pos = 2
    Do While fname <> ""
        Dim fpath As String: fpath = root & "/" & fname
        Dim orig As Boolean: orig = Application.ScreenUpdating
        Application.ScreenUpdating = False
        Dim wb As Workbook: Set wb = Workbooks.Open(fpath)
        Dim prop As DocumentProperties: Set prop = wb.BuiltinDocumentProperties
        On Error Resume Next
        Dim lastPrintDate As String: lastPrintDate = prop("Last print date")
        On Error GoTo 0
        If Val(lastPrintDate) = 0 Then lastPrintDate = ""
        ws.Cells(pos, 1).Resize(1,7) = Array( _
            wb.Name, _
            prop("Last Author"), _
            prop("Author"), _
            prop("Creation date"), _
            prop("Last Save Time"), _
            lastPrintDate, _
            FileLen(fpath) _
        )
        
        wb.Close SaveChanges:=False
        Application.ScreenUpdating = orig
        pos = pos +1
        fname = Dir()
        ' display up to 10
        If pos > 11 Then Exit Do
    Loop
End Sub
