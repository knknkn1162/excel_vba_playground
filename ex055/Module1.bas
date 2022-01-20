Option Explicit

Sub main()
    Dim root As String: root = Thisworkbook.Path
    Dim fpath As String: fpath = root & "/ex055/test.xlsm"
    Application.EnableEvents = False
    Dim wb As Workbook: Set wb = Workbooks.Open(fpath)
    Dim i As Integer
    i = Application.Run("'" & Replace(fpath,"'","''") & "'" & "!mult",3,5)
    wb.Close SaveChanges:=False
    Range("A1") = i
    Application.EnableEvents = True
End Sub
