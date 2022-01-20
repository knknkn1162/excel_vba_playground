Option Explicit

Sub SetAppConfig(ByVal b As Boolean)
    Application.ScreenUpdating = b
    ' Workbook_Openを止めるにはApplication.EnableEventsをFalseにします
    Application.EnableEvents = b
End Sub

Sub main()
    Dim root As String: root = Thisworkbook.Path
    Dim fpath As String: fpath = root & "/ex055/test.xlsm"
    Call SetAppConfig(False)
    Dim wb As Workbook
    Set wb = Workbooks.Open(Filename:=fpath, ReadOnly:=True)
    Dim i As Integer
    ' 他ブックのマクロを起動するには[Application.]Runを使います。
    i = Application.Run("'" & Replace(fpath,"'","''") & "'" & "!mult",3,5)
    wb.Close SaveChanges:=False
    Range("A1") = i
    Call SetAppConfig(True)
End Sub
