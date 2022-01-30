Option Explicit

Type AppConfig
    EnableEvents As Boolean
    ScreenUpdating As Boolean
End Type

Function setAppConfig(conf As AppConfig) As AppConfig
    Dim orig As AppConfig
　　With Application
        orig.EnableEvents = .EnableEvents
        orig.ScreenUpdating = .ScreenUpdating
　　　　.ScreenUpdating = conf.ScreenUpdating
        .EnableEvents = conf.EnableEvents
　　End With
    setAppConfig = orig
End Function

Sub main()
    Dim root As String: root = Thisworkbook.Path
    Dim fpath As String: fpath = root & "/ex055/test.xlsm"
    Dim conf As AppConfig
    conf.ScreenUpdating = False: conf.EnableEvents = False
    Dim orig As AppConfig: orig = SetAppConfig(conf)

    Dim wb As Workbook
    Set wb = Workbooks.Open(Filename:=fpath, ReadOnly:=True)
    Dim i As Integer
    ' 他ブックのマクロを起動するには[Application.]Runを使います。
    i = Application.Run("'" & Replace(fpath,"'","''") & "'" & "!mult",3,5)
    wb.Close SaveChanges:=False
    Range("A1") = i
    Call SetAppConfig(orig)
End Sub
