Option Explicit

Type AppConfig
    Calculation As Boolean
    DisplayAlerts As Boolean
    ScreenUpdating As Boolean
End Type

Function switchAppConfig(b As Boolean) As AppConfig
    Dim orig As AppConfig
　　With Application
        orig.Calculation = .Calculation
        orig.DisplayAlerts = .DisplayAlerts
        orig.ScreenUpdating = .ScreenUpdating
　　　　.Calculation = IIf(b, xlCalculationAutomatic, xlCalculationManual)
　　　　.DisplayAlerts = b
　　　　.ScreenUpdating = b
　　End With
    switchAppConfig = orig
End Function

Function setAppConfig(conf As AppConfig) As AppConfig
    Dim orig As AppConfig
　　With Application
        orig.Calculation = .Calculation
        orig.DisplayAlerts = .DisplayAlerts
        orig.ScreenUpdating = .ScreenUpdating
　　　　.Calculation = IIf(conf.Calculation, xlCalculationAutomatic, xlCalculationManual)
　　　　.DisplayAlerts = conf.DisplayAlerts
　　　　.ScreenUpdating = conf.ScreenUpdating
　　End With
    setAppConfig = orig
End Function

Sub CreateDir(d As String)
    On Error Resume Next
    Mkdir d
    On Error GoTo 0
End Sub

Sub TestOpenBooks()
    Dim bdir As String: bdir = ThisWorkbook.Path & "/ex032"
    Dim fname As String
    fname = Dir(bdir & "/*.xls")
    Do While fname <> ""
        Workbooks.open(bdir & "/" & fname)
        fname = Dir()
    Loop
End Sub

Sub main()
    Dim orig As AppConfig
    orig = switchAppConfig(false)
    Call TestOpenBooks()
    Dim wb As Workbook
    Dim cur_wb As Workbook: Set cur_wb = ThisWorkbook
    Dim txtPath As String
    Dim bdir As String: bdir = cur_wb.Path & "/ex032_out"
    Call CreateDir(bdir)
    txtPath = bdir & "/log_" & Format(Now(), "yyyymmddhhmmss") & ".txt"
    Dim fnumber As Integer
    fnumber = FreeFile
    Open txtPath For Output As #fnumber
    For each wb In Workbooks
        ' wb.Save
        Print #fnumber, wb.Path & "/" & wb.Name
        If wb.Path <> cur_wb.Path Then wb.Close SaveChanges:=False
    Next
    Close #fnumber
    Call setAppConfig(orig)
    ' run command in vbac
    ' Application.Quit
End Sub
