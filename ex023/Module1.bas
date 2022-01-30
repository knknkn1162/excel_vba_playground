Option Explicit

Type AppConfig
    Calculation As Boolean
    DisplayAlerts As Boolean
    ScreenUpdating As Boolean
End Type

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

Sub main()
    Dim root As String: root = ThisWorkbook.Path
    Dim str1 As String: str1 = "Book_20201101.xlsx"
    Dim str2 As String: str2 = "Book_20201102.xlsx"

    Dim conf As AppConfig, orig As AppConfig
    Dim ret As String
    conf.Calculation = false
    conf.DisplayAlerts = false
    conf.ScreenUpdating = false
    orig = SetAppConfig(conf)
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
        On Error GoTo 0
        If sht2 Is Nothing Then GoTo PrintNotEq
    Next

PrintEq:
    ret = "一致"
    GoTo Release
PrintNotEq:
    ret = "不一致"
Release:
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=False
    Msgbox ret
    Call SetAppConfig(orig)
End Sub
