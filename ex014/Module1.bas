Option Explicit

Type AppConfig
    Calculation As Boolean
    DisplayAlerts As Boolean
    ScreenUpdating As Boolean
End Type

Function setApp(conf As AppConfig) As AppConfig
    Dim orig As AppConfig
　　With Application
        conf.Calculation = .Calculation
        conf.DisplayAlerts = .DisplayAlerts
        conf.ScreenUpdating = .ScreenUpdating
　　　　.Calculation = IIf(conf.Calculation, xlCalculationAutomatic, xlCalculationManual)
　　　　.DisplayAlerts = conf.DisplayAlerts
　　　　.ScreenUpdating = conf.ScreenUpdating
　　End With
End Function

Sub main()
    Dim rng As Range
    Dim conf As AppConfig
    conf.Calculation = false
    conf.DisplayAlerts = false
    conf.ScreenUpdating = false
    Dim orig As AppConfig: orig = setApp(conf)

    Dim ws As Worksheet
    For Each ws In WorkSheets
        With ws.Cells
            On Error Resume Next
            Set rng = InterSect(.Cells, .SpecialCells(xlCellTypeFormulas))
            Err.Clear
        End With
        If rng Is Nothing Then Exit For
        Dim r As Range
        For each r In rng.Areas 
            r.Value = r.Value
        Next
    Next

    ' including Graph
    Dim sht As Object
    For Each sht In Sheets
        If Instr(sht.Name, "社外秘") <> 0 Then
            Application.DisplayAlerts = false
            sht.Visible = xlSheetVisible
            sht.Delete
            Application.DisplayAlerts = true
        End If
    Next
    Call setApp(orig)
End Sub
