Option Explicit

Sub setApp(ByVal arg As Boolean)
　　With Application
　　　　.Calculation = IIf(arg, xlCalculationAutomatic, xlCalculationManual)
　　　　.DisplayAlerts = arg
　　　　.ScreenUpdating = arg
　　End With
End Sub

Sub main()
    Dim rng As Range
    Call setApp(False)

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
    Call setApp(True)

End Sub
