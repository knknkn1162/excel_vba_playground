Option Explicit

Sub main()
    Dim ws As WorkSheet
    Dim rng As Range
    For Each ws In WorkSheets
        With ws.Cells
            On Error Resume Next
            Set rng = InterSect(.Cells, .SpecialCells(xlCellTypeFormulas))
            Err.Clear
        End With
        If rng Is Nothing Then
            Exit For
        End If
        Dim r As Range
        For each r In rng.Areas 
            r.Value = r.Value
        Next
    Next

    ' including Graph
    For Each ws In Sheets
        If Instr(ws.Name, "社外秘") <> 0 Then
            Application.DisplayAlerts = false
            ws.Visible = xlSheetVisible
            ws.Delete
            Application.DisplayAlerts = true
        End If
    Next

End Sub
