Option Explicit

Sub main()
    Dim ws As Worksheet, ws2 As Worksheet
    Dim i As Integer
    Dim r As Range
    For Each ws In Worksheets
        Dim origName As String: origName = ws.Name
        ' change name that requires single quortation
        ws.Name = ws.Name & " "
        Dim str As String: str = "'" & Replace(ws.Name,"'","''") & "'!"
        Dim rng As Range

        On Error Resume Next
        Set rng = ws.Cells.SpecialCells(xlCellTypeFormulas)
        If rng Is Nothing Then GoTo Continue
        Err.Clear

        For Each r In rng
            ' .Formula2 property considers spill
            Dim prev As String: prev = r.Formula2
            Dim cur
            cur = Replace(r.Formula2, str, "")
            Debug.Print prev & " -> " & r.Formula2
            If r.HasArray Then
                r.FormulaArray = cur
            Else
                r.Formula2 = cur
            End If
Continue:
        Next
        ws.Name = origName
    Next
End Sub
