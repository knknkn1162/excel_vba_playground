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
            If r.HasSpill Or r.HasArray Then
                GoTo Continue
            End If
            Dim prev As String: prev = r.Formula
            r.Formula = Replace(r.Formula, str, "")
            Debug.Print prev & " -> " & r.Formula
Continue:
        Next
        ws.Name = origName
    Next
End Sub
