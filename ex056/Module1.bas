Option Explicit

Sub main()
    Dim ws As Worksheet, ws2 As Worksheet
    Dim i As Integer
    Dim r As Range
    For Each ws In Worksheets
        Dim str As String: str = ws.Name & "!"
        Dim str2 As String: str2 = "'" & ws.Name & "'!"
        For Each r In ws.Cells.SpecialCells(xlCellTypeFormulas)
            If Instr(r.Formula, ":" & ws.Name) <> 0 Then
                GoTo Continue
            End If
            If Instr(r.Formula, ws.Name & ":") <> 0 Then
                GoTo Continue
            End If
            If r.HasSpill Or r.HasArray Then
                GoTo Continue
            End If
            Dim prev As String: prev = r.Formula
            r.Formula = Replace(r.Formula, str, "")
            r.Formula = Replace(r.Formula, str2, "")
            Debug.Print prev & " -> " & r.Formula
Continue:
        Next
    Next
End Sub
