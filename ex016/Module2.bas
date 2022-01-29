Option Explicit

Sub main2()
    Dim rng As Range
    Dim cands As Range
    ' ignore 該当するセルが見つかりません
    On Error Resume Next
    Set cands = Cells.SpecialCells(xlCellTypeConstants, xlTextValues)
    Err.clear
    On Error Goto 0
    If cands Is Nothing Then Exit Sub
    For Each rng In cands
        Dim str As String, buf As String
        Dim v As Variant
        str = rng.Value
        buf = ""
        str = Replace(str, vbCrLf, vbLf)
        For Each v In Split(str, vbLf)
            If v <> "" Then buf = buf & v & vbLf
        Next
        If Len(buf) = 0 Then
            rng.Value = ""
            GoTo Continue
        End If
        ' remove lastword; vbLf
        rng.Value = Left(buf, Len(buf)-1)
Continue:
    Next
    
End Sub
