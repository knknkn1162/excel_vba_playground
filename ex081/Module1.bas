Option Explicit

Function GetCand() As Range
    Dim rng(1) As Range
    On Error Resume Next
    Set rng(0) = Cells.SpecialCells(xlCellTypeConstants)
    Set rng(1) = Cells.SpecialCells(xlCellTypeFormulas)
    Err.Clear
    On Error GoTo 0
    Set GetCand = rng(0)
    Dim i As integer
    For i = 0 To 1
        If Not rng(i) Is Nothing Then Set GetCand = Union(GetCand, rng(i))
    Next
End Function

Sub main()
    Dim ws As Worksheet: Set ws = ActiveSheet
    ws.UsedRange.EntireRow.Hidden = False
    Dim area As Range
    Dim rng As Range: Set rng = GetCand()
    For Each area In rng.Areas
        Dim r As Range: Set r = area.Cells(1,1)
        ' This is necessary
        r.Activate
        Dim fl As Filter
        Dim i As Integer
        On Error Resume Next
        ws.AutoFilter.ShowAllData
        Err.Clear
        On Error GoTo 0
Continue:
    Next
End Sub
