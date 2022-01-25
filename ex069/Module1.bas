Option Explicit

Function FindMergeCells(rng As Range) As Range
    With Application.FindFormat
        .Clear
        .MergeCells = True
    End With
    Dim ws As Worksheet: Set ws = rng.Worksheet
    Set FindMergeCells = ws.Cells.Find(What:="", After:=rng, SearchFormat:=True)
End Function

Sub main()
    Dim ws As Worksheet
    For Each ws In Worksheets
        Dim rng As Range
        Set rng = FindMergeCells(ws.Range("A1"))
        Do While Not rng Is Nothing
            Dim val As Variant: val = rng.MergeArea(1).Value
            Dim r As Range: Set r = rng.MergeArea
            rng.UnMerge
            r.Value() = val
            Set rng = FindMergeCells(rng)
        Loop
Continue:
    Next
End Sub
