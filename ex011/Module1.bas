Option Explicit

Sub main()
    Dim rng As Range
    For Each rng In Range("A1").CurrentRegion
        If rng.Address <> rng.MergeArea(1).Address Then
            GoTo Continue
        End If
        If rng.MergeCells Then
            rng.AddComment "セル結合されています"
            ' rng must be the head of merged Cells
            rng.Comment.Visible = True
        End If
Continue:
    Next
End Sub
