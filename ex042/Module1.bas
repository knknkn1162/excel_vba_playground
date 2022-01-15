Option Explicit

Sub main()
    Dim rng As Range
    Set rng = Range("A1").CurrentRegion
    Dim i As Integer, j As Integer
    Dim pos As Integer: pos = 1
    Dim ws As Worksheet
    Set ws = Worksheets("階層DB")
    Range("A1").Resize(,4).Copy Destination:=ws.Range("A1")
    Dim db_row As Integer: db_row = 2

    Dim arr(1 To 4) As String
    For i = 2 To rng.Rows.Count
        If Cells(i, pos) <> "" Then
            ' pass
        ElseIf Cells(i,pos+1) <> "" Then
            pos = pos + 1
        ElseIf pos = 4 Then
            ' Descend
            For j = 3 To 1 Step -1
                If Cells(i, j) <> "" Then
                    pos = j
                    Exit For
                End If
            Next
        Else
            Err.Raise number:=1000, Description:="階層のフォーマットが正しくありません"
        End If
        arr(pos) = Cells(i, pos)
        If pos = 4 Then
            For j = 1 To 4
                ws.Cells(db_row,j) = arr(j)
            Next
            db_row = db_row + 1
        End If
    Next
End Sub
