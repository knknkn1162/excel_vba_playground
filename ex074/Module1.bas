Option Explicit

Sub main()
    Dim dbws As Worksheet: Set dbws = Worksheets("DB")
    Dim tblws As Worksheet: Set tblws = Worksheets("売上")
    Dim pos As Integer: pos = 1
    Dim k As Integer: k = 2
    Dim i As Integer
    Dim r As Range
    Do While True
        If Cells(pos,1) = "" Then
            pos = pos + 1
            GoTo Continue
        End If
        Dim rng As Range: Set rng = Cells(pos,1).CurrentRegion
        Dim cols As Integer: cols = rng.Columns.Count-2
        For Each r In Intersect(rng, rng.Columns(1).Offset(2))
            dbws.Cells(k, 1).Resize(cols,4).Value() = Array( _
                rng.Cells(1,1), rng.Cells(1,2), _
                Cells(r.Row,1), Cells(r.Row, 2) _
            )
            dbws.Cells(k,5).Resize(cols).Value() = _
                WorksheetFunction.transpose(Cells(pos+1,3).Resize(1, cols))
            dbws.Cells(k,6).Resize(cols).Value() = _
                WorksheetFunction.transpose(r.Offset(,2).Resize(1,cols))
            k = k + cols
        Next
        ' next cell
        pos = pos + rng.Rows.Count - 1 + 1
Continue:
        If pos > tblws.Cells(Rows.Count,1).End(xlup).Row Then
            Exit Do
        End If
    Loop
    With dbws
        Dim drng As Range: Set drng = .Range("A1").CurrentRegion
        .AutoFilterMode = False
        ' exclude [1-4]Q計
        .Range("A1").Autofilter Field:=5, Criteria1:="*Q計"
        drng.Offset(1).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        .AutoFilterMode = False
    End With
End Sub
