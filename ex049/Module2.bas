Option Explicit

Sub main2()
    Dim wsIn As Worksheet: Set wsIn = Worksheets("49In")
    Dim wsOut As Worksheet: Set wsOut = Worksheets("49Out")
    ' header
    wsOut.Range("A1").Resize(,4) = wsIn.Range("A1").Resize(,4).Value
    Dim rng As Range
    Set rng = wsIn.Range("A1").CurrentRegion
    wsIn.AutoFilterMode = False
    wsOut.Range("A1").CurrentRegion.Offset(1).Clear
    Dim pos As Integer: pos = 2
    pos = setDisplayFormat(wsOut, rng, pos, 4, vbRed, xlFilterFontColor)
    pos = setDisplayFormat(wsOut, rng, pos, 4, vbRed, xlFilterCellColor)
    pos = setDisplayFormat(wsOut, rng, pos, 4, vbYellow, xlFilterCellColor)
    wsOut.Range("A1").CurrentRegion.Sort key1:=wsOut.Range("A1"), Header:=xlYes
    wsIn.AutoFilterMode = False
End Sub

Function setDisplayFormat(wsOut As Worksheet, _
    rng As Range, _
    pos As Integer, _
    Field As Integer, _
    Criteria1 As Variant, _
    Operator As XlAutoFilterOperator _
) As Integer
    rng.AutoFilter Field:=Field, Criteria1:=Criteria1, Operator:=Operator
    Dim cnt As Integer
    cnt = rng.Columns(1).SpecialCells(xlCellTypeVisible).Count-1
    rng.Offset(1).Copy Destination:=wsOut.Cells(pos, 1)

    With wsOut.Cells(pos, Field).Resize(cnt)
        .ClearFormats
        Select Case Operator
            Case xlFilterFontColor
                .Font.Color = Criteria1
            Case xlFilterCellColor
                .Interior.Color = Criteria1
　　　　End Select
　　End With
    setDisplayFormat = pos+cnt
End Function
