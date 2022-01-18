Option Explicit

Function CreateWorksheet(name As String, header As Range) As Worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    On Error Resume Next
    Worksheets(name).Delete
    Err.Clear
    Dim nws As Worksheet
    Set nws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    With nws
        .Name = name
        .Range("A1").Resize(,header.Columns.Count) = header.Value
    End With
    ws.Activate
    Set CreateWorksheet = nws
End Function

Sub main()
    Dim rng As Range
    Set rng = Range("A1").CurrentRegion
    Dim ws As Worksheet: Set ws = CreateWorksheet("49Out", rng.Resize(1))
    Dim fc As FormatCondition
    Dim r As Range
    Dim pos As Integer: pos = 2
    For Each fc In Columns("D").FormatConditions
        For Each r In fc.AppliesTo
            Dim flag As Boolean: flag = True
            With r.DisplayFormat
                ' If the condition meets..
                If (Not IsNull(fc.Font.Color)) And .Font.Color <> fc.Font.Color Then
                    flag = False
                End If
                If (Not IsNull(fc.Interior.Color)) And _
                    (Not IsNull(fc.Interior.ColorIndex)) Then
                    If .Interior.ColorIndex = fc.Interior.ColorIndex Then
                    ' 塗りつぶしなしか判定
                    ElseIf .Interior.Color <> fc.Interior.Color Then
                        flag = False
                    End If
                End If
            End With
            If flag Then
                r.EntireRow.Copy
                ws.Cells(pos,1).PasteSpecial Paste:=xlPasteFormats
                ws.Cells(pos,1).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                pos = pos + 1
            End If
        Next
    Next
End Sub

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
