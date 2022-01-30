Option Explicit

Sub main()
    Dim rng As Range
    Set rng = Range("A1").CurrentRegion
    Dim ws As Worksheet: Set ws = Worksheets("49Out")
    ws.Range("A1").Resize(1,rng.Columns.Count).Value() = rng.Resize(1).Value()

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
