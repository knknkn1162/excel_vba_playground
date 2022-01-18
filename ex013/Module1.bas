Option Explicit

Sub main()
    Dim c As Range
    Dim srng As Range
    Set c = Cells.Find(What:="注意")
    Set srng = c
    If c Is Nothing Then
        Exit Sub
    End If
    Do
        Dim pos As Long
        pos = 1
        Do While True
            pos = Instr(pos, c.Value, "注意")
            If pos = 0 Then
                Exit Do
            End If
            With c.Characters(pos, 2).Font
                .Color = vbRed
                .Bold = True
            End With
            pos = pos + 2
        Loop
        Set c = Cells.FindNext(c)
    Loop Until c.Address = srng.Address
End Sub

Sub ReplaceFormat(ByRef rng As Range)
    Dim pos As Long
    pos = 1
    Do While True
        pos = Instr(pos, rng.Value, "注意")
        If pos = 0 Then Exit Do
        With rng.Characters(pos, 2).Font
            .Color = vbRed
            .Bold = True
        End With
        pos = pos + 2
    Loop
End Sub

Sub main2()
    Dim target As Range, c As Range
    ' SpecialCellsメソッドは、指定に一致するセルが存在しない場合はエラーとなります。
    On Error Resume Next
    Set target = Intersect(Selection, Selection.SpecialCells(xlCellTypeConstants, xlTextValues))
    Err.Clear
    If target Is Nothing Then Exit Sub
    For Each c In target
        ' Call statement is must-need.
        Call ReplaceFormat(c)
    Next
End Sub

