Option Explicit

Function checkRemarks(str As String) As Boolean
    Dim blackList() As String
    blackList = Split("不要,削除", ",")
    checkRemarks = false
    ' for eachを配列で使用する場合は、バリアント型の配列でなければなりません
    Dim item As Variant
    For Each item In blackList
        If Instr(str, item) <> 0 Then
            checkRemarks = True
            Exit Function
        End If
    Next
End Function

Sub main()
    Dim i As Integer
    For i = Cells(Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If Cells(i, 3) = "" And checkRemarks(Cells(i, 4)) Then
            Rows(i).Delete
        End If
    Next
End Sub

Sub main2()
    Dim ws As Worksheet
    Set ws = WorkSheets(1)
    ws.AutoFilterMode = false
    Dim rng As Range
    With ws.Range("A1").CurrentRegion
        .AutoFilter field:=3, Criteria1:=""
        .AutoFilter field:=4, Criteria1:="*削除*", Operator:=xlOr, Criteria2:="*不要*"
        Set rng = Intersect(.Offset(1), .SpecialCells(xlCellTypeVisible))
        If Not rng Is Nothing Then rng.EntireRow.Delete
    End With
    ws.AutoFilterMode = false
End Sub
