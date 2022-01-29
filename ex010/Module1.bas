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
