Option Explicit

Sub main()
    Dim n As Name
    Dim visibleCnt As Integer, deleteCnt As Integer
    visibleCnt = 0: deleteCnt = 0
    For Each n in Names
        If n.Visible = False Then visibleCnt = visibleCnt + 1
        ' 非表示の名前定義は表示にする
        n.Visible = True
        Debug.Print "Name: " & n.Name & ", RefersTo: " & n.RefersTo
        ' [#] escapes `#`
        If n.RefersTo Like "*[#]REF!*" Then
            deleteCnt = deleteCnt + 1
            n.Delete
        End If
    Next
    Msgbox "非表示件数: " & visibleCnt & vbLf & _
        "削除件数: " & deleteCnt
End Sub
