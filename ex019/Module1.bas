Option Explicit

Sub main()
    const checked As String = "checked"
    Dim sp As Shape
    Dim ws As Worksheet
    Set ws = WorkSheets(1)
    For Each sp In ws.Shapes
        ' See https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
        If sp.Type = msoFormControl Or sp.Type = msoOLEControlObject Then
            GoTo Continue
        End If
        ' 繰り返し実行しても増殖しないように工夫
        If sp.Name = checked Then GoTo Continue
        sp.Name = checked
        With sp.Duplicate
            .left = sp.left + sp.width
            .top = sp.top
            .Name = checked
        End With
Continue:
    Next
End Sub
