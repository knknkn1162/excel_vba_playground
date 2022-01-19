Option Explicit

Function CreateWorksheet(fname As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Worksheets(fname).Delete
    Err.Clear
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    With ws
        .Name = fname
        .Range("A1").resize(,2) = Array("シート名", "印刷ページ数")
    End With
    Set CreateWorksheet = ws
End Function

Sub main()
    Dim ws As Worksheet: Set ws = CreateWorkSheet("目次")
    Dim sht As Worksheet
    Dim pos As Integer: pos = 2
    For each sht In WorkSheets
        If ws.Name = sht.Name Then GoTo Continue
        If sht.Visible <> xlSheetVisible Then
            ws.Cells(pos, 1) = sht.Name
            ws.Cells(pos,2) = 0
            pos = pos + 1
            GoTo Continue
        End If
        ' シートへのハイパーリンクでは、シングルクォートでシート名を囲まないと、記号やスペースを含むシートへのリンクが正しく機能しません。
        ws.Hyperlinks.Add Anchor:=ws.Cells(pos,1), _
            Address:="", _
            SubAddress:="'" & Replace(sht.Name, "'", "''") & "'!A1", _
            TextToDisplay:=sht.Name
        ws.Cells(pos,2) = sht.PageSetup.Pages.Count
        pos = pos + 1
Continue:
    Next
End Sub
