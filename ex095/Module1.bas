Option Explicit

Function CreateWorksheet(str As String)
    On Error Resume Next
    Worksheets.Delete(str)
    On Error GoTo 0
    Dim tws As Worksheet: tws = Worksheets.Add(Before:=Worksheets.Count)
End Function

Function GetText(ByRef shp As Shape) As String
    GetText = ""
    On Error Resume Next
    GetText = shp.TextFrame.Characters.Text
    On Error GoTo 0
End Function

Sub main()
    Dim ws As Worksheet
    Dim tws As Worksheet: Set tws = Worksheets.Add(Before:=Worksheets(1))
    tws.Name = "検索結果"
    Dim pos As Integer: pos = 1
    Dim pat As String: pat = Inputbox("検索文字列を入力してください")
    For Each ws In Worksheets
        Dim shp As Shape
        For Each shp In ws.Shapes
            Dim str As String: str = GetText(shp)
            If str Like "*" & pat & "*" Then
                Dim addr As String
                addr = shp.TopLeftCell.Address(External:=True)
                addr = Mid(addr, Instr(addr, "]")+1)
                tws.Hyperlinks.Add Anchor:=tws.Cells(pos, 1), _
                    Address:="", _
                    SubAddress:=Replace(addr, "'", "''"), TextToDisplay:=addr
                tws.Cells(pos, 2) = str
                pos = pos + 1
            End If
        Next
    Next
    tws.UsedRange.EntireColumn.AutoFit
End Sub
