Option Explicit

Function isPathRooted(str As String) As Boolean
    isPathRooted = False
    If Left(str, 1) = "/" Then
        isPathRooted = True
    End If
    If (Not Application.OperatingSystem Like "*Mac*") And Mid(str, 2, 1) = ":" Then
        isPathRooted = True
    End If
End Function

Sub main()
    Dim ws As Worksheet
    Set ws = Worksheets("ファイル一覧")
    ws.Cells(1,1).Resize(1,3) = _
        Array("ファイル一覧", "更新日時", "サイズ")
    Dim searchDir As String
    Dim root As String
    root = ThisWorkbook.Path
    searchDir = InputBox("フォルダを入力してください" & vbLf & "(相対パスでも可)")
    If Not isPathRooted(searchDir) Then
        searchDir = root & "/" & searchDir
    End If
    Dim fname As String
    fname = Dir(searchDir & "/*")
    Dim pos As Integer
    pos = 2
    Columns("C").NumberFormatLocal = "#,##0"
    Do While fname <> ""
        Dim fpath As String
        fpath = searchDir & "/" & fname
        If fpath Like "*.xls" Or fpath Like "*.xls[mx]" Then  
            ws.Hyperlinks.Add Anchor:=Cells(pos, 1), _
                Address:=fpath, TextToDisplay:=fname
        Else
            Cells(pos, 1) = fname
        End If
        Cells(pos, 2) = FileDateTime(fpath)
        Cells(pos, 3) = FileLen(fpath)
        pos = pos + 1
        fname = Dir()
    Loop
    ws.UsedRange.EntireColumn.AutoFit
End Sub
