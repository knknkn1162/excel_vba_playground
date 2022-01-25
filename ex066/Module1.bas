Option Explicit


' search directory under ./books/ex066 as follows:
' ./books/ex066/
' ├── ex066.xlsm
' ├── ex966.xlsm
' ├── sub1
' ├── sub2
' │   └── ex966.xlsm
' └── sub3
'     └── ex066.xlsm
Function Walk(path As String, comp As String, pos As Integer) As Integer
    Dim fname As String
    fname = Dir(path & "/*.*", vbDirectory + vbNormal + vbReadOnly + vbHidden)
    Dim arr() As String
    Dim aidx As Integer: aidx = 0
    Redim arr(aidx)
    ' get file only
    Do While fname <> ""
        If fname = "." Then GoTo Continue
        If fname = ".." Then GoTo Continue
        Dim fpath As String: fpath = path & "/" & fname
        If GetAttr(fpath) = vbDirectory Then
            arr(aidx) = fpath
            aidx = aidx+1
            Redim Preserve arr(aidx)
        ElseIf fname = comp Then
            ActiveSheet.Cells(pos, 1).Resize(1,3) = Array( _
                fpath, FileDateTime(fpath), FileLen(fpath) _
            )
            pos = pos + 1
        End If
Continue:
        fname = Dir()
    Loop
    ' Lastly, walk subdirectory at once
    Dim i As Integer
    For i = 0 To aidx-1
        pos = Walk(arr(i), comp, pos)
    Next
    Walk = pos
End Function

Sub main()
    Dim root As String: root = ThisWorkbook.Path & "/" & "ex066"
    Dim fname As String
    Range("A1").Resize(1,3) = Array("フルパス", "更新日時", "ファイルサイズ")
    Dim pos As Integer
    pos = Walk(root, ThisWorkbook.Name, 2)
    ActiveSheet.UsedRange.EntireColumn.AutoFit
End Sub
