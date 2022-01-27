Option Explicit

Sub CreateDir(n As String)
    On Error Resume Next
    Mkdir n
    On Error GoTo 0
End Sub

Function GetFileUpdateTime(fpath As String) As Date
    On Error Resume Next
    GetFileUpdateTime = FileDateTime(fpath)
    On Error GoTo 0
End Function

Sub WalkCopy(src As String, dst As String)
    Dim fname As String: fname = Dir(src & "/*.*", vbDirectory + vbNormal + vbReadOnly + vbHidden)
    Dim arr() As String
    Dim brr() As String
    Dim pos As Integer: pos = 0
    Do While fname <> ""
        If fname = "." Then GoTo Continue
        If fname = ".." Then GoTo Continue
        Dim srcpath As String: srcpath = src & "/" & fname
        Dim dstpath As String: dstpath = dst & "/" & fname
        If GetAttr(srcpath) = vbDirectory Then
            Redim Preserve arr(pos)
            Redim Preserve brr(pos)
            arr(pos) = srcpath
            brr(pos) = dstpath
            pos = pos + 1
            GoTo Continue
        End If
        ' Assume that srcpath is file
        Dim d As Date: d = GetFileUpdateTime(dstpath)
        If Val(d) <> 0 Then
            If d > FileDateTime(srcpath) Then GoTo Continue
        End If
        FileCopy srcpath, dstpath
Continue:
        fname = Dir()
    Loop
    ' Lastly, walk subdirectory at once
    Dim i As Integer
    For i = 0 To pos-1
        CreateDir(brr(i))
        Call WalkCopy(arr(i), brr(i))
    Next
End Sub

Sub main()
    Dim adir As String: adir = ThisWorkbook.Path & "/ex089_A"
    Dim bdir As String: bdir = ThisWorkbook.Path & "/ex089_B"
    Dim cdir As String: cdir = ThisWorkbook.Path & "/ex089_C"

    Call CreateDir(cdir)
    Call WalkCopy(adir, cdir)
    Call WalkCopy(bdir, cdir)
End Sub
