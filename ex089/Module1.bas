Option Explicit

Sub CreateDir(n As String)
    On Error Resume Next
    RmDir n
    Mkdir n
    On Error GoTo 0
End Sub

Function GetFileUpdateTime(fpath As String) As Date
    On Error Resume Next
    GetFileUpdateTime = FileDateTime(fpath)
    On Error GoTo 0
End Function

Sub FilesCopy(src As String, dst As String)
    Dim fname As String: fname = Dir(src & "/*.*")
    Do While fname <> ""
        Dim srcpath As String: srcpath = src & "/" & fname
        Dim dstpath As String: dstpath = dst & "/" & fname
        Dim d As Date: d = GetFileUpdateTime(dstpath)
        If Val(d) <> 0 Then
            If d > FileDateTime(srcpath) Then GoTo Continue
        End If
        FileCopy srcpath, dstpath
Continue:
        fname = Dir()
    Loop
End Sub

Sub main()
    Dim adir As String: adir = ThisWorkbook.Path & "/ex089_A"
    Dim bdir As String: bdir = ThisWorkbook.Path & "/ex089_B"
    Dim cdir As String: cdir = ThisWorkbook.Path & "/ex089_C"

    Call CreateDir(cdir)
    Call FilesCopy(adir, cdir)
    Call FilesCopy(bdir, cdir)
End Sub
