Option Explicit

Sub main()
    Dim root As String: root = ThisWorkbook.Path
    Dim bdir As String: bdir = root & "/ex057_BACKUP"
    Dim fname As String
    Dim cand As Date: cand = #1970/1/1 00:00:00#
    Dim latestFile As String
    fname = Dir(bdir & "/*.*")
    Do While fname <> ""
        Dim dt As Date: dt = FileDateTime(bdir & "/" & fname)
        If cand < dt Then
            cand = dt
            latestFile = fname
        End If
        fname = Dir()
    Loop

    ' delete exclude: str
    fname = Dir(bdir & "/*.*")
    Do While fname <> ""
        If latestFile = fname Then GoTo Continue
        On Error Resume Next
        Kill bdir & "/" & fname
        Err.Clear
Continue:
        fname = Dir()
    Loop
End Sub
