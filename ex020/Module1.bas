Option Explicit

Sub main()
    Dim ws As Workbook
    Set ws = ThisWorkbook
    Dim bdir As String: bdir = ws.Path & "/ex020_BACKUP"
    ' mkdir -p
    If Dir(bdir, vbDirectory) = "" Then
        MkDir bdir
    End If

    Dim str As String
    str = ws.Name
    str = Replace(str, ".", "_" & Format(Now(), "yyyymmddhhmm") & ".")
    ws.SaveCopyAs FileName:= ws.Path & "/BACKUP/" & str
End Sub
