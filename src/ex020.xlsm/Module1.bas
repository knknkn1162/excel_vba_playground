Option Explicit

Sub main()
    Dim ws As Workbook
    Set ws = ThisWorkbook
    On Error Resume Next
    MkDir ws.Path & "/BACKUP"
    Err.Clear

    Dim str As String
    str = ws.Name
    str = Replace(str, ".", "_" & Format(Now(), "yyyymmddhhmm") & ".")
    ws.SaveCopyAs FileName:= ws.Path & "/BACKUP/" & str
End Sub
