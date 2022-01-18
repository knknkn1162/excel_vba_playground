Option Explicit

Sub main()
    Dim ws As Workbook
    Set ws = ThisWorkbook
    ' mkdir -p
    If Dir(ws.Path & "/BACKUP", vbDirectory) = "" Then
        MkDir ws.Path & "/BACKUP"
    End If

    Dim str As String
    str = ws.Name
    str = Replace(str, ".", "_" & Format(Now(), "yyyymmddhhmm") & ".")
    ws.SaveCopyAs FileName:= ws.Path & "/BACKUP/" & str
End Sub
