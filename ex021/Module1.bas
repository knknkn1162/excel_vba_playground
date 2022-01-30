Option Explicit

Sub main()
    Dim root As String
    root = ThisWorkbook.Path
    Dim fname As String
    Dim backup_dir As String
    backup_dir = root & "/ex021_BACKUP"
    fname = Dir(backup_dir & "/")
    Dim prev As String
    prev = Format(Date()-30, "yyyymmdd")
    Do while fname <> ""
        Dim pos As Integer
        pos = InStr(fname, ".")
        ' len(yyyymmddhhmm)=12
        Dim fDate As String: fDate = Mid(fname, pos-12, 8)
        Debug.Print fDate & " vs " & prev
        ' Trash
        If fDate <= prev Then
            ' Killは読み取り専用は削除できない
            On Error Resume Next
            ' Msgbox "KIll " & fname
            Kill backup_dir & "/" & fname
            Err.Clear
            On Error GoTo 0
        End If
        fname = Dir()
    loop
End Sub
