Option Explicit

Sub main()
    Dim root As String
    root = ThisWorkbook.Path
    Dim fname As String
    Dim backup_dir As String
    backup_dir = root & "/BACKUP"
    fname = Dir(backup_dir & "/")
    Dim prev As String
    prev = Format(Date()-30, "yyyymmdd")
    Do while fname <> ""
        Dim pos As Integer
        pos = InStr(fname, ".")
        ' len(yyyymmddhhmm)=12
        Dim fDate As String
        ' Trash
        fDate = Mid(fname, pos-12, 4) & "/" & Mid(fname, pos-8, 2) & "/" & Mid(fname, pos-6, 2)
        If fDate <= prev Then
            ' Killは読み取り専用は削除できない
            On Error Resume Next
            ' Msgbox "KIll " & fname
            Kill backup_dir & "/" & fname
            Err.Clear
        End If
        fname = Dir()
    loop
End Sub
