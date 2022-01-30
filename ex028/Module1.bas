Option Explicit

Sub main()
    Dim ws As Worksheet
    Dim root As String
    root = ThisWorkbook.Path & "/ex028_out_"
    Dim busyo As String, namae As String
    Dim d As String
    ' Remove old
    For Each ws In Worksheets
        ws.Visible = xlSheetVisible
        busyo = Split(ws.Name, "_")(0)
        d = root & busyo
        ' フォルダ内にファイルが残っている場合はエラーとなります。
        If Dir(d, vbDirectory) <> "" Then
            Kill d & "/*.xlsx"
            RmDir(d)
        End If
    Next

    ' create
    For Each ws In Worksheets
        busyo = Split(ws.Name, "_")(0)
        d = root & busyo
        If Dir(d, vbDirectory) = "" Then
            MkDir(d)
        End If
        ws.Copy
        ActiveWorkbook.SaveAs Filename:= d & "/" & ws.Name
        ActiveWorkbook.Close
    Next
End Sub
