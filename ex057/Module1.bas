Option Explicit

Sub main()
    Dim root As String: root = ThisWorkbook.Path
    Dim bdir As String: bdir = root & "/ex057_BACKUP"
    Dim fname As String
    Dim latestFile As String
    fname = Dir(bdir & "/*.*")
    Dim tmp_ws As Worksheet: Set tmp_ws = Worksheets(1)
    Dim pos As Integer: pos = 1
    Do While fname <> ""
        Dim dt As Date: dt = FileDateTime(bdir & "/" & fname)
        tmp_ws.Cells(pos,1) = dt
        tmp_ws.Cells(pos,2) = fname
        pos = pos + 1
        fname = Dir()
    Loop

    ' remove files
    If pos = 1 Then Exit Sub
    Range("A1").Sort _
        key1:= Range("A1"), order1:=xlAscending, Header:=xlNo
    Dim i As Integer
    Dim day As String: day = "2100/12/31"
    For i = pos-1 To 1 Step -1
        Dim str As String: str = Format(Cells(i,1),"yyyy/mm/dd")
        If str <> day Then GoTo Continue
        ' Msgbox bdir & "/" & Cells(i,2)
        On Error Resume Next
        Kill bdir & "/" & Cells(i,2)
        Err.Clear
Continue:
        day = str
    Next
    tmp_ws.UsedRange.Clear
End Sub
