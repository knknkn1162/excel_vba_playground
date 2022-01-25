Option Explicit

Function pad(str As String, tp As String, digit As Integer) As String
    Dim zeros As String: zeros = "000000000000000000000000000000000000"
    Dim spaces As String: spaces = Space(50)
    If tp = "N" Then
        ' 0
        pad = Right(zeros & str, digit)
    Else
        pad =  Left(str & spaces, digit)
    End If
End Function

Sub main()
    Dim ws As Worksheet: Set ws = Worksheets("data")
    Dim fws As Worksheet: Set fws = Worksheets("フォーマット")
    Dim rng As Range: Set rng = ws.Range("A1").CurrentRegion
    Dim rows As Integer: rows = rng.Rows.Count
    Dim cols As Integer: cols = rng.Columns.Count

    Dim padding() As String
    Dim cnt() As Integer
    ReDim padding(cols)
    Redim cnt(cols)
    Dim i As Integer, j As Integer
    For i = 1 To cols
        Dim pos As Integer
        pos = WorksheetFunction.match( _
            ws.Cells(1,i), _
            fws.Range("A1").CurrentRegion.Columns(1), 0 _
        )
        padding(i) = fws.Cells(pos, 2)
        cnt(i) = fws.Cells(pos, 3)
    Next

    Dim fname As String: fname =  "out.txt"
    Dim fd as Integer: fd = FreeFile
    Dim fpath As String: fpath = ThisWorkbook.Path & "/" & fname
    Open fpath For output As #fd
    For i = 2 To rows
        Dim str As String: str = ""
        For j = 1 To cols
            str = str & pad(ws.Cells(i,j), padding(j), cnt(j))
        Next
        Print #fd, str
    Next
    Close #fd
End Sub
