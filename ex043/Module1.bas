Option Explicit

Sub CreateDir(d As String)
    Dim fname As String: fname = Dir(d & "/*.*")
    Do While fname <> ""
        Kill d & "/" & fname
        fname = Dir()
    Loop
    On Error Resume Next
    RmDir d
    Mkdir d
    On Error GoTo 0
End Sub

Sub main()
    Dim i As Integer
    Dim fd as Integer: fd = FreeFile
    Dim bdir As String: bdir = ThisWorkbook.Path & "/ex043_out"
    Call CreateDir(bdir)
    Dim fpath As String: fpath = bdir & "/out.csv"
    Open fpath For output As #fd
    Dim box(1 To 4) As String
    For i = 1 To 4
        box(i) = Cells(1,i)
    Next
    Print #fd, Join(box, ",")
    For i = 2 To Cells(Rows.Count,1).End(xlUp).Row
        Dim arr(1 To 4) As String
        arr(1) = Format(Cells(i,1), "yyyy/mm/dd")
        Dim b As Long: b = Cells(i,2)
        Dim cpos As Integer
        arr(2) = b & ""
        arr(3) = Format(Cells(i,3), "0.00")
        ' If Cells(i,4) contains double-quote, escape.
        arr(4) = """" & Replace(Cells(i,4), """", """""") & """"
        Print #fd, Join(arr, ",")
    Next
    ' write as SHIFT_JIS encoding
    ' In Linux, check with `nkf --ic=SHIFT_JIS out.csv`
    Close #fd
End Sub
