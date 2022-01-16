Option Explicit

Sub main()
    Dim i As Integer
    Dim fd as Integer: fd = FreeFile
    Dim fpath As String: fpath = ThisWorkbook.Path & "/out.csv"
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

Sub main2()
    Dim ws As Worksheet: Set ws = Activesheet
    Dim wb As Workbook: Set wb = ActiveWorkbook
    ws.Copy
    With wb.Worksheets(1)
        .Columns(1).NumberFormatLocal = "yyyy/mm/dd"
        .Columns(2).NumberFormatLocal = "0"
        .Columns(3).NumberFormatLocal = "0.00"
    End With

    Dim fpath As String
    fpath = wb.Path & "/out.csv"
    ' write as SHIFT_JIS encoding
    ' In Linux, check with `nkf --ic=SHIFT_JIS out.csv`
    wb.SaveAs FileName:=fpath, FileFormat:=xlCSV
    wb.Close SaveChanges:=False
End Sub
