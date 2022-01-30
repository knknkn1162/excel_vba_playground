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

Function setAppConfig(conf As Boolean) As Boolean
    Dim orig As Boolean
　　With Application
        orig = .DisplayAlerts
　　　　.DisplayAlerts = conf
　　End With
    setAppConfig = orig
End Function

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
    Dim bdir As String: bdir = ThisWorkbook.Path & "/ex043_out"
    Call CreateDir(bdir)
    fpath = bdir & "/out.csv"
    Dim orig As Boolean: orig = setAppConfig(false)
    ' write as SHIFT_JIS encoding
    ' In Linux, check with `nkf --ic=SHIFT_JIS out.csv`
    wb.SaveAs FileName:=fpath, FileFormat:=xlCSV
    Call setAppConfig(orig)
    ActiveWorkbook.Close SaveChanges:=False
End Sub
