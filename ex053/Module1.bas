Option Explicit

Sub main()
    Dim d As Date: d = #2021/12/31#
    Dim rng As Range
    d = DateAdd("yyyy",-35, d)
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects(1)
    ' See the link; https://excel-ubara.com/excelvba5/EXCELVBA212.html
    With tbl.DataBodyRange
        .AutoFilter Field:=tbl.ListColumns("誕生日").Index, _
            Criteria1:="<=" & CLng(d)
        .AutoFilter Field:=tbl.ListColumns("性別").Index, _
            Criteria1:="*男*"
        tbl.ListColumns("備考").Range.SpecialCells(xlCellTypeVisible) = "対象"
    End With
    tbl.AutoFilter.ShowAllData
End Sub
