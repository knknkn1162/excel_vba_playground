Option Explicit

Sub main()
    Dim d As Date: d = #2020/4/1#
    Dim s As WorkSheet
    Dim i As Integer
    For i = 1 To 12
        Dim ws As WorkSheet: Set ws = WorkSheets(Format(d, "yyyy年mm月"))
        ' move backmost
        ws.Move After:=Sheets(Sheets.Count)
        d = DateAdd("m", 1, d)
    Next
End Sub

Sub main2()
    Dim i, s, ary
    ReDim ary(1 To Sheets.Count)
    For i = 1 To Sheets.Count
        ary(i) = Sheets(i).Name
    Next
    For Each s In WorksheetFunction.Sort(ary, , 1, 1)
        Sheets(s).Move After:=Sheets(Sheets.Count)
    Next
End Sub
