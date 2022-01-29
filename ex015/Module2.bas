Option Explicit

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
