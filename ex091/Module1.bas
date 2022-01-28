Option Explicit

Sub main()
    Dim kws As Worksheet: Set kws = Worksheets("勤怠")
    Dim zws As Worksheet: Set zws = Worksheets("残業")
    kws.Activate

    Range("A1").CurrentRegion.Sort _
        key1:= Range("A1"), order1:=xlAscending, _
        key1:= Range("B1"), order1:=xlAscending, _
        Header:=xlYes
    Dim previdx As Integer: previdx = Range("A2").Value()
    Dim prevd As Date: prevd = Range("B2").Value()
    Dim sum As Long: sum = 0
    Dim pos As Integer: pos = 2
    Dim i As Integer
    For i = 2 To Cells(Rows.Count,1).End(xlUp).Row + 1
        If previdx <> Cells(i,1) Or _
            Month(prevd) <> Month(Cells(i,2)) Or _
            Year(prevd) <> Year(Cells(i,2)) _
            Then
            zws.Cells(pos, 1).Resize(1,3).Value() = Array( _
                previdx, _
                Format(prevd, "yyyymm"), _
                TimeSerial(0,Fix(sum/30)*30,0) _
            )
            pos = pos + 1
            previdx = Cells(i,1)
            prevd = Cells(i,2)
            sum = 0
        End If
        Dim overworkn As Long
        overworkn = WorksheetFunction.max(0, DateDiff("n", WorksheetFunction.max(#9:00:00#, Cells(i,3)), Cells(i,4)) - 60 - 8*60)
        sum = sum + overworkn
    Next
End Sub
