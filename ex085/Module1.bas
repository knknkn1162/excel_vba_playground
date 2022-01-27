Option Explicit
Function GetFirstDay(d As Date, n As Integer)
    GetFirstDay = DateSerial(Year(d), Month(d)+n,1)
End Function

Sub main()
    Dim tws As Worksheet: Set tws = Worksheets("入金予定")
    Dim mws As Worksheet: Set mws = Worksheets("取引先マスタ")
    Dim pay_ws As Worksheet: Set pay_ws = Worksheets("支払パターン")
    Dim date_ws As Worksheet: Set date_ws = Worksheets("祝日マスタ")

    Dim rng As Range: Set rng = tws.Range("A1").CurrentRegion
    rng.Columns(4).NumberFormatLocal = "yyyy/mm/dd(aaa)"
    Set rng = Intersect(rng.Columns(1), rng.Offset(1))
    Dim r As Range
    For Each r In rng
        Dim pay As String
        pay = WorksheetFunction.vlookup(r.Value(), mws.Range("A1").CurrentRegion, 3, False)
        Dim idx As Integer
        idx = WorksheetFunction.match(pay, pay_ws.Range("A1").CurrentRegion.Columns(1), False)
        Dim ret As Date
        ret = r.Offset(,2)
        With pay_ws
            If .Cells(idx, 2) <> "末" Then
                If Day(ret) > Day(.Cells(idx,2)) Then
                    ret = GetFirstDay(ret, 1)
                End If
            End If
            ret = GetFirstDay(ret, Val(.Cells(idx,3)))
            If .Cells(idx, 4) <> "末" Then
                If Day(ret) > Day(.Cells(idx,4)) Then
                    ret = GetFirstDay(ret, 1)
                End If
            End If
            If .Cells(idx, 4) = "末" Then
                ret = DateSerial(Year(ret), Month(ret) + 1, 0)
            Else
                ret = DateSerial(Year(ret), Month(ret), .Cells(idx,4))
            End If
            Dim comp As Date: comp = ret
            Do While True
                If WorksheetFunction.networkdays(comp, ret) <> 0 Then
                    ret = comp
                    Exit Do
                End If
                comp = DateAdd("d", -1, comp)
            Loop
        End With
        r.Offset(,3) = ret
    Next
    tws.UsedRange.EntireColumn.AutoFit
End Sub
