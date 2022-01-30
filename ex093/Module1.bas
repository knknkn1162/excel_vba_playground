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
    Dim twb As Workbook: Set twb = ThisWorkbook
    Dim root As String: root = twb.Path & "/ex093"
    Dim srcs As String: srcs = root & "/月別"
    Dim dsts As String: dsts = root & "/支店別"
    CreateDir(dsts)

    Dim tws As Worksheet: Set tws = twb.Worksheets(1)
    Dim flag As boolean: flag = True
    Dim fname As String: fname = Dir(srcs & "/*.xls")

    Dim ws As Worksheet
    Dim rng As Range
    Dim pos As Integer: pos = 2
    ' copy data to thisworkbook
    Dim orig As Boolean
    Do While fname <> ""
        orig = Application.ScreenUpdating
        Application.ScreenUpdating = False
        Dim wb As Workbook: Set wb = Workbooks.Open(srcs & "/" & fname)
        If flag Then
            wb.Worksheets(1).Range("A1").CurrentRegion.Rows(1).Copy _
                Destination:=tws.Cells(1,1)
            flag = False
        End If
        For Each ws In wb.Worksheets
            Set rng = ws.Range("A1").CurrentRegion
            Set rng = Intersect(rng, rng.Offset(1))
            rng.Copy Destination:=tws.Cells(pos,1)
            pos = pos + rng.Rows.Count
        Next
        wb.Close SaveChanges:=False
        Application.ScreenUpdating = orig
        fname = Dir()
    Loop
    tws.UsedRange.EntireColumn.AutoFit
    ' sort
    tws.Range("A1").CurrentRegion.Sort _
        key1:= tws.Range("A1"), order1:=xlAscending, Header:=xlYes

    ' filter -> export
    pos= 2
    Dim comp As String: comp = tws.Range("A2").Value()
    Do While comp <> ""
        tws.AutoFilterMode = False
        tws.Range("A1").AutoFilter Field:=1, Criteria1:=comp
        Dim rs As Range: Set rs = tws.Cells.SpecialCells(xlCellTypeVisible)
        pos = pos + WorksheetFunction.Subtotal(3, tws.Range("A1").CurrentRegion.Columns(1))-1

        orig = Application.ScreenUpdating
        Application.ScreenUpdating = False
        With Workbooks.Add
            With .Worksheets(1)
                .Name = comp
                tws.Cells.SpecialCells(xlCellTypeVisible).Copy Destination:=.Range("A1")
            End With
            .SaveAs dsts & "/" & comp & ".xlsx"
            .Close
        End With
        Application.ScreenUpdating = orig
        comp = tws.Cells(pos,1)
        tws.AutoFilterMode = False
    Loop
End Sub
