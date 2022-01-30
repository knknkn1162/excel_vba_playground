Option Explicit

Sub setLinkPicture(rng As Range, ws As Worksheet)
    Dim mws As Worksheet: Set mws = rng.Worksheet
    ws.Range("A1").CurrentRegion.Copy
    With mws.Pictures.Paste(Link:=True)
        .ShapeRange.LockAspectRatio = True
        .Width = rng.Width
        .Height = rng.Height
        .Width = WorksheetFunction.min(rng.Width, .Width)
        .Height = WorksheetFunction.min(rng.Height, .Height)
        .Top = rng.Top
        .Left = rng.Left + (rng.Width - .Width)/2
    End With
    Application.CutCopyMode = False
End Sub

Sub main()
    Dim ws1 As Worksheet: Set ws1 = Worksheets("元表1")
    Dim ws2 As Worksheet: Set ws2 = Worksheets("元表2")
    Dim mws As Worksheet: Set mws = Worksheets("まとめ")
    Dim p As Picture
    For each p In mws.Pictures
        On Error Resume Next
        p.Delete
    Next
    Call SetLinkPicture(mws.Range("A1:J20"), ws1)
    Call SetLinkPicture(mws.Range("A21:J40"), ws2)
    Windows(1).SheetViews("まとめ").DisplayGridlines = False
End Sub
