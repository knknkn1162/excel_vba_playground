Option Explicit

Sub main()
    Dim ws1 As Worksheet: Set ws1 = Worksheets("元表1")
    Dim ws2 As Worksheet: Set ws2 = Worksheets("元表2")
    Dim mws As Worksheet: Set mws = Worksheets("まとめ")
    Dim shp1 As Shape, shp2 As Shape
    ws1.Range("A1").CurrentRegion.Copy
    mws.Pictures.Paste Link:=True
    Application.CutCopyMode = False
    ws2.Range("A1").CurrentRegion.Copy
    mws.Pictures.Paste Link:=True
    Application.CutCopyMode = False
    With mws.Pictures(2)
        .Top = Range("A21").Top
        .Left = Range("A21").Left
    End With
    Windows(1).SheetViews("まとめ").DisplayGridlines = False
End Sub
