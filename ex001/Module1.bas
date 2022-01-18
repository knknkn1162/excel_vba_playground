Option Explicit

Sub main()
    Worksheets("Sheet1").Range("A1").CurrentRegion.Copy _
        Destination:=Worksheets("Sheet2").Range("A1")
End Sub
