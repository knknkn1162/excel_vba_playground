Option Explicit

Sub main()
    Worksheets("Sheet1").Range("A1:C5").Copy Destination:=Worksheets("Sheet2").Range("A1")
End Sub
