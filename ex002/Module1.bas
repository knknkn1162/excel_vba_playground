Option Explicit

Sub main()
    Range("A1:C5").Copy
    Dim ws2 As Worksheet
    Set ws2 = Worksheets("Sheet2")
    ws2.Range("A1").PasteSpecial Paste:=xlPasteValues
    ws2.Range("A1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub
