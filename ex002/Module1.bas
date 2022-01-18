Option Explicit

Sub main()
    Range("A1:C5").Copy
    With Worksheets("Sheet2")
        .Range("A1").PasteSpecial Paste:=xlPasteValues
        .Range("A1").PasteSpecial Paste:=xlPasteFormats
    End With
    Application.CutCopyMode = False
End Sub
