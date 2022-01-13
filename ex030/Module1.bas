Option Explicit

Sub main()
    Dim meibo_ws As Worksheet, nafuda_ws As Worksheet
    Set meibo_ws = Worksheets("名簿")
    Set nafuda_ws = Worksheets("名札")
    meibo_ws.Activate
    
    Dim pos As Integer, i As Integer
    pos = 1
    With nafuda_ws
        For i = 2 To Cells(Rows.Count, 1).End(xlup).Row Step 2
            .Rows("1:2").Copy
            .Rows(pos).PasteSpecial Paste:=xlPasteFormats
            .Cells(pos, 1) = Cells(i, 2)
            .Cells(pos+1, 1) = Cells(i, 3)
            .Cells(pos, 2) = Cells(i+1, 2)
            .Cells(pos+1, 2) = Cells(i+1, 3)
            pos = pos + 2
        Next
    End With

End Sub
