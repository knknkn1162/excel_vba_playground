Option Explicit

Sub ClearReceiptData()
    Dim ws As Worksheet: Set ws = Worksheets("請求書")
    With Application.FindFormat
        .Clear
        .Interior.Color = vbYellow
    End With
    Dim st As Range: Set st = Range("A1")
    Dim r As Range: Set r = ws.Cells.Find(What:="", After:=st, SearchFormat:=True)
    Set st = r
    Dim rng As Range: Set rng = st
    ' findall
    Do
        Set rng = Union(rng, r)
        Set r = ws.Cells.Find(What:="", After:=r, SearchFormat:=True)
    Loop While r.Address <> st.Address
    rng.ClearContents
End Sub

Sub main()
    Call ClearReceiptData
End Sub
