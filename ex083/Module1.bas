Option Explicit

Function NewReceipt(shtName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Worksheets(shtName).Delete
    On Error GoTo 0

    Worksheets("請求書").Copy After:=Worksheets(Worksheets.Count)
    Set ws = Worksheets(Worksheets.Count)
    ws.Name = shtName
    With Application.FindFormat
        .Clear
        .Interior.Color = vbYellow
    End With
    Dim st As Range: Set st = ws.Range("A1")
    Dim r As Range: Set r = ws.Cells.Find(What:="", After:=st, SearchFormat:=True)
    Set st = r
    Dim rng As Range: Set rng = st
    ' findall
    Do
        Set rng = Union(rng, r)
        Set r = ws.Cells.Find(What:="", After:=r, SearchFormat:=True)
    Loop While r.Address <> st.Address
    rng.ClearContents
    rng.Interior.ColorIndex = xlNone
    Set NewReceipt = ws
End Function

Sub main()
    Const maxcnt As Integer = 10
    Dim uws As Worksheet: Set uws = Worksheets("売上")
    Dim mws As Worksheet: Set mws = Worksheets("取引先マスタ")
    Dim rng As Range: Set rng = mws.Range("A1").CurrentRegion
    Set rng = Intersect(rng, rng.Offset(1))
    Dim r As Range

    Dim rdir As String: rdir = ThisWorkbook.Path & "/ex083"
    On Error Resume Next
    RmDir rdir
    MkDir rdir
    On Error GoTo 0

    For Each r In rng.Columns(2).Cells
        Dim ws As Worksheet: Set ws = NewReceipt(r.Value & "_請求書")
        ws.Range("A2").Resize(4,1) = WorksheetFunction.transpose(r.Resize(1,4))
        uws.AutoFilterMode = False
        uws.Range("A1").Autofilter Field:=2, Criteria1:=r.Value()
        Dim rng2 As Range: Set rng2 = uws.Range("A1").CurrentRegion.Offset(1)
        rng2.Columns("C").Copy
        ws.Cells(10,1).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        rng2.Columns("D:E").Copy
        ws.Cells(10,3).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        uws.AutoFilterMode = False
        Dim fname As String: fname = r.Value() & "_" & Format(Now(), "yyyymm") & ".pdf"
        fname = Replace(fname, "株式会社", "")
        ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:= rdir & "/" & fname
    Next
End Sub
