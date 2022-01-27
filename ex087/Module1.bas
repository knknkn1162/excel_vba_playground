Option Explicit

Function ParseDecendents(rng As Range, ByRef arr As Variant) As Variant
    Dim i As Integer
    Dim ret() As String
    ReDim ret(LBound(arr) To UBound(arr))
    For i = LBound(arr) To UBound(arr)
        ret(i) = ""
        If rng.Worksheet.Name = arr(i) Then GoTo FCon
        Dim pat As String: pat = "*'" & Replace(arr(i), "'","''") & "'!*"
        If Not rng.Find(What:=pat, LookIn:=xlFormulas, LookAt:=xlPart) Is Nothing Then
            ret(i) = "〇"
        End If
FCon:
    Next
    ParseDecendents = ret
End Function

Function CollectFormulaCells(ws As Worksheet) As Range
    On Error Resume Next
    Set CollectFormulaCells = ws.Cells.SpecialCells(XlCellTypeFormulas)
    On Error GoTo 0
End Function

Sub main()
    Dim sz As Integer: sz = Worksheets.Count-1
    Dim origs() As String
    Dim arr() As String
    Redim orig(1 To sz)
    Redim arr(1 To sz)
    Dim i As Integer, j As Integer
    Dim tws As Worksheet: Set tws = Worksheets("相関表")
    tws.Move Before:=Worksheets(1)
    For i = 1 To sz
        orig(i) = Worksheets(i+1).Name
        arr(i) = orig(i) & vbTab
        Worksheets(i+1).Name = arr(i)
    Next

    Dim trng As Range: Set trng = tws.Range("B2")
    trng.CurrentRegion.Offset(1,1).ClearContents

    Dim r As Range
    For i = 1 To sz
        Dim ws As Worksheet: Set ws = Worksheets(arr(i))
        Dim rng As Range: Set rng = CollectFormulaCells(ws)
        Dim bs() As String
        ReDim bs(1 To sz)
        If Not rng Is Nothing Then
            bs = parseDecendents(rng, arr)
        End If
        trng.Offset(i,1).Resize(1,sz).Value() = bs
    Next
    For i = 1 To sz
        Worksheets(arr(i)).Name = orig(i)
    Next
End Sub
