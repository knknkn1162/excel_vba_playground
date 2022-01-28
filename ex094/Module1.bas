Option Explicit

Function formatCode(str As String, depth As Integer) As String
    const unit As Integer = 2
    formatCode = Space(depth * unit) & str
    ' Debug.Print formatCode
End Function

Function formatThTd(tag As String, _
    rowspan As Integer, colspan As Integer, data As String, depth As Integer) As String
    Dim rowstr As String: rowstr = ""
    Dim colstr As String: colstr = ""
    Dim datastr As String: datastr = IIf(data="", "&nbsp;", data)
    Dim str As String: str = "<" & tag
    If rowspan >= 2 Then str = str & " rowspan=""" & rowspan & """"
    If colspan >= 2 Then str = str & " colspan=""" & colspan & """"
    str = str & ">"
    str = str & datastr
    str = str & "</" & tag & ">"
    formatThTd = formatCode(str, depth)
End Function

Function ParseThtd(ByRef rng As Range, ByRef arr As Variant, tag As String, pos As Integer, depth As Integer) As Integer
    Dim i As Integer
    arr(pos) = formatCode("<tr>", depth): pos = pos + 1
    depth = depth + 1
    For i = 1 To rng.Columns.Count
        Dim r As Range: Set r = rng.Columns(i)
        Dim mr As Range: Set mr = r.MergeArea
        If mr.Cells(1,1).Address = r.Address Then
            arr(pos) = formatthtd( _
                tag, _
                mr.Rows.Count, _
                mr.Columns.Count, _
                r.Value(), depth _
            )
            pos = pos + 1
        End If
    Next
    depth = depth - 1
    arr(pos) = formatCode("</tr>", depth): pos = pos + 1
    ParseThTd = pos
End Function

Function ParseThead(ByRef rng As Range, ByRef arr As Variant, pos As Integer, depth As Integer) As Integer
    arr(pos) = formatCode("<thead>", depth): pos = pos + 1
    depth = depth + 1
    Dim i As Integer
    For i = 1 To rng.Rows.Count
        pos = ParseThTd(rng.Rows(i), arr, "th", pos, depth)
    Next
    depth = depth - 1
    arr(pos) = formatCode("</thead>", depth): pos = pos + 1
    ParseThead = pos
End Function

Function ParseTbody(ByRef rng As Range, ByRef arr As Variant, pos As Integer, depth As Integer) As Integer
    arr(pos) = formatCode("<tbody>", depth): pos = pos + 1
    depth = depth + 1
    Dim i As Integer
    For i = 1 To rng.Rows.Count
        pos = ParseThTd(rng.Rows(i), arr, "td", pos, depth)
    Next
    depth = depth - 1
    arr(pos) = formatCode("</tbody>", depth): pos = pos + 1
    ParseTbody = pos
End Function

Function ParseTable(ByRef rng As Range, ByRef arr As Variant, pos As Integer, nh As Integer) As Integer
    Dim depth As Integer: depth = 0
    arr(pos) = formatCode("<table border=""1"">", depth): pos = pos + 1
    depth = depth + 1
    pos = ParseThead(rng.Resize(nh), arr, pos, depth)
    pos = ParseTbody(Intersect(rng, rng.Offset(nh)), arr, pos, depth)
    depth = depth - 1
    arr(pos) = formatCode("</table>", depth): pos = pos + 1
    ParseTable = pos
End Function

Function ConvertHTML(rng As Range, nh As Integer) As String
    Dim arr() As String
    Redim arr(100)
    Dim pos As Integer: pos = 0
    pos = ParseTable(rng, arr, pos, nh)
    Redim Preserve arr(pos-1)
    ConvertHTML = Join(arr, vbLf)
End Function

Sub main()
    Open Replace(Thisworkbook.FullName, ".xlsm", ".html") For Output As #1
    Dim str As String
    str = ConvertHTML(Range("B2").CurrentRegion, 2)
    Print #1, str
    Close #1
End Sub
