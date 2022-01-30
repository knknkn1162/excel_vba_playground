Option Explicit

Function changeStr(str As String) As String
    Dim res As String
    Dim ch As String
    res = ""
    Dim i As Integer
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        If Not ch Like "[ア-ン]" Then
            ch = StrConv(ch, vbUpperCase + vbNarrow)
        End If
        res = res + ch
Continue:
    Next
    changeStr = res
End Function

Sub main()
    Range("A1") = "あいうＡＢＣアイウａｂｃ１２３"
    Range("A1") = changeStr(Range("A1"))
End Sub
