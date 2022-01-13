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

Sub main2()
    Dim i As Integer, tmp As String
    Dim str As String
    str = "あいうＡＢＣアイウａｂｃ１２３"
    str = UCase(str)
    Dim ans As String
    ans = ""
    For i = 1 To Len(str)
        tmp = Mid(str, i, 1)
        Select Case tmp
            Case StrConv("A", vbWide) To StrConv("Z", vbWide)
                Mid(str, i, 1) = StrConv(tmp, vbNarrow)
            Case StrConv("0", vbWide) To StrConv("9", vbWide)
                Mid(str, i, 1) = StrConv(tmp, vbNarrow)
        End Select
    Next
    Range("A1") = str
End Sub
