Option Explicit

Function IsAlpha(str As String) As Integer
    Select Case str
        Case "A" To "Z", ",", "'"
            isAlpha = False
        Case Else
            isAlpha = True
    End Select
End Function

Function ReplaceIT(str As String) As String
    Dim str2 As String
    str2 = StrConv(str, vbUpperCase + vbNarrow)
    Dim pos As Integer: pos = 1
    Dim i As Integer
    Do While True
        pos = Instr(pos, str2, "IT")
        If pos = 0 Then Exit Do
        Dim flag As Boolean: flag = True
        Dim ch As String
        For i = pos-1 To 1 Step -1
            ch = Mid(str2,i,1)
            If ch = " " Then GoTo Continue1
            flag = flag And IsAlpha(ch)
            Exit For
Continue1:
        Next
        For i = pos+2 To Len(str2)
            ch = Mid(str2,i,1)
            If ch = " " Then GoTo Continue2
            flag = flag And IsAlpha(ch)
            Exit For
Continue2:
        Next
        If flag Then Mid(str, pos, 2) = "DX"
        pos = pos + 2
    Loop
    ReplaceIT = str
End Function

Sub main()
    Dim v As Variant
    Dim pos As Integer: pos = 1
    ' test
    For Each v In Array("ＩＴ","と IT","itは","IT 99", "ＧＩＴ","site","It's","it is")
        Dim str As String: str = v
        Cells(pos, 1) = ReplaceIT(str)
        pos = pos + 1
    Next

End Sub
