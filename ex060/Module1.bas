Option Explicit

Sub main()
    Dim r As Range
    Dim reg_str As String: reg_str = "株式会社"
    Dim pos As Integer
    For Each r In WorkSheets(1).Cells.SpecialCells(xlCellTypeConstants)
        Dim str As String: str = r.Value
        Dim str1 As String: str1 = strConv(str, vbNarrow)
        pos = InStr(str1, ")")
        If pos <> 0 Then
            Select Case Replace(Left(str1, pos), " ","")
                Case "株)", "(株)"
                    str = reg_str & Mid(str, pos+1)
            End Select
        End If

        pos = InStrRev(str1, "(")
        If pos <> 0 Then
            Select Case Replace(Mid(str1, pos), " ", "")
                Case "(株", "(株)"
                    str = Left(str, pos-1) & reg_str
            End Select
        End If
        ' support special word
        ' ㈱
        str = Replace(str, ChrW(&H3231), reg_str)
        ' ㍿
        str = Replace(str, ChrW(&H337F), reg_str)
        r.Value = str
    Next
End Sub
