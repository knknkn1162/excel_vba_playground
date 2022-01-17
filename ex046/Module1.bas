Option Explicit

Function Normalize(str As String) As String
    ' See the link, https://docs.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.clean
    str = WorksheetFunction.Clean(str)
    Dim i As Integer
    For i = 1 To Len(str)
        Dim ch As String: ch = Mid(str,i,1)
        'Msgbox ch & "," & StrConv(ch, vbNarrow) & "," & StrConv(ch, vbHiragana)
        ' detect kanji and hiragana
        If ch = StrConv(ch, vbNarrow) And ch = StrConv(ch,vbWide) Then
            GoTo Continue
        End If
        ' detect katakana
        If ch <> StrConv(ch, vbHiragana) Then
            GoTo Continue
        End If
        ' detect alphabet and 0-9
        Select Case StrConv(ch, vbNarrow)
            Case "0" To "9"
            Case "A" To "Z"
            Case "a" To "z"
            Case StrConv("ー", vbNarrow)
            Case "-"
            ' " !""#$%&'()*+,-./:;<=>?@[\]^`{|}~"
            Case Else
                Mid(str,i,1) = "_"
                GoTo Continue
        End Select
Continue:
    Next
    Normalize = str
End Function

Sub main()
    Dim col As Integer
    Dim row As Integer
    For col=1 To Range("A1").CurrentRegion.Columns.Count
        Dim rng As Range
        Set rng = Range(Cells(1,col), Cells(Rows.Count, col).End(xlUp))
        Dim n As String
        n = Normalize(rng.Resize(1,1).Value)
        On Error Resume Next
        Names.Add Name:=n, RefersTo:=rng.Address
        Dim flag As Boolean: flag = True
        If Err.Number <> 0 Then
            Err.Clear
            On Error Resume Next
            Names.Add Name:="_"&n, RefersTo:=rng.Address
            If Err.Number <> 0 Then
                Debug.Print "回避不可な文字列です:" & n
                flag = False
            End If
        End If
    Next
    If flag <> True Then Msgbox "名前定義に登録できない名前がありました。"
End Sub
