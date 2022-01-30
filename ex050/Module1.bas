Option Explicit

Sub main()
    Cells(1,1).Resize(3) = WorksheetFunction.transpose(Array(0,0,1))
    Dim i As Integer
    For i = 4 To 10000
        On Error Resume Next
        Cells(i,1) = Cells(i-3,1) + Cells(i-2,1) + Cells(i-1,1)
        If Err.Number <> 0 Then
            ' i = 1169
            '  Excel の場合、格納できる最大数は 1.79769313486232E+308 で、保存できる最小正数は 2.2250738555072E-308 
            Debug.Print i
            Exit Sub
        End If
    Next
End Sub
