Option Explicit


Function convert(ByRef arr() As Variant, ByVal seq As Integer) As String
    Dim i As Integer, j As Integer
    ' seq=1 vs seq=2 should be the same result
    seq = IIF(seq=1,2,seq)
    If seq <= 0 Then Exit Function
    Dim cnt As Integer: cnt = UBound(arr,1)
    Redim Preserve arr(cnt+1)
    Dim brr() As String
    Redim brr(cnt+1)
    Dim pos As Integer: pos = 0
    ' guard
    arr(cnt+1) = -2
    Dim step As Integer: step = 0
    Dim prev As Long: prev = -2

    For i = LBound(arr,1) To cnt+1
        If arr(i) = prev+step Then
            step = step+1
        Else
            If step >= seq Then
                brr(pos) = prev & "-" & (prev+step-1)
                pos = pos + 1
            Else
                ' the first `prev=-2` will be trashed because step=0
                For j = 1 To step
                    brr(pos) = prev+j-1
                    pos = pos + 1
                Next
            End If
            ' The last value arr(cnt+1)=-2 will be trashed
            step = 1
            prev = arr(i)
        End If
    Next
    Redim Preserve arr(cnt)
    Redim Preserve brr(pos-1)
    convert = Join(brr, ",")
End Function

Sub main()
    Dim arr() As Variant
    arr = Array(1,2,3,5,8,9,11,12,13,14,15,17,19,20,21,22)
    ' output to cell
    Dim i As Integer
    For i = 1 To 5
        Cells(i,1) = convert(arr, i)
    Next
End Sub
