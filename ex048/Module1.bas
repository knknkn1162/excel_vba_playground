Option Explicit

' True If type of v is Single or Double
Function IsDecimal(v As variant) As String
    Dim t As Integer: t = VarType(v)
    IsDecimal = (t = vbSingle) Or (t = vbDouble)
End Function

Sub transform1(ByRef v As Variant)
    Dim i As Integer
    For i = LBound(v,1) To Ubound(v,1)
        If IsDecimal(v(i)) Then
            v(i) = Fix(v(i))
        End If
    Next
End Sub

Sub transform2(ByRef v As Variant)
    Dim i,j As Integer
    For i = LBound(v,1) To Ubound(v,1)
        For j = LBound(v,2) To Ubound(v,2)
            If IsDecimal(v(i,j)) Then
                v(i,j) = Fix(v(i,j))
            End If
        Next
    Next
End Sub

Function transform(v As Variant) as Variant
    Dim tmp As Integer
    Dim dimension As Integer: dimension = 100
    Dim i As Integer
    For i = 1 To 3
        On Error Resume Next
        tmp = UBound(v,i)
        If Err.Number <> 0 Then
            dimension = i-1
            Err.Clear
            Exit For
        End If
    Next
    Debug.Print "dimension = " & dimension
    Select Case dimension
        Case 1
            Call transform1(v)
        Case 2
            Call transform2(v)
    End Select
    transform = v
End Function

Sub main()
    ' create Testcase
    Dim arr0 As Variant: arr0 = 3.5
    Dim arr1() As Variant: arr1 = Array(-1.5, 1.5, "1.5", #2020/1/1#)
    Dim arr2(1,3) As Variant
    Dim i As Integer, j As Integer
    For i = 0 To UBound(arr2,1)
        For j = 0 To UBound(arr2,2)
            arr2(i,j) = arr1(j)
        Next
    Next
    Dim brr1() As Variant, brr2() As Variant
    Dim brr0 As Variant
    brr0 = transform(arr0)
    brr1 = transform(arr1)
    brr2 = transform(arr2)
    Range("A1") = brr0
    Range("C1").Resize(,4) = brr1
    Range("G1").Resize(2,4) = brr2
End Sub
