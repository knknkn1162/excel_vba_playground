Option Explicit

Function transpose(ByRef mat As Variant, ByVal w As Boolean) As Variant
    Dim arr2() As Variant
    Dim r,c As Integer
    r = UBound(mat, 1): c = UBound(mat, 2)
    ReDim arr2(1 To c, 1 To r)
    Dim i As Integer, j As Integer
    For i = 1 To c
        For j =1 To r
            If w Then
                arr2(i, j) = mat(r-j+1,i)
            Else
                arr2(i, j) = mat(j,c-i+1)
            End If
        Next
    Next
    transpose = arr2
End Function

Sub main()
    Dim arr() As Variant
    ' No need `Set`
    arr = Range("A1").CurrentRegion.Value
    ' CW
    Range("E1").Resize(UBound(arr, 2), UBound(arr, 1)).Value = transpose(arr, true)
    ' CCW
    Range("E6").Resize(UBound(arr, 2), UBound(arr, 1)).Value = transpose(arr, false)
End Sub
