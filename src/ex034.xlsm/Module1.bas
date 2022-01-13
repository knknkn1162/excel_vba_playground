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

Function formatMatrix(ByRef mat As Variant) As String
    Dim str As String
    str = ""
    Dim i As Integer, j As Integer
    For i = 1 To UBound(mat, 1)
        For j = 1 To UBound(mat, 2)
            str = str & mat(i,j) & ","
        Next
        str = str & vbLf
    Next
    ' remove comma and vbLf
    formatMatrix = Left(str, Len(str)-2)
End Function

Sub main()
    Dim arr() As Variant
    ' No need `Set`
    arr = Range("A1").CurrentRegion.Value
    ' CW
    Msgbox formatMatrix(transpose(arr, true))
    ' CCW
    Msgbox formatMatrix(transpose(arr, false))
End Sub
