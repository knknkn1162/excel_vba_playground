Option Explicit

Function zlookup(str As String, rng As Range, idx As Long, ord As Long) As Variant
    Dim firstrow As Integer: firstrow = rng.Resize(1,1).Row
    Dim lastrow As Integer: lastrow = rng.Resize(1,1).Row + rng.Rows.Count - 1
    Dim st As Integer, en As Integer, sp As Integer
    If ord = -1 Then
        st = lastrow: en = firstrow: sp = -1
    Else
        st = firstrow: en = lastrow: sp = 1
    End If

    Dim i As Integer
    Dim v As Variant
    If ord = 0 Then ord = 1
    Dim pos As Integer: pos = 0
    For i = firstrow To lastrow
        If rng.Cells(i, 1) = str Then
            pos = pos + 1
            If pos = ord Then
                v = rng.Cells(i,idx)
                Exit For
            End If
        End If
    Next
    If IsEmpty(v) And pos = 0 Then v = CVErr(XlErrNA)
    If IsEmpty(v) Then v = ""
    zlookup = v
End Function

Sub main()
    Dim n1 As Variant: n1 = zlookup("sample20", Range("A1").CurrentRegion, 3, 3)
    Dim n2 As Variant: n2 = zlookup("sample50", Range("A1").CurrentRegion, 3, 3)
    Dim n3 As Variant: n3 = zlookup("sample20", Range("A1").CurrentRegion, 3, -1)
    Dim n4 As Variant: n4 = zlookup("sample20", Range("A1").CurrentRegion, 3, 100)
    Range("E1").Resize(1,4).Value = Array(n1,n2,n3,n4)
End Sub
