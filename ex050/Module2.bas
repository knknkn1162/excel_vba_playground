Option Explicit

Sub main2()
    Dim i As Long
    Dim ntmp
    Dim n1,n2,n3
    n1 = 0: Cells(1,1) = "'" & n1
    n2 = 0: Cells(2,1) = "'" & n2
    n3 = 1: Cells(3,1) = "'" & n3
    For i = 4 To 10000
        On Error GoTo ErrExit
        ntmp = CDec(n1+n2+n3)
        Cells(i,1) = "'" & ntmp
        n1 = n2: n2 = n3: n3 = ntmp
    Next
ErrExit:
    Columns(1).AutoFit
    Debug.Print i
End Sub
