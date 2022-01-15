Option Explicit

Function Calc(a1 As Integer, a2 As Integer, op As Integer) As Integer
    Dim res As Integer
    Select Case op
        Case 0
            res = a1+a2
        Case 1
            res = a1-a2
        Case 2
            res = a1*a2
        Case 3
            res = Int(a1/a2)
    End Select
    Calc = res
End Function

Sub main()
    Dim ops() As String
    ops = Split("+,-,*,/", ",")
    Dim ts As Integer
    ts = 10
    Dim cnt As Integer
    cnt = 0
    Dim i As Integer
    For i = 1 To ts
        Dim a1 As Integer, a2 As Integer
        Dim op As Integer
        Dim a As Integer
        Do While True
            a1 = WorksheetFunction.RandBetween(10,99)
            a2 = WorksheetFunction.RandBetween(2,a1)
            op = WorksheetFunction.RandBetween(0,3)
            a = Calc(a1,a2,op)
            If a <= 400 Then
                Exit Do
            End If
        Loop
        Dim str As String
        str = Inputbox( _
            "第"  & i & "問: " & a1 & ops(op) & a2 &"= ?")
        Dim ans As Integer
        If IsNumeric(str) Then
            If CInt(str) = a Then
                cnt = cnt + 1
            End If
        End If
    Next
    Msgbox cnt & "問解けた
End Sub
