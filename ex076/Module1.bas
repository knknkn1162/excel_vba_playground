Option Explicit

Sub btn_click(str As String)
    On Error Resume Next
    ' Precedence
    ' 1.GoTo Cells -> 2:GoTo Sheet -> 3.Do Nothing
    Application.GoTo Reference:=Range(str), Scroll:=True
    Sheets(str).Activate
    Err.Clear
End Sub

Sub main()
    If TypeName(Application.Caller) <> "String" Then Exit Sub
    Msgbox Application.Caller

    Dim btn As Button
    For Each btn in Worksheets(1).Buttons
        Dim addr As String: addr = btn.Caption
        btn.OnAction = "'btn_click(""" & btn.Caption & """)'"
    Next
End Sub
