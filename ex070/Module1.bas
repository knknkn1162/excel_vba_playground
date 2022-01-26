Option Explicit

Dim elapse As Integer

Sub UpdateStatusBar()
    Application.StatusBar = Format(Now(), "hh:mm:ss") & " " & (elapse-1) &"秒経過"
End Sub

Sub stopUpdate()
    Application.StatusBar = "時間表示終了"
End Sub

Sub main()
    elapse = elapse + 1
    Call UpdateStatusBar
    If elapse < 10 Then
        Call Application.OnTime(Now + TimeValue("00:00:01"), "main")
    Else
        Call Application.OnTime( _
            EarliestTime:=Now + TimeValue("00:00:01"), _
            Procedure:="stopUpdate" _
        )
    End If
End Sub
