Option Explicit

Sub UpdateStatusBar()
    Application.StatusBar = Format(Now(), "hh:mm:ss") & " " & (dtStart-1) &"秒経過"
End Sub

Sub stopUpdate()
    Application.StatusBar = "時間表示終了"
End Sub

Sub main()
    Static dtStart As Date
    If dtStart = 0 Then dtStart = Now()
    Call UpdateStatusBar
    If Now() <= dtStart + DateSerial(0,0,10) Then
        Call Application.OnTime(Now + TimeValue("00:00:01"), "main")
    Else
        Application.StatusBar = False
    End If
End Sub
