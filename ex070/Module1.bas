Option Explicit

Sub main()
    Static dtStart As Date
    If dtStart = 0 Then dtStart = Now()
    If Now() <= dtStart + TimeSerial(0,0,60) Then
        Application.StatusBar = Format(Now(), "hh:mm:ss")
        Call Application.OnTime(Now + TimeValue("00:00:01"), "main")
    Else
        Application.StatusBar = "時間表示完了"
    End If
End Sub
