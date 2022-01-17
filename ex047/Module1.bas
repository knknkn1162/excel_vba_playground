Option Explicit

Sub main()
    Dim w As Window
    For each w In Windows
        w.Activate
        With w
            ' ズームを85%
            .Zoom = 85
            ' 表示を標準 
            .View = xlNormalView
        End With
        For each sv In w.SheetViews
            If TypeName(sv) <> "WorksheetView" Then
                GoTo Continue
            End If
            sv.Sheet.Activate
            ' autoscroll=True
            Application.GoTo sv.Sheet.Range("A1"), True
            ' 枠線を非表示
            sv.DisplayGridlines = False
Continue:
        Next
    Next
    Dim sht As Worksheet
    For each sht In WorkSheets
        Application.PrintCommunication = False
        ' 印刷の向き「横」
        sht.PageSetup.Orientation = xlLandscape
        Application.PrintCommunication = True
    Next
End Sub
