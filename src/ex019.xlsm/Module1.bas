Option Explicit

Sub main()
    Dim sp As Shape
    Dim ws As Worksheet
    Set ws = WorkSheets(1)
    For Each sp In ws.Shapes
        If sp.Type = msoFormControl Or sp.Type = msoOLEControlObject Then
            GoTo Continue
        End If
        If sp.Name = "checked" Then
            GoTo Continue
        End If
        sp.Name = "checked"
        With sp.Duplicate
            .left = sp.left + sp.width
            .top = sp.top
            .Name = "checked"
        End With
Continue:

    Next

End Sub
