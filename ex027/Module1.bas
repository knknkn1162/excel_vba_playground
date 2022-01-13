Option Explicit

Sub main()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim link As Hyperlink
    For each link In ws.Hyperlinks
        ' ※図は無視してください。
        If link.type <> msoHyperlinkRange Then
            GoTo Continue
        End If
        Dim rng As Range
        Set rng = link.Range
        rng.offset(,1) = link.Address
        link.Delete
Continue:
    Next
End Sub
