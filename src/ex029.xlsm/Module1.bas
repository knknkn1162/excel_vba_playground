Option Explicit

Sub main()
    Dim rng As Range
    Set rng = Selection
    Dim d As String
    d = inputbox("画像のpath選択")
    Dim root As String: root = ThisWorkbook.Path
    Dim fpath As String
    fpath = root & "/" & d
    Dim shp As Shape
    Set shp = Activesheet.Shapes.AddPicture( _
        FileName := fpath, _
        LinkToFile := False, _
        SaveWithDocument := True, _
        Left := rng.Left, _
        Top := rng.Top, _
        Width:=-1, Height:=-1)
    shp.LockAspectRatio = msoTrue
    'shp.Placement = xlMoveAndSize

    Dim ratio As Double
    shp.Width = WorksheetFunction.min(rng.Width, shp.Width)
    shp.Height = WorksheetFunction.min(rng.Height, shp.Height)
    shp.Left = shp.Left + (rng.Width-shp.Width)/2
    shp.Top = shp.Top + (rng.Height-shp.Height)/2
End Sub
