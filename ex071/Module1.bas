Option Explicit

Sub main()
    Dim ws As Worksheet: Set ws = Worksheets(1)
    ws.ChartObjects(1).Chart.CopyPicture
    DoEvents

    Dim ppApp As New PowerPoint.Application
    Dim ppPt As Presentation
    Dim bdir As String: bdir = ThisWorkbook.Path & "/ex071"
    Set ppPt = ppApp.Presentations.Open(bdir & "/prezen1.pptx")
    Dim ppSlide As Slide: Set ppSlide = ppPt.Slides(1)
    Dim pptOrigShape As PowerPoint.Shape: Set pptOrigShape = ppSlide.Shapes(1)
    With ppSlide.Shapes.Paste
        .LockAspectRatio = msoTrue
        .Width = WorksheetFunction.min(pptOrigShape.Width, .Width)
        .Height = WorksheetFunction.min(pptOrigShape.Height, .Height)
        ' aligned center
        .Top = pptOrigShape.Top
        .Left = pptOrigShape.Left + (pptOrigShape.Width - .Width)/2
    End With
    Application.CutCopyMode = False
    pptOrigShape.Delete

    ppPt.SaveAs fileName:= bdir & "/out.pdf", FileFormat:=ppSaveAsPDF 
    ppApp.Quit
    Set ppSlide = Nothing
    Set ppPt = Nothing
    Set ppApp = Nothing
End Sub
