Option Explicit

Sub main()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet: Set ws = Worksheets(1)
    Dim wdApp As New Word.Application
    Dim wdDoc As Word.Document
    Dim wpath As String: wpath = ThisWorkbook.Path & "/ex079/doc1.docx"
    Set wdDoc = wdApp.Documents.Open(wpath)
    wdDoc.Bookmarks("エクセル表").Select

    ' save as picture
    ws.Range("A1").CurrentRegion.Copy
    With wdApp.Selection
        .TypeText(wb.Name & vbCrLf)
        .TypeText(ws.Name & vbCrLf)
        .PasteSpecial DataType:=wdPasteMetafilePicture
    End With
    Application.CutCopyMode = False

    'save as pdf
    wdDoc.ExportAsFixedFormat _
        OutputFileName:= Replace(wpath, ".docx", ".pdf"), _
        ExportFormat:=wdExportFormatPDF
    wdDoc.Close SaveChanges:=False
    wdApp.Quit
End Sub
