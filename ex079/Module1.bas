Option Explicit

Sub main()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet: Set ws = Worksheets(1)
    Dim wdApp As New Word.Application
    Dim wdDoc As Word.Document
    Dim wpath As String: wpath = ThisWorkbook.Path & "/ex079/doc1.docx"
    Set wdDoc = wdApp.Documents.Open(wpath)
    wdApp.Visible = True

    ' save as picture
    ws.Range("A1").CurrentRegion.Copy
    wdDoc.Content.InsertAfter Text:=wb.Name & vbCrLf
    wdDoc.Content.InsertAfter Text:=ws.Name & vbCrLf
    With wdApp.Selection
        .EndKey Unit:=wdStory
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
