Option Explicit

Sub main2()
　　Const cnsURL = "https://excel-ubara.com/vba100sample/vba100list.html"
　　
　　Dim wb As Workbook: Set wb = ActiveWorkbook
　　Dim ws As Worksheet: Set ws = wb.ActiveSheet
　　ws.Cells.Clear
　　
　　With ws.QueryTables.Add(Connection:="URL;" & cnsURL, Destination:=ws.Range("A1"))
　　　　.FieldNames = True
　　　　.WebSelectionType = xlAllTables
　　　　.WebFormatting = xlWebFormattingNone
　　　　.Refresh BackgroundQuery:=False
　　　　.Delete
　　End With
End Sub
