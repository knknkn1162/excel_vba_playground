Option Explicit

Function GetHeader(elm As Selenium.WebElement) As Variant
    Dim idx As Integer
    Dim arr() As String
    Dim th As Selenium.WebElement
    For Each th In elm.FindElementsByTag("th")
        Redim Preserve arr(idx)
        arr(idx) = th.text
        idx = idx + 1
    Next
    GetHeader = arr
End Function

Sub main()
    Dim ws As WOrksheet: Set ws = Worksheets(1)
    ws.Name = "スクレイピング結果"
　　Dim Driver As New Selenium.WebDriver
　　Driver.Start "chrome", "https://excel-ubara.com/vba100sample"
    Driver.Get "/vba100list.html"
    Dim elm As Selenium.WebElement
    Dim tr As Selenium.WebElement
    Dim td As Selenium.WebElement
    Set elm = Driver.FindElementByXPath("//table/thead")
    ws.Range("A1").Resize(1,5) = GetHeader(elm)

    Set elm = Driver.FindElementByXPath("//table/tbody")
    Dim pos As Integer: pos = 2
    For Each tr In elm.FindElementsByTag("tr")
        Dim arr(1 To 5) As String
        Dim i As Integer
        For i = 1 To tr.FindElementsByTag("td").Count
            arr(i) = tr.FindElementsByTag("td").item(i).text
        Next
        ws.Cells(pos,1).Resize(1,5) = arr
        pos = pos + 1
    Next
    ws.UsedRange.EntireColumn.Autofit
End Sub

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
