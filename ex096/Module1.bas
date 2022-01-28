Option Explicit

Sub main()
　　Dim ws As Worksheet: Set ws = Worksheets(1)
　　Dim sDb As String
　　
　　Const cnsDate As Date = #1/1/2021#
　　Const cnsAmount As Long = 1000000
　　
　　sDb = ThisWorkbook.Path & "/ex096/DB1.accdb"
　　Call VBA100_96_ADO(sDb, ws, Array(cnsDate, cnsAmount))
End Sub

Sub VBA100_96_ADO(ByVal aDb As String, ws As Worksheet, ByRef aParam)
　　Dim adoCn As New ADODB.Connection
　　Dim adoRs As ADODB.Recordset
　　Dim isExcel As Boolean
　　
　　Set adoCn = getConnection(aDb, isExcel)
　　adoCn.Open aDb
　　Set adoRs = adoCn.Execute(createSql(aParam, isExcel))
　　
　　Call outputSheet(ws, adoRs)
　　
　　adoRs.Close: Set adoRs = Nothing
　　adoCn.Close: Set adoCn = Nothing
End Sub

Function getConnection(ByVal aDb As String, ByRef isExcel As Boolean) As ADODB.Connection
　　Dim adoCn As New ADODB.Connection
　　adoCn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0"
　　Select Case Mid(aDb, InStrRev(aDb, ".") + 1)
　　　　Case "accdb"
　　　　　　isExcel = False
　　　　Case "xlsx", "xlsm"
　　　　　　adoCn.Properties("Extended Properties") = "Excel 12.0"
　　　　　　isExcel = True
　　End Select
　　Set getConnection = adoCn
End Function

Sub outputSheet(ByVal ws As Worksheet, adoRs As ADODB.Recordset)
　　Dim i As Long
　　With ws
　　　　.Cells.Clear
　　　　For i = 0 To adoRs.Fields.Count - 1
　　　　　　.Cells(1, i + 1) = adoRs.Fields(i).Name
　　　　Next
　　　　.Range("A2").CopyFromRecordset adoRs
　　　　.Columns("E").NumberFormatLocal = "yyyy/mm/dd"
　　　　.Columns("F:H").NumberFormatLocal = "#,##0"
　　　　.Range("A1").CurrentRegion.EntireColumn.AutoFit
　　End With
End Sub

Function createSql(ByRef aParam, Optional ByVal isExcel As Boolean = False) As String
　　Dim sql() As String: ReDim sql(0)

    sqlAppend sql, "SELECT"
    sqlAppend sql, " T1.取引先CD"
    sqlAppend sql, ",M1.取引先名"
    sqlAppend sql, ",T1.商品CD"
    sqlAppend sql, ",M2.商品名"
    sqlAppend sql, ",T1.日付"
    sqlAppend sql, ",T1.単価"
    sqlAppend sql, ",T1.数量"
    sqlAppend sql, ",T1.数量 * T1.単価 AS 金額"
    sqlAppend sql, " FROM (([T売上] As T1"
    sqlAppend sql, " LEFT JOIN [M取引先] AS M1 ON T1.取引先CD = M1.取引先CD)"
    sqlAppend sql, " LEFT JOIN [M商品] AS M2 ON T1.商品CD = M2.商品CD)"
    sqlAppend sql, " WHERE T1.日付 >= #" & Format(aParam(0), "yyyy/mm/dd") & "#"
    sqlAppend sql, " AND T1.数量 * T1.単価 >= " & aParam(1)
　　
　　createSql = Join(sql)
　　
　　If isExcel Then
　　　　createSql = Replace(createSql, "[T売上]", "[T売上$]")
　　　　createSql = Replace(createSql, "[M取引先]", "[M取引先$]")
　　　　createSql = Replace(createSql, "[M商品]", "[M商品$]")
　　End If
End Function

Sub sqlAppend(ByRef sql, ByVal aString As String)
　　ReDim Preserve sql(1 To UBound(sql) + 1)
　　sql(UBound(sql)) = aString & vbCrLf
End Sub
