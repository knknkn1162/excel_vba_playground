Option Explicit

Sub main()
    Dim rng As Range
    Set rng = ActiveSheet.Range("A1")

    Dim sh As Object
    Dim formula As String
    formula = ""
    For Each sh In ThisWorkbook.Sheets
        formula = formula & sh.Name & ","
    Next
    formula = Left(formula, Len(formula)-1)
    rng.NumberFormatLocal = "@"
    With rng.Validation
        .Delete
        .Add _
            Type:= xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Formula1:=formula
        .ErrorTitle = "エラー発生"
        .ShowError=True
        .ErrorMessage="シート名が無効です"
    End With
End Sub
