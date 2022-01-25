Option Explicit

Sub main()
    Dim mws As Worksheet: Set mws = Worksheets("都道府県")
    Dim lst As Range
    Set lst = mws.Range("A1").CurrentRegion
    Set lst = Intersect(lst, lst.Offset(1))
    Dim rng As Range: set rng = Range("F2")
    Range("F1") = "都道府県"
    With rng.Validation
        .Delete
        .Add _
            Type:=xlValidateList, _
            Operator:=xlBetween, _
            AlertStyle:=xlValidAlertStop, _
            Formula1:= "=都道府県!" & lst.Address
    End With
End Sub
