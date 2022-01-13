Option Explicit

Sub main()
    With Range("B2").CurrentRegion.EntireColumn
        .FormatConditions.Delete
    End With

    With Intersect(Range("B2").CurrentRegion, Range("E:E, G:G")).FormatConditions
        With .Add( _
            Type:=xlCellValue, Operator:=xlLess, Formula1:="90%")
                .Interior.Color = vbRed
        End With
        With .Add( _
            Type:=xlCellValue, Operator:=xlLess, Formula1:="100%")
            .Font.Color = vbRed
        End With
    End With
End Sub
