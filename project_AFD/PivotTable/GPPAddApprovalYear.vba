Sub add_année_autorisation()

    With ThisWorkbook.Worksheets("BDD_GPP")
        .Columns("D:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range("D3").FormulaR1C1 = "Année d'autorisation"
        .Range("D4:D83").FormulaR1C1 = "=YEAR(RC[1])"
        .Range("D4:D83").NumberFormat = "General"
    End With
    
End Sub
