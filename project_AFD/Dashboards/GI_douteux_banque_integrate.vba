Sub GI_douteux()

    Dim wbkThis As Workbook
    Dim wbkOpen As Workbook

    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\GI_douteux_31-03-16_TdB.xlsm")
    
    wbkOpen.Worksheets("Feuil1").Range("A7:D9").Copy wbkThis.Worksheets("Feuil1").Range("B38")
    
    wbkThis.Worksheets("Feuil1").Range("B38").FormulaR1C1 = "GI_douteux (en M€)"
    wbkThis.Worksheets("Feuil1").Range("C38").FormulaR1C1 = "montant des garantie douteux"
    wbkThis.Worksheets("Feuil1").Range("D38").FormulaR1C1 = "encours"
    wbkThis.Worksheets("Feuil1").Range("E38").FormulaR1C1 = "provision"

    wbkOpen.Close False
    
End Sub
