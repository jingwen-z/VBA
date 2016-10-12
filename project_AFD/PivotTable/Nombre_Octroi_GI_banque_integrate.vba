Sub Octroi_GI_en_nombre()

    Dim wbkThis As Workbook
    Dim wbkOpen As Workbook

    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\Table_Principale_30-06-16_TdB.xlsm")
    
    wbkOpen.Worksheets("Feuil1").Range("A86:K94").Copy wbkThis.Worksheets("Feuil1").Range("B27")
    
    wbkThis.Worksheets("Feuil1").Range("B27").FormulaR1C1 = "Octroi GI (en nombre)"
    wbkThis.Worksheets("Feuil1").Range("B35").FormulaR1C1 = "Total"
    wbkThis.Worksheets("Feuil1").Range("K27").FormulaR1C1 = "2016 act."
    wbkThis.Worksheets("Feuil1").Range("L27").FormulaR1C1 = "Total"
    
    wbkOpen.Close False

End Sub
