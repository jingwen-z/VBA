Sub Octroi_GI_et_GP()
    
    Dim wbkThis As Workbook
    Dim wbkOpen As Workbook
    
    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\Table_Principale_30-06-16_TdB.xlsm")
    
    wbkOpen.Worksheets("Feuil1").Range("A6:K9").Copy wbkThis.Worksheets("Feuil1").Range("B4")
    wbkOpen.Worksheets("Feuil1").Range("B14:B17").Copy wbkThis.Worksheets("Feuil1").Range("M4")
    
    wbkThis.Worksheets("Feuil1").Range("B4").FormulaR1C1 = "Octroi (en M€) GI et GP"
    wbkThis.Worksheets("Feuil1").Range("B6").FormulaR1C1 = "GP"
    wbkThis.Worksheets("Feuil1").Range("K4").FormulaR1C1 = "2016 act."
    wbkThis.Worksheets("Feuil1").Range("L4").FormulaR1C1 = "Total"
    wbkThis.Worksheets("Feuil1").Range("M4").FormulaR1C1 = "Encours act."
    
    wbkOpen.Close False
    
End Sub
