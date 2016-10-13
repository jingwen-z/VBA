Sub MEJ_montant_max_nature()

    Dim wbkThis As Workbook
    Dim wbkOpen As Workbook
    
    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\MEJ_30-06-16_TdB.xlsm")
    
    wbkOpen.Worksheets("Feuil1").Range("AH24:AM26").Copy wbkThis.Worksheets("Feuil1").Range("B109")

    wbkOpen.Worksheets("Feuil1").Range("AH35:AM35").Copy
    wbkThis.Worksheets("Feuil1").Range("B111:G111").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("AH36:AM36").Copy
    wbkThis.Worksheets("Feuil1").Range("B113:G113").Insert Shift:=xlDown

    With wbkThis.Worksheets("Feuil1")
         .Range("C110").FormulaR1C1 = .Range("C110").Value / 1000000
         .Range("D110").FormulaR1C1 = .Range("D110").Value / 1000000
         .Range("E110").FormulaR1C1 = .Range("E110").Value / 1000000
         .Range("F110").FormulaR1C1 = .Range("F110").Value / 1000000
         .Range("G110").FormulaR1C1 = .Range("G110").Value / 1000000
         
         .Range("C112").FormulaR1C1 = .Range("C112").Value / 1000000
         .Range("D112").FormulaR1C1 = .Range("D112").Value / 1000000
         .Range("E112").FormulaR1C1 = .Range("E112").Value / 1000000
         .Range("F112").FormulaR1C1 = .Range("F112").Value / 1000000
         .Range("G112").FormulaR1C1 = .Range("G112").Value / 1000000
    
         .Range("C110:G113").NumberFormat = "0.00"
        
         .Range("B109").FormulaR1C1 = "MEJ (en M€) montant max (GI)"
         .Range("B111").FormulaR1C1 = "Taux de sinistralité"
         .Range("B113").FormulaR1C1 = "Taux de sinistralité"
         .Range("G109").FormulaR1C1 = "Total"
    End With

    wbkOpen.Close False

End Sub
