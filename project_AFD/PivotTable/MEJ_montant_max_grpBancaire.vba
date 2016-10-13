Sub MEJ_montant_max_grpBancaire()

    Dim wbkThis As Workbook
    Dim wbkOpen As Workbook
    Dim Rng As Range
    
    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\MEJ_30-06-16_TdB.xlsm")
    
    With wbkOpen.Worksheets("Feuil1")
         .Range("AH7:AM8").Copy wbkThis.Worksheets("Feuil1").Range("B104")
         .Range("AH16:AM16").Copy wbkThis.Worksheets("Feuil1").Range("B106")
    End With
    
    With wbkThis.Worksheets("Feuil1")
         .Range("C105").FormulaR1C1 = .Range("C105").Value / 1000000
         .Range("D105").FormulaR1C1 = .Range("D105").Value / 1000000
         .Range("E105").FormulaR1C1 = .Range("E105").Value / 1000000
         .Range("F105").FormulaR1C1 = .Range("F105").Value / 1000000
         .Range("G105").FormulaR1C1 = .Range("G105").Value / 1000000
    
         .Range("C105:G106").NumberFormat = "0.00"
        
         .Range("B104").FormulaR1C1 = "MEJ (en M€) montant max"
         .Range("B106").FormulaR1C1 = "Taux de sinistralité"
         .Range("G104").FormulaR1C1 = "Total"
    End With
        
    wbkOpen.Close False
    
End Sub
