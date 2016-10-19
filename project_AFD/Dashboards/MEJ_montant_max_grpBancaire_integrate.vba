Sub MEJ_montant_max_grpBancaire()
    
    Dim wbkOpen2 As Workbook
    
    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\MEJ_30-06-16_TdB.xlsm")
    Set wbkOpen2 = Workbooks.Open(wbkThis.Path & "\Table_Principale_30-06-16_TdB.xlsm")
    
    wbkOpen.Worksheets("Feuil1").Range("AX6:BB7").Copy wbkThis.Worksheets("Feuil1").Range("B102")
    
    With wbkThis.Worksheets("Feuil1")
         .Range("C103").FormulaR1C1 = .Range("C103").Value / 1000000
         .Range("D103").FormulaR1C1 = .Range("D103").Value / 1000000
         .Range("E103").FormulaR1C1 = .Range("E103").Value / 1000000
         .Range("F103").FormulaR1C1 = .Range("F103").Value / 1000000
    
         .Range("C103:F103").NumberFormat = "0.00"
    End With
    
    wbkOpen2.Worksheets("Feuil1").Range("A177:D177").Copy wbkThis.Worksheets("Feuil1").Range("B105")
    wbkOpen2.Worksheets("Feuil1").Range("G177").Copy wbkThis.Worksheets("Feuil1").Range("F105")
    
    With wbkThis.Worksheets("Feuil1")
         .Range("C104").FormulaR1C1 = .Range("C103").Value / .Range("C105").Value
         .Range("D104").FormulaR1C1 = .Range("D103").Value / .Range("D105").Value
         .Range("E104").FormulaR1C1 = .Range("E103").Value / .Range("E105").Value
         .Range("F104").FormulaR1C1 = .Range("F103").Value / .Range("F105").Value
    
         .Range("B105:F105").Delete Shift:=xlToLeft
         .Range("C104:F105").NumberFormat = "0.00%"
        
         .Range("B102").FormulaR1C1 = "MEJ (en M€) montant max"
         .Range("B104").FormulaR1C1 = "Taux de sinistralité"
         
         With .Range("B104:F104")
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 5
                .TintAndShade = 0.399914548173467
                .Weight = xlThin
            End With
         
         End With
         
    End With
    
    wbkOpen.Close False
    wbkOpen2.Close False
    
End Sub
