Sub MEJ_montant_max_nature()
    
    Dim wbkOpen2 As Workbook
    
    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\MEJ_30-06-16_TdB.xlsm")
    Set wbkOpen2 = Workbooks.Open(wbkThis.Path & "\Table_Principale_30-06-16_TdB.xlsm")
    
    wbkOpen.Worksheets("Feuil1").Range("AX21:BB23").Copy wbkThis.Worksheets("Feuil1").Range("B107")

    With wbkThis.Worksheets("Feuil1")
         .Range("C108").FormulaR1C1 = .Range("C108").Value / 1000000
         .Range("D108").FormulaR1C1 = .Range("D108").Value / 1000000
         .Range("E108").FormulaR1C1 = .Range("E108").Value / 1000000
         .Range("F108").FormulaR1C1 = .Range("F108").Value / 1000000
         
         .Range("C109").FormulaR1C1 = .Range("C109").Value / 1000000
         .Range("D109").FormulaR1C1 = .Range("D109").Value / 1000000
         .Range("E109").FormulaR1C1 = .Range("E109").Value / 1000000
         .Range("F109").FormulaR1C1 = .Range("F109").Value / 1000000
    
         .Range("C108:F109").NumberFormat = "0.00"
         
         .Range("B109:F109").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("B109:F109").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         
    End With
    
    wbkOpen2.Worksheets("Feuil1").Range("A196:D196").Copy wbkThis.Worksheets("Feuil1").Range("B110")
    wbkOpen2.Worksheets("Feuil1").Range("G196").Copy wbkThis.Worksheets("Feuil1").Range("F110")

    With wbkThis.Worksheets("Feuil1")
         .Range("C109").FormulaR1C1 = 0
         .Range("D109").FormulaR1C1 = .Range("D108").Value / .Range("D110").Value
         .Range("E109").FormulaR1C1 = .Range("E108").Value / .Range("E110").Value
         .Range("F109").FormulaR1C1 = .Range("F108").Value / .Range("F110").Value
    
         .Range("B110:F110").Delete Shift:=xlUp
         .Range("C109:F109").NumberFormat = "0.00%"
        
         .Range("B107").FormulaR1C1 = "MEJ (en M€) montant max"
         .Range("B109").FormulaR1C1 = "Taux de sinistralité"
         
         .Range("B111:F111").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("B111:F111").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
    
    wbkOpen2.Worksheets("Feuil1").Range("A198:D198").Copy wbkThis.Worksheets("Feuil1").Range("B112")
    wbkOpen2.Worksheets("Feuil1").Range("G198").Copy wbkThis.Worksheets("Feuil1").Range("F112")

    With wbkThis.Worksheets("Feuil1")
         .Range("C111").FormulaR1C1 = .Range("C110").Value / .Range("C112").Value
         .Range("D111").FormulaR1C1 = .Range("D110").Value / .Range("D112").Value
         .Range("E111").FormulaR1C1 = .Range("E110").Value / .Range("E112").Value
         .Range("F111").FormulaR1C1 = .Range("F110").Value / .Range("F112").Value
    
         .Range("B112:F112").Delete Shift:=xlUp
         .Range("C111:F111").NumberFormat = "0.00%"
        
         .Range("B111").FormulaR1C1 = "Taux de sinistralité"
    End With

    With wbkThis.Worksheets("Feuil1").Range("B109:F109,B111:F111")
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

    wbkOpen.Close False
    wbkOpen2.Close False

End Sub
