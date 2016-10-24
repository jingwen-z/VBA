Sub MEJ_SGBCI()

    Dim wbkOpen2 As Workbook

    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\MEJ_30-06-16_TCD.xlsm")
    Set wbkOpen2 = Workbooks.Open(wbkThis.Path & "\GPP_31-12-15_TCD.xlsm")

    wbkOpen.Worksheets("Feuil1").Range("N7:P8").Copy wbkThis.Worksheets("Feuil1").Range("B73")
    wbkOpen2.Worksheets("Feuil1").Range("C59").Copy wbkThis.Worksheets("Feuil1").Range("C76")
    
    With wbkThis.Worksheets("Feuil1")
        .Range("D76").FormulaR1C1 = "=SUM(RC[-1]:RC[-1])"
        
        .Range("C75").FormulaR1C1 = .Range("C74").Value / .Range("C76").Value
        .Range("D75").FormulaR1C1 = .Range("D74").Value / .Range("D76").Value
    
        .Range("B76:D76").Delete Shift:=xlToLeft
        .Range("C75:D75").NumberFormat = "0.00%"
        .Range("B74:D74").Font.Bold = False
    
        With .Range("B74:D74").Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With
    
    End With
    
    wbkOpen.Worksheets("Feuil1").Range("N16:P16").Copy wbkThis.Worksheets("Feuil1").Range("B76")
    wbkOpen2.Worksheets("Feuil1").Range("C59").Copy wbkThis.Worksheets("Feuil1").Range("C78")
    
    With wbkThis.Worksheets("Feuil1")
        .Range("D78").FormulaR1C1 = "=SUM(RC[-1]:RC[-1])"
        
        .Range("C77").FormulaR1C1 = .Range("C76").Value / .Range("C78").Value
        .Range("D77").FormulaR1C1 = .Range("D76").Value / .Range("D78").Value
    
        .Range("B78:D78").Delete Shift:=xlToLeft
        .Range("C77:D77").NumberFormat = "0.00%"
        .Range("B76:D76").Font.Bold = False
    
        With .Range("B76:D76").Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With
    
    End With
    
    wbkOpen.Worksheets("Feuil1").Range("N24:P24").Copy wbkThis.Worksheets("Feuil1").Range("B78")
    wbkOpen2.Worksheets("Feuil1").Range("C59").Copy wbkThis.Worksheets("Feuil1").Range("C80")
    
    With wbkThis.Worksheets("Feuil1")
        .Range("D80").FormulaR1C1 = "=SUM(RC[-1]:RC[-1])"
        
        .Range("C79").FormulaR1C1 = .Range("C78").Value / .Range("C80").Value
        .Range("D79").FormulaR1C1 = .Range("D78").Value / .Range("D80").Value
    
        .Range("B80:D80").Delete Shift:=xlToLeft
        .Range("C79:D79").NumberFormat = "0.00%"
        .Range("B78:D78").Font.Bold = False
    
        With .Range("B78:D78").Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With
    
    End With
    
    wbkOpen.Worksheets("Feuil1").Range("N35:P35").Copy wbkThis.Worksheets("Feuil1").Range("B80")
    wbkOpen2.Worksheets("Feuil1").Range("C59").Copy wbkThis.Worksheets("Feuil1").Range("C82")
    
    With wbkThis.Worksheets("Feuil1")
        .Range("D82").FormulaR1C1 = "=SUM(RC[-1]:RC[-1])"
        
        .Range("C81").FormulaR1C1 = .Range("C80").Value / .Range("C82").Value
        .Range("D81").FormulaR1C1 = .Range("D80").Value / .Range("D82").Value
    
        .Range("B82:D82").Delete Shift:=xlToLeft
        .Range("C81:D81").NumberFormat = "0.00%"
        .Range("B80:D80").Font.Bold = False
    
        With .Range("B80:D80").Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With
    
    End With
    
    With wbkThis.Worksheets("Feuil1")
        .Range("B73").FormulaR1C1 = "MEJ (en M€) GP"
        .Range("B74").FormulaR1C1 = "montant d'engagement garanti"
        .Range("B75").FormulaR1C1 = "Taux de sinistralité 1"
        .Range("B76").FormulaR1C1 = "montant d'indemnisation max"
        .Range("B77").FormulaR1C1 = "Taux de sinistralité 2"
        .Range("B78").FormulaR1C1 = "montant d'indemnisation réel"
        .Range("B79").FormulaR1C1 = "Taux de sinistralité 3"
        .Range("B80").FormulaR1C1 = "perte provisoire calculée par la banque"
        .Range("B81").FormulaR1C1 = "Taux de sinistralité 4"
        .Range("D73").FormulaR1C1 = "Avant 2016"
                
        .Range("B75:D75").Borders(xlEdgeBottom).LineStyle = xlNone
        .Range("B77:D77").Borders(xlEdgeBottom).LineStyle = xlNone
    End With

    With wbkThis.Worksheets("Feuil1").Range("B75:D75,B77:D77,B79:D79,B81:D81")
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
