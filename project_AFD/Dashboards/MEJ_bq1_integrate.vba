Sub MEJ_SGBCI()

    Dim wbkOpen2 As Workbook

    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\MEJ_30-06-16_TCD.xlsm")
    Set wbkOpen2 = Workbooks.Open(wbkThis.Path & "\Table_Principale_30-06-16_TCD.xlsm")

    wbkOpen.Worksheets("Feuil1").Range("X7:AB8").Copy wbkThis.Worksheets("Feuil1").Range("B85")
    wbkOpen2.Worksheets("Feuil1").Range("A156:D156").Copy wbkThis.Worksheets("Feuil1").Range("B88")
    wbkOpen2.Worksheets("Feuil1").Range("G156").Copy wbkThis.Worksheets("Feuil1").Range("F88")
    
    With wbkThis.Worksheets("Feuil1")
        .Range("C87").FormulaR1C1 = .Range("C86").Value / .Range("C88").Value
        .Range("D87").FormulaR1C1 = .Range("D86").Value / .Range("D88").Value
        .Range("E87").FormulaR1C1 = .Range("E86").Value / .Range("E88").Value
        .Range("F87").FormulaR1C1 = .Range("F86").Value / .Range("F88").Value
    
        .Range("B88:F88").Delete Shift:=xlToLeft
        .Range("C87:F87").NumberFormat = "0.00%"
        .Range("B86:F86").Font.Bold = False
    
        With .Range("B86:F86").Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With
    
    End With
    
    wbkOpen.Worksheets("Feuil1").Range("X16:AB16").Copy wbkThis.Worksheets("Feuil1").Range("B88")
    wbkOpen2.Worksheets("Feuil1").Range("A156:D156").Copy wbkThis.Worksheets("Feuil1").Range("B90")
    wbkOpen2.Worksheets("Feuil1").Range("G156").Copy wbkThis.Worksheets("Feuil1").Range("F90")
    
    With wbkThis.Worksheets("Feuil1")
        .Range("C89").FormulaR1C1 = .Range("C88").Value / .Range("C90").Value
        .Range("D89").FormulaR1C1 = .Range("D88").Value / .Range("D90").Value
        .Range("E89").FormulaR1C1 = .Range("E88").Value / .Range("E90").Value
        .Range("F89").FormulaR1C1 = .Range("F88").Value / .Range("F90").Value
    
        .Range("B90:F90").Delete Shift:=xlToLeft
        .Range("C89:F89").NumberFormat = "0.00%"
        .Range("B88:F88").Font.Bold = False
    
        With .Range("B88:F88").Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With
    
    End With
    
    wbkOpen.Worksheets("Feuil1").Range("X24:AB24").Copy wbkThis.Worksheets("Feuil1").Range("B90")
    wbkOpen2.Worksheets("Feuil1").Range("A156:D156").Copy wbkThis.Worksheets("Feuil1").Range("B92")
    wbkOpen2.Worksheets("Feuil1").Range("G156").Copy wbkThis.Worksheets("Feuil1").Range("F92")
    
    With wbkThis.Worksheets("Feuil1")
        .Range("C91").FormulaR1C1 = .Range("C90").Value / .Range("C92").Value
        .Range("D91").FormulaR1C1 = .Range("D90").Value / .Range("D92").Value
        .Range("E91").FormulaR1C1 = .Range("E90").Value / .Range("E92").Value
        .Range("F91").FormulaR1C1 = .Range("F90").Value / .Range("F92").Value
    
        .Range("B92:F92").Delete Shift:=xlToLeft
        .Range("C91:F91").NumberFormat = "0.00%"
        .Range("B90:F90").Font.Bold = False
    
        With .Range("B90:F90").Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With
    
    End With
    
    wbkOpen.Worksheets("Feuil1").Range("X35:AB35").Copy wbkThis.Worksheets("Feuil1").Range("B92")
    wbkOpen2.Worksheets("Feuil1").Range("A156:D156").Copy wbkThis.Worksheets("Feuil1").Range("B94")
    wbkOpen2.Worksheets("Feuil1").Range("G156").Copy wbkThis.Worksheets("Feuil1").Range("F94")
    
    With wbkThis.Worksheets("Feuil1")
        .Range("C93").FormulaR1C1 = .Range("C92").Value / .Range("C94").Value
        .Range("D93").FormulaR1C1 = .Range("D92").Value / .Range("D94").Value
        .Range("E93").FormulaR1C1 = .Range("E92").Value / .Range("E94").Value
        .Range("F93").FormulaR1C1 = .Range("F92").Value / .Range("F94").Value
    
        .Range("B94:F94").Delete Shift:=xlToLeft
        .Range("C93:F93").NumberFormat = "0.00%"
        .Range("B92:F92").Font.Bold = False
    
        With .Range("B92:F92").Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With
    
    End With
    
    With wbkThis.Worksheets("Feuil1")
        .Range("B85").FormulaR1C1 = "MEJ (en M€) SGBCI"
        .Range("B86").FormulaR1C1 = "montant d'engagement garanti"
        .Range("B87").FormulaR1C1 = "Taux de sinistralité 1"
        .Range("B88").FormulaR1C1 = "montant d'indemnisation max"
        .Range("B89").FormulaR1C1 = "Taux de sinistralité 2"
        .Range("B90").FormulaR1C1 = "montant d'indemnisation réel"
        .Range("B91").FormulaR1C1 = "Taux de sinistralité 3"
        .Range("B92").FormulaR1C1 = "perte provisoire calculée par la banque"
        .Range("B93").FormulaR1C1 = "Taux de sinistralité 4"

    With wbkThis.Worksheets("Feuil1").Range("B87:F87,B89:F89,B91:F91,B93:F93")
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

        '.Range("B87:G87").Borders(xlEdgeBottom).LineStyle = xlNone
        '.Range("B89:G89").Borders(xlEdgeBottom).LineStyle = xlNone
        
    End With
    
    wbkOpen.Close False
    wbkOpen2.Close False

End Sub
