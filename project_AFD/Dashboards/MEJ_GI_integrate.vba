Sub MEJ_GI()

    Dim wbkThis As Workbook
    Dim wbkOpen As Workbook

    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\MEJ_30-06-16_TdB.xlsm")

    With wbkOpen.Worksheets("Feuil1")
        .Range("A7:F8").Copy wbkThis.Worksheets("Feuil1").Range("B63")
        .Range("A16:F16").Copy wbkThis.Worksheets("Feuil1").Range("B65")
        .Range("A24:F24").Copy wbkThis.Worksheets("Feuil1").Range("B66")
        .Range("A32:F32").Copy wbkThis.Worksheets("Feuil1").Range("B67")
        .Range("A40:F40").Copy wbkThis.Worksheets("Feuil1").Range("B68")
        .Range("A48:F48").Copy wbkThis.Worksheets("Feuil1").Range("B69")
    End With

    With wbkThis.Worksheets("Feuil1")
        .Range("B63").FormulaR1C1 = "MEJ (en M€) GI"
        .Range("B64").FormulaR1C1 = "montant d'engagement garanti"
        .Range("B65").FormulaR1C1 = "Taux de sinistralité 1"
        .Range("B66").FormulaR1C1 = "montant d'indemnisation max"
        .Range("B67").FormulaR1C1 = "Taux de sinistralité 2"
        .Range("B68").FormulaR1C1 = "montant d'indemnisation réel"
        .Range("B69").FormulaR1C1 = "Taux de sinistralité 3"
        .Range("G63").FormulaR1C1 = "Avant 2016"
        .Range("B64:F69").Font.Bold = False
        
        With .Range("B64:G69").Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With
        
        With .Range("B65:G68")
             .Borders(xlEdgeTop).LineStyle = xlNone
             .Borders(xlEdgeLeft).LineStyle = xlNone
             .Borders(xlEdgeRight).LineStyle = xlNone
             .Borders(xlEdgeBottom).LineStyle = xlNone
             .Borders(xlDiagonalUp).LineStyle = xlNone
             .Borders(xlDiagonalDown).LineStyle = xlNone
             .Borders(xlInsideVertical).LineStyle = xlNone
             .Borders(xlInsideHorizontal).LineStyle = xlNone
        End With
    
    End With
    
    wbkOpen.Close False

End Sub
