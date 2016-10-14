Sub MEJ_SGBCI()

    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\MEJ_30-06-16_TdB.xlsm")

    With wbkOpen.Worksheets("Feuil1")
        .Range("P7:T8").Copy wbkThis.Worksheets("Feuil1").Range("B77")
        .Range("P16:T16").Copy wbkThis.Worksheets("Feuil1").Range("B79")
        .Range("P24:T24").Copy wbkThis.Worksheets("Feuil1").Range("B80")
        .Range("P32:T32").Copy wbkThis.Worksheets("Feuil1").Range("B81")
        .Range("P40:T40").Copy wbkThis.Worksheets("Feuil1").Range("B82")
        .Range("P48:T48").Copy wbkThis.Worksheets("Feuil1").Range("B83")
    End With

    With wbkThis.Worksheets("Feuil1")
        .Range("B77").FormulaR1C1 = "MEJ (en M€) SGBCI"
        .Range("B78").FormulaR1C1 = "montant d'engagement garanti"
        .Range("B79").FormulaR1C1 = "Taux de sinistralité 1"
        .Range("B80").FormulaR1C1 = "montant d'indemnisation max"
        .Range("B81").FormulaR1C1 = "Taux de sinistralité 2"
        .Range("B82").FormulaR1C1 = "montant d'indemnisation réel"
        .Range("B83").FormulaR1C1 = "Taux de sinistralité 3"
        .Range("B78:F83").Font.Bold = False
        
        With .Range("B78:F83").Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With
        
        With .Range("B79:F82")
             .Borders(xlDiagonalDown).LineStyle = xlNone
             .Borders(xlDiagonalUp).LineStyle = xlNone
             .Borders(xlEdgeLeft).LineStyle = xlNone
             .Borders(xlEdgeTop).LineStyle = xlNone
             .Borders(xlEdgeBottom).LineStyle = xlNone
             .Borders(xlEdgeRight).LineStyle = xlNone
             .Borders(xlInsideVertical).LineStyle = xlNone
             .Borders(xlInsideHorizontal).LineStyle = xlNone
        End With
    
    End With
    
    wbkOpen.Close False

End Sub
