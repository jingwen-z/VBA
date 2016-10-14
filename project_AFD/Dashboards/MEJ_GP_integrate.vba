Sub MEJ_GP()

    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\MEJ_30-06-16_TdB.xlsm")

    With wbkOpen.Worksheets("Feuil1")
        .Range("J7:L8").Copy wbkThis.Worksheets("Feuil1").Range("B72")
        .Range("J16:L16").Copy wbkThis.Worksheets("Feuil1").Range("B74")
    End With

    With wbkThis.Worksheets("Feuil1")
        .Range("B72").FormulaR1C1 = "MEJ (en M€) GP"
        .Range("B73").FormulaR1C1 = "montant d'indemnisation réel"
        .Range("B74").FormulaR1C1 = "taux de sinistralité GP"
        .Range("D72").FormulaR1C1 = "Avant 2016"
        .Range("B73:D74").Font.Bold = False
        
        With .Range("B73:D74").Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With
        
        With .Range("B74:D74")
             .Borders(xlEdgeTop).LineStyle = xlNone
        End With
    
    End With
    
    wbkOpen.Close False

End Sub
