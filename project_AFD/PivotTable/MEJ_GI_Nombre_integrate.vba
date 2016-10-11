Sub MEJ_GI_Nombre()

    Dim wbkThis As Workbook
    Dim wbkOpen As Workbook
    Dim wbkOpen2 As Workbook

    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\MEJ_30-06-16_TdB.xlsm")
    Set wbkOpen2 = Workbooks.Open(wbkThis.Path & "\Table_Principale_30-06-16_TdB.xlsm")
    
    wbkOpen.Worksheets("Feuil1").Range("Y7:AC8").Copy wbkThis.Worksheets("Feuil1").Range("B58")
    wbkOpen2.Worksheets("Feuil1").Range("A101:D101").Copy wbkThis.Worksheets("Feuil1").Range("B61")
    wbkOpen2.Worksheets("Feuil1").Range("G101").Copy wbkThis.Worksheets("Feuil1").Range("F61")
    
    With wbkThis.Worksheets("Feuil1")
        
        .Range("C60").FormulaR1C1 = .Range("C59").Value / .Range("C61").Value
        .Range("D60").FormulaR1C1 = .Range("D59").Value / .Range("D61").Value
        .Range("E60").FormulaR1C1 = .Range("E59").Value / .Range("E61").Value
        .Range("F60").FormulaR1C1 = .Range("F59").Value / .Range("F61").Value
                
        .Range("B58").FormulaR1C1 = "MEJ (en nombre) GI"
        .Range("B59").FormulaR1C1 = "nb. de demande"
        .Range("B60").FormulaR1C1 = "Taux de sinistralité en nombre"
        
        .Range("B61:F61").Delete Shift:=xlToLeft
        .Range("C60:F60").NumberFormat = "0.00"
        .Range("B59:F59").Font.Bold = False
        
        With .Range("B59:F59").Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

    End With

    wbkOpen.Close False
    wbkOpen2.Close False

End Sub
