Sub MEJ_GI()

    Dim wbkThis As Workbook
    Dim wbkOpen As Workbook
    Dim wbkOpen2 As Workbook    

    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\MEJ_30-06-16_TdB.xlsm")
    Set wbkOpen2 = Workbooks.Open(wbkThis.Path & "\Table_Principale_30-06-16_TdB.xlsm")

    wbkOpen.Worksheets("Feuil1").Range("A7:F8").Copy wbkThis.Worksheets("Feuil1").Range("B57")
    wbkOpen2.Worksheets("Feuil1").Range("A7:D7").Copy wbkThis.Worksheets("Feuil1").Range("B60")
    wbkOpen2.Worksheets("Feuil1").Range("G7").Copy wbkThis.Worksheets("Feuil1").Range("F60")
    
    With wbkThis.Worksheets("Feuil1")
        .Range("G60").FormulaR1C1 = "=SUM(RC[-4]:RC[-1])"
        
        .Range("C59").FormulaR1C1 = .Range("C58").Value / .Range("C60").Value
        .Range("D59").FormulaR1C1 = .Range("D58").Value / .Range("D60").Value
        .Range("E59").FormulaR1C1 = .Range("E58").Value / .Range("E60").Value
        .Range("F59").FormulaR1C1 = .Range("F58").Value / .Range("F60").Value
        .Range("G59").FormulaR1C1 = .Range("G58").Value / .Range("G60").Value
    
        .Range("B60:G60").Delete Shift:=xlToLeft
        .Range("C59:G59").NumberFormat = "0.00%"
        .Range("B58:G58").Font.Bold = False
    
        With .Range("B58:G58").Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With
    
    End With
    
    wbkOpen.Worksheets("Feuil1").Range("A24:F24").Copy wbkThis.Worksheets("Feuil1").Range("B60")
    wbkOpen2.Worksheets("Feuil1").Range("A7:D7").Copy wbkThis.Worksheets("Feuil1").Range("B62")
    wbkOpen2.Worksheets("Feuil1").Range("G7").Copy wbkThis.Worksheets("Feuil1").Range("F62")
    
    With wbkThis.Worksheets("Feuil1")
        .Range("G62").FormulaR1C1 = "=SUM(RC[-4]:RC[-1])"
        
        .Range("C61").FormulaR1C1 = .Range("C60").Value / .Range("C62").Value
        .Range("D61").FormulaR1C1 = .Range("D60").Value / .Range("D62").Value
        .Range("E61").FormulaR1C1 = .Range("E60").Value / .Range("E62").Value
        .Range("F61").FormulaR1C1 = .Range("F60").Value / .Range("F62").Value
        .Range("G61").FormulaR1C1 = .Range("G60").Value / .Range("G62").Value
    
        .Range("B62:G62").Delete Shift:=xlToLeft
        .Range("C61:G61").NumberFormat = "0.00%"
        .Range("B60:G60").Font.Bold = False
    
        With .Range("B60:G60").Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With
    
    End With
    
    wbkOpen.Worksheets("Feuil1").Range("A40:F40").Copy wbkThis.Worksheets("Feuil1").Range("B62")
    wbkOpen2.Worksheets("Feuil1").Range("A7:D7").Copy wbkThis.Worksheets("Feuil1").Range("B64")
    wbkOpen2.Worksheets("Feuil1").Range("G7").Copy wbkThis.Worksheets("Feuil1").Range("F64")
    
    With wbkThis.Worksheets("Feuil1")
        .Range("G64").FormulaR1C1 = "=SUM(RC[-4]:RC[-1])"
        
        .Range("C63").FormulaR1C1 = .Range("C62").Value / .Range("C64").Value
        .Range("D63").FormulaR1C1 = .Range("D62").Value / .Range("D64").Value
        .Range("E63").FormulaR1C1 = .Range("E62").Value / .Range("E64").Value
        .Range("F63").FormulaR1C1 = .Range("F62").Value / .Range("F64").Value
        .Range("G63").FormulaR1C1 = .Range("G62").Value / .Range("G64").Value
    
        .Range("B64:G64").Delete Shift:=xlToLeft
        .Range("C63:G63").NumberFormat = "0.00%"
        .Range("B62:G62").Font.Bold = False
    
        With .Range("B62:G62").Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With
    
    End With

    With wbkThis.Worksheets("Feuil1")
        .Range("B57").FormulaR1C1 = "MEJ (en M€) GI"
        .Range("B58").FormulaR1C1 = "montant d'engagement garanti"
        .Range("B59").FormulaR1C1 = "Taux de sinistralité 1"
        .Range("B60").FormulaR1C1 = "montant d'indemnisation max"
        .Range("B61").FormulaR1C1 = "Taux de sinistralité 2"
        .Range("B62").FormulaR1C1 = "montant d'indemnisation réel"
        .Range("B63").FormulaR1C1 = "Taux de sinistralité 3"
        .Range("G57").FormulaR1C1 = "Avant 2016"
                
        .Range("B59:G59").Borders(xlEdgeBottom).LineStyle = xlNone
        .Range("B61:G61").Borders(xlEdgeBottom).LineStyle = xlNone
        
    End With
    
    wbkOpen.Close False
    wbkOpen2.Close False

End Sub
