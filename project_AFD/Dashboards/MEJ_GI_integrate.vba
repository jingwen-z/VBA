Sub MEJ_GI()

    Dim wbkThis As Workbook
    Dim wbkOpen As Workbook
    Dim wbkOpen2 As Workbook    

    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\MEJ_30-06-16_TCD.xlsm")
    Set wbkOpen2 = Workbooks.Open(wbkThis.Path & "\Table_Principale_30-06-16_TCD.xlsm")

    wbkOpen.Worksheets("Feuil1").Range("A7:F8").Copy wbkThis.Worksheets("Feuil1").Range("B61")
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
    
    wbkOpen.Worksheets("Feuil1").Range("A16:F16").Copy wbkThis.Worksheets("Feuil1").Range("B64")
    wbkOpen2.Worksheets("Feuil1").Range("A7:D7").Copy wbkThis.Worksheets("Feuil1").Range("B66")
    wbkOpen2.Worksheets("Feuil1").Range("G7").Copy wbkThis.Worksheets("Feuil1").Range("F66")
    
    With wbkThis.Worksheets("Feuil1")
        .Range("G66").FormulaR1C1 = "=SUM(RC[-4]:RC[-1])"
        
        .Range("C65").FormulaR1C1 = .Range("C64").Value / .Range("C66").Value
        .Range("D65").FormulaR1C1 = .Range("D64").Value / .Range("D66").Value
        .Range("E65").FormulaR1C1 = .Range("E64").Value / .Range("E66").Value
        .Range("F65").FormulaR1C1 = .Range("F64").Value / .Range("F66").Value
        .Range("G65").FormulaR1C1 = .Range("G64").Value / .Range("G66").Value
    
        .Range("B66:G66").Delete Shift:=xlToLeft
        .Range("C65:G65").NumberFormat = "0.00%"
        .Range("B64:G64").Font.Bold = False
    
        With .Range("B64:G64").Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With
    
    End With
    
    wbkOpen.Worksheets("Feuil1").Range("A24:F24").Copy wbkThis.Worksheets("Feuil1").Range("B66")
    wbkOpen2.Worksheets("Feuil1").Range("A7:D7").Copy wbkThis.Worksheets("Feuil1").Range("B68")
    wbkOpen2.Worksheets("Feuil1").Range("G7").Copy wbkThis.Worksheets("Feuil1").Range("F68")
    
    With wbkThis.Worksheets("Feuil1")
        .Range("G68").FormulaR1C1 = "=SUM(RC[-4]:RC[-1])"
        
        .Range("C67").FormulaR1C1 = .Range("C66").Value / .Range("C68").Value
        .Range("D67").FormulaR1C1 = .Range("D66").Value / .Range("D68").Value
        .Range("E67").FormulaR1C1 = .Range("E66").Value / .Range("E68").Value
        .Range("F67").FormulaR1C1 = .Range("F66").Value / .Range("F68").Value
        .Range("G67").FormulaR1C1 = .Range("G66").Value / .Range("G68").Value
    
        .Range("B68:G68").Delete Shift:=xlToLeft
        .Range("C67:G67").NumberFormat = "0.00%"
        .Range("B66:G66").Font.Bold = False
    
        With .Range("B66:G66").Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With
    
    End With
    
    wbkOpen.Worksheets("Feuil1").Range("A35:F35").Copy wbkThis.Worksheets("Feuil1").Range("B68")
    wbkOpen2.Worksheets("Feuil1").Range("A7:D7").Copy wbkThis.Worksheets("Feuil1").Range("B70")
    wbkOpen2.Worksheets("Feuil1").Range("G7").Copy wbkThis.Worksheets("Feuil1").Range("F70")
    
    With wbkThis.Worksheets("Feuil1")
        .Range("G70").FormulaR1C1 = "=SUM(RC[-4]:RC[-1])"
        
        .Range("C69").FormulaR1C1 = .Range("C68").Value / .Range("C70").Value
        .Range("D69").FormulaR1C1 = .Range("D68").Value / .Range("D70").Value
        .Range("E69").FormulaR1C1 = .Range("E68").Value / .Range("E70").Value
        .Range("F69").FormulaR1C1 = .Range("F68").Value / .Range("F70").Value
        .Range("G69").FormulaR1C1 = .Range("G68").Value / .Range("G70").Value
    
        .Range("B70:G70").Delete Shift:=xlToLeft
        .Range("C69:G69").NumberFormat = "0.00%"
        .Range("B68:G68").Font.Bold = False
    
        With .Range("B68:G68").Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With
    
    End With
    
    With wbkThis.Worksheets("Feuil1")
        .Range("B61").FormulaR1C1 = "MEJ (en M€) GI"
        .Range("B62").FormulaR1C1 = "montant d'engagement garanti"
        .Range("B63").FormulaR1C1 = "Taux de sinistralité 1"
        .Range("B64").FormulaR1C1 = "montant d'indemnisation max"
        .Range("B65").FormulaR1C1 = "Taux de sinistralité 2"
        .Range("B66").FormulaR1C1 = "montant d'indemnisation réel"
        .Range("B67").FormulaR1C1 = "Taux de sinistralité 3"
        .Range("B68").FormulaR1C1 = "perte provisoire calculée par la banque"
        .Range("B69").FormulaR1C1 = "Taux de sinistralité 4"
        .Range("G61").FormulaR1C1 = "Avant 2016"
    End With

    With wbkThis.Worksheets("Feuil1").Range("B63:G63,B65:G65,B67:G67,B69:G69")
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
