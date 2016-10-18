Sub MEJ_GI_Nombre()

    Dim wbkThis As Workbook
    Dim wbkOpen As Workbook
    Dim wbkOpen2 As Workbook

    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\MEJ_30-06-16_TdB.xlsm")
    Set wbkOpen2 = Workbooks.Open(wbkThis.Path & "\Table_Principale_30-06-16_TdB.xlsm")
    
    wbkOpen.Worksheets("Feuil1").Range("Y7:AD8").Copy wbkThis.Worksheets("Feuil1").Range("B52")
    wbkOpen2.Worksheets("Feuil1").Range("A101:D101").Copy wbkThis.Worksheets("Feuil1").Range("B5")
    wbkOpen2.Worksheets("Feuil1").Range("G101").Copy wbkThis.Worksheets("Feuil1").Range("F55")
    
    With wbkThis.Worksheets("Feuil1")
        .Range("G55").FormulaR1C1 = "=SUM(RC[-4]:RC[-1])"        
        
        .Range("C54").FormulaR1C1 = .Range("C53").Value / .Range("C55").Value
        .Range("D54").FormulaR1C1 = .Range("D53").Value / .Range("D55").Value
        .Range("E54").FormulaR1C1 = .Range("E53").Value / .Range("E55").Value
        .Range("F54").FormulaR1C1 = .Range("F53").Value / .Range("F55").Value
        .Range("G54").FormulaR1C1 = .Range("G53").Value / .Range("G55").Value

        .Range("B52").FormulaR1C1 = "MEJ (en nombre) GI"
        .Range("B53").FormulaR1C1 = "nb. de demande"
        .Range("B54").FormulaR1C1 = "Taux de sinistralité en nombre"
        .Range("G52").FormulaR1C1 = "Avant 2016"

        .Range("B55:G55").Delete Shift:=xlToLeft
        .Range("C54:G54").NumberFormat = "0.00%"
        .Range("B53:G53").Font.Bold = False
        
        With .Range("B53:G53").Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
        End With

    End With

    wbkOpen.Close False
    wbkOpen2.Close False

End Sub
