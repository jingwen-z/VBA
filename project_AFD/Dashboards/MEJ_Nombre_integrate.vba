Sub MEJ_Nombre()

    Dim wbkThis As Workbook
    Dim wbkOpen As Workbook
    Dim wbkOpen2 As Workbook

    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\MEJ_30-06-16_TCD.xlsm")
    Set wbkOpen2 = Workbooks.Open(wbkThis.Path & "\Table_Principale_30-06-16_TCD.xlsm")
    
    wbkOpen.Worksheets("Feuil1").Range("AK7:AP8").Copy wbkThis.Worksheets("Feuil1").Range("B52")
    
    wbkThis.Worksheets("Feuil1").Range("B54:G54").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    wbkThis.Worksheets("Feuil1").Range("B54:G54").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    wbkOpen2.Worksheets("Feuil1").Range("A136:D136").Copy wbkThis.Worksheets("Feuil1").Range("B55")
    wbkOpen2.Worksheets("Feuil1").Range("G136").Copy wbkThis.Worksheets("Feuil1").Range("F55")
    
    With wbkThis.Worksheets("Feuil1")
        .Range("G55").FormulaR1C1 = "=SUM(RC[-4]:RC[-1])"
        
        .Range("C54").FormulaR1C1 = .Range("C53").Value / .Range("C55").Value
        .Range("D54").FormulaR1C1 = .Range("D53").Value / .Range("D55").Value
        .Range("E54").FormulaR1C1 = .Range("E53").Value / .Range("E55").Value
        .Range("F54").FormulaR1C1 = .Range("F53").Value / .Range("F55").Value
        .Range("G54").FormulaR1C1 = .Range("G53").Value / .Range("G55").Value
                
        .Range("B52").FormulaR1C1 = "MEJ (en nombre)GI"
        .Range("B53").FormulaR1C1 = "nb. de demande"
        .Range("B54").FormulaR1C1 = "Taux de sinistralité en nombre"
        .Range("G52").FormulaR1C1 = "Avant 2016"
        
        .Range("B55:G55").Delete Shift:=xlUp
        .Range("C54:G54").NumberFormat = "0.00%"
    End With

    wbkOpen.Close False
    wbkOpen2.Close False

End Sub
