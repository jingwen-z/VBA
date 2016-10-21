Sub MEJ_Nombre()

    Dim wbkThis As Workbook
    Dim wbkOpen As Workbook
    Dim wbkOpen2 As Workbook
    Dim wbkOpen3 As Workbook

    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\MEJ_30-06-16_TCD.xlsm")
    Set wbkOpen2 = Workbooks.Open(wbkThis.Path & "\Table_Principale_30-06-16_TCD.xlsm")
    Set wbkOpen3 = Workbooks.Open(wbkThis.Path & "\GPP_31-12-15_TCD.xlsm")
    
    wbkOpen.Worksheets("Feuil1").Range("AK7:AP9").Copy wbkThis.Worksheets("Feuil1").Range("B52")
    
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
                
        .Range("B52").FormulaR1C1 = "MEJ (en nombre)"
        .Range("B53").FormulaR1C1 = "nb. de demande(GI)"
        .Range("B54").FormulaR1C1 = "Taux de sinistralité en nombre(GI)"
        .Range("G52").FormulaR1C1 = "Avant 2016"
        
        .Range("B55:G55").Delete Shift:=xlUp
        .Range("C54:G54").NumberFormat = "0.00%"

        .Range("B56:G56").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range("B56:G56").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End With

    wbkOpen3.Worksheets("Feuil1").Range("A51").Copy wbkThis.Worksheets("Feuil1").Range("B57")
    wbkOpen3.Worksheets("Feuil1").Range("B51:C51").Copy wbkThis.Worksheets("Feuil1").Range("D57")
    wbkOpen3.Worksheets("Feuil1").Range("E51").Copy wbkThis.Worksheets("Feuil1").Range("F57")

    With wbkThis.Worksheets("Feuil1")
        .Range("G57").FormulaR1C1 = "=SUM(RC[-4]:RC[-1])"
        
        .Range("C56").FormulaR1C1 = 0
        .Range("D56").FormulaR1C1 = .Range("D55").Value / .Range("D57").Value
        .Range("E56").FormulaR1C1 = 1 / 2
        .Range("F56").FormulaR1C1 = .Range("F55").Value / .Range("F57").Value
        .Range("G56").FormulaR1C1 = .Range("G55").Value / .Range("G57").Value
                
        .Range("B55").FormulaR1C1 = "nb. de demande(GP)"
        .Range("B56").FormulaR1C1 = "Taux de sinistralité en nombre(GP)"
        
        .Range("B57:G57").Delete Shift:=xlUp
        .Range("C56:G56").NumberFormat = "0.00%"
    End With

    wbkOpen.Close False
    wbkOpen2.Close False
    wbkOpen3.Close False

End Sub
