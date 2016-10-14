Sub MEJ_montant_max_secteur()

    Dim wbkThis As Workbook
    Dim wbkOpen As Workbook

    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\MEJ_30-06-16_TdB.xlsm")
    
    wbkOpen.Worksheets("Feuil1").Range("AH44:AM59").Copy wbkThis.Worksheets("Feuil1").Range("B116")

    wbkOpen.Worksheets("Feuil1").Range("AH68:AM68").Copy
    wbkThis.Worksheets("Feuil1").Range("B118:G118").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("AH69:AM69").Copy
    wbkThis.Worksheets("Feuil1").Range("B120:G120").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("AH70:AM70").Copy
    wbkThis.Worksheets("Feuil1").Range("B122:G122").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("AH71:AM71").Copy
    wbkThis.Worksheets("Feuil1").Range("B124:G124").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("AH72:AM72").Copy
    wbkThis.Worksheets("Feuil1").Range("B126:G126").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("AH73:AM73").Copy
    wbkThis.Worksheets("Feuil1").Range("B128:G128").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("AH74:AM74").Copy
    wbkThis.Worksheets("Feuil1").Range("B130:G130").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("AH75:AM75").Copy
    wbkThis.Worksheets("Feuil1").Range("B132:G132").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("AH76:AM76").Copy
    wbkThis.Worksheets("Feuil1").Range("B134:G134").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("AH77:AM77").Copy
    wbkThis.Worksheets("Feuil1").Range("B136:G136").Insert Shift:=xlDown

    wbkOpen.Worksheets("Feuil1").Range("AH78:AM78").Copy
    wbkThis.Worksheets("Feuil1").Range("B138:G138").Insert Shift:=xlDown

    wbkOpen.Worksheets("Feuil1").Range("AH79:AM79").Copy
    wbkThis.Worksheets("Feuil1").Range("B140:G140").Insert Shift:=xlDown

    wbkOpen.Worksheets("Feuil1").Range("AH80:AM80").Copy
    wbkThis.Worksheets("Feuil1").Range("B142:G142").Insert Shift:=xlDown

    wbkOpen.Worksheets("Feuil1").Range("AH81:AM81").Copy
    wbkThis.Worksheets("Feuil1").Range("B144:G144").Insert Shift:=xlDown

    wbkOpen.Worksheets("Feuil1").Range("AH82:AM82").Copy
    wbkThis.Worksheets("Feuil1").Range("B146:G146").Insert Shift:=xlDown

    With wbkThis.Worksheets("Feuil1")
         .Range("C117").FormulaR1C1 = .Range("C117").Value / 1000000
         .Range("D117").FormulaR1C1 = .Range("D117").Value / 1000000
         .Range("E117").FormulaR1C1 = .Range("E117").Value / 1000000
         .Range("F117").FormulaR1C1 = .Range("F117").Value / 1000000
         .Range("G117").FormulaR1C1 = .Range("G117").Value / 1000000
         
         .Range("C119").FormulaR1C1 = .Range("C119").Value / 1000000
         .Range("D119").FormulaR1C1 = .Range("D119").Value / 1000000
         .Range("E119").FormulaR1C1 = .Range("E119").Value / 1000000
         .Range("F119").FormulaR1C1 = .Range("F119").Value / 1000000
         .Range("G119").FormulaR1C1 = .Range("G119").Value / 1000000
    
         .Range("C121").FormulaR1C1 = .Range("C121").Value / 1000000
         .Range("D121").FormulaR1C1 = .Range("D121").Value / 1000000
         .Range("E121").FormulaR1C1 = .Range("E121").Value / 1000000
         .Range("F121").FormulaR1C1 = .Range("F121").Value / 1000000
         .Range("G121").FormulaR1C1 = .Range("G121").Value / 1000000
    
         .Range("C123").FormulaR1C1 = .Range("C123").Value / 1000000
         .Range("D123").FormulaR1C1 = .Range("D123").Value / 1000000
         .Range("E123").FormulaR1C1 = .Range("E123").Value / 1000000
         .Range("F123").FormulaR1C1 = .Range("F123").Value / 1000000
         .Range("G123").FormulaR1C1 = .Range("G123").Value / 1000000
    
         .Range("C125").FormulaR1C1 = .Range("C125").Value / 1000000
         .Range("D125").FormulaR1C1 = .Range("D125").Value / 1000000
         .Range("E125").FormulaR1C1 = .Range("E125").Value / 1000000
         .Range("F125").FormulaR1C1 = .Range("F125").Value / 1000000
         .Range("G125").FormulaR1C1 = .Range("G125").Value / 1000000
    
         .Range("C127").FormulaR1C1 = .Range("C127").Value / 1000000
         .Range("D127").FormulaR1C1 = .Range("D127").Value / 1000000
         .Range("E127").FormulaR1C1 = .Range("E127").Value / 1000000
         .Range("F127").FormulaR1C1 = .Range("F127").Value / 1000000
         .Range("G127").FormulaR1C1 = .Range("G127").Value / 1000000
    
         .Range("C129").FormulaR1C1 = .Range("C129").Value / 1000000
         .Range("D129").FormulaR1C1 = .Range("D129").Value / 1000000
         .Range("E129").FormulaR1C1 = .Range("E129").Value / 1000000
         .Range("F129").FormulaR1C1 = .Range("F129").Value / 1000000
         .Range("G129").FormulaR1C1 = .Range("G129").Value / 1000000
    
         .Range("C131").FormulaR1C1 = .Range("C131").Value / 1000000
         .Range("D131").FormulaR1C1 = .Range("D131").Value / 1000000
         .Range("E131").FormulaR1C1 = .Range("E131").Value / 1000000
         .Range("F131").FormulaR1C1 = .Range("F131").Value / 1000000
         .Range("G131").FormulaR1C1 = .Range("G131").Value / 1000000
    
         .Range("C133").FormulaR1C1 = .Range("C133").Value / 1000000
         .Range("D133").FormulaR1C1 = .Range("D133").Value / 1000000
         .Range("E133").FormulaR1C1 = .Range("E133").Value / 1000000
         .Range("F133").FormulaR1C1 = .Range("F133").Value / 1000000
         .Range("G133").FormulaR1C1 = .Range("G133").Value / 1000000
    
         .Range("C135").FormulaR1C1 = .Range("C135").Value / 1000000
         .Range("D135").FormulaR1C1 = .Range("D135").Value / 1000000
         .Range("E135").FormulaR1C1 = .Range("E135").Value / 1000000
         .Range("F135").FormulaR1C1 = .Range("F135").Value / 1000000
         .Range("G135").FormulaR1C1 = .Range("G135").Value / 1000000
    
         .Range("C137").FormulaR1C1 = .Range("C137").Value / 1000000
         .Range("D137").FormulaR1C1 = .Range("D137").Value / 1000000
         .Range("E137").FormulaR1C1 = .Range("E137").Value / 1000000
         .Range("F137").FormulaR1C1 = .Range("F137").Value / 1000000
         .Range("G137").FormulaR1C1 = .Range("G137").Value / 1000000
    
         .Range("C139").FormulaR1C1 = .Range("C139").Value / 1000000
         .Range("D139").FormulaR1C1 = .Range("D139").Value / 1000000
         .Range("E139").FormulaR1C1 = .Range("E139").Value / 1000000
         .Range("F139").FormulaR1C1 = .Range("F139").Value / 1000000
         .Range("G139").FormulaR1C1 = .Range("G139").Value / 1000000
    
         .Range("C141").FormulaR1C1 = .Range("C141").Value / 1000000
         .Range("D141").FormulaR1C1 = .Range("D141").Value / 1000000
         .Range("E141").FormulaR1C1 = .Range("E141").Value / 1000000
         .Range("F141").FormulaR1C1 = .Range("F141").Value / 1000000
         .Range("G141").FormulaR1C1 = .Range("G141").Value / 1000000
    
         .Range("C143").FormulaR1C1 = .Range("C143").Value / 1000000
         .Range("D143").FormulaR1C1 = .Range("D143").Value / 1000000
         .Range("E143").FormulaR1C1 = .Range("E143").Value / 1000000
         .Range("F143").FormulaR1C1 = .Range("F143").Value / 1000000
         .Range("G143").FormulaR1C1 = .Range("G143").Value / 1000000
    
         .Range("C145").FormulaR1C1 = .Range("C145").Value / 1000000
         .Range("D145").FormulaR1C1 = .Range("D145").Value / 1000000
         .Range("E145").FormulaR1C1 = .Range("E145").Value / 1000000
         .Range("F145").FormulaR1C1 = .Range("F145").Value / 1000000
         .Range("G145").FormulaR1C1 = .Range("G145").Value / 1000000
    
         .Range("C117:G146").NumberFormat = "0.000"
        
         .Range("B116").FormulaR1C1 = "MEJ (en M€) montant max (GI)"
         .Range("G116").FormulaR1C1 = "Total"
         .Range("B118").FormulaR1C1 = "Taux de sinistralité"
         .Range("B120").FormulaR1C1 = "Taux de sinistralité"
         .Range("B122").FormulaR1C1 = "Taux de sinistralité"
         .Range("B124").FormulaR1C1 = "Taux de sinistralité"
         .Range("B126").FormulaR1C1 = "Taux de sinistralité"
         .Range("B128").FormulaR1C1 = "Taux de sinistralité"
         .Range("B130").FormulaR1C1 = "Taux de sinistralité"
         .Range("B132").FormulaR1C1 = "Taux de sinistralité"
         .Range("B134").FormulaR1C1 = "Taux de sinistralité"
         .Range("B136").FormulaR1C1 = "Taux de sinistralité"
         .Range("B138").FormulaR1C1 = "Taux de sinistralité"
         .Range("B140").FormulaR1C1 = "Taux de sinistralité"
         .Range("B142").FormulaR1C1 = "Taux de sinistralité"
         .Range("B144").FormulaR1C1 = "Taux de sinistralité"
         .Range("B146").FormulaR1C1 = "Taux de sinistralité"
    End With

    wbkOpen.Close False

End Sub
