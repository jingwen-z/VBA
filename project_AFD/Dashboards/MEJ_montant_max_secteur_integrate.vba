Sub MEJ_montant_max_secteur()
    
    Dim wbkOpen2 As Workbook
    
    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\MEJ_30-06-16_TdB.xlsm")
    Set wbkOpen2 = Workbooks.Open(wbkThis.Path & "\Table_Principale_30-06-16_TdB.xlsm")
    
    wbkOpen.Worksheets("Feuil1").Range("AX36:BB51").Copy wbkThis.Worksheets("Feuil1").Range("B114")


    With wbkThis.Worksheets("Feuil1")
         .Range("C115").FormulaR1C1 = .Range("C115").Value / 1000000
         .Range("D115").FormulaR1C1 = .Range("D115").Value / 1000000
         .Range("E115").FormulaR1C1 = .Range("E115").Value / 1000000
         .Range("F115").FormulaR1C1 = .Range("F115").Value / 1000000
         
         .Range("C116").FormulaR1C1 = .Range("C116").Value / 1000000
         .Range("D116").FormulaR1C1 = .Range("D116").Value / 1000000
         .Range("E116").FormulaR1C1 = .Range("E116").Value / 1000000
         .Range("F116").FormulaR1C1 = .Range("F116").Value / 1000000
        
         .Range("C117").FormulaR1C1 = .Range("C117").Value / 1000000
         .Range("D117").FormulaR1C1 = .Range("D117").Value / 1000000
         .Range("E117").FormulaR1C1 = .Range("E117").Value / 1000000
         .Range("F117").FormulaR1C1 = .Range("F117").Value / 1000000
    
         .Range("C118").FormulaR1C1 = .Range("C118").Value / 1000000
         .Range("D118").FormulaR1C1 = .Range("D118").Value / 1000000
         .Range("E118").FormulaR1C1 = .Range("E118").Value / 1000000
         .Range("F118").FormulaR1C1 = .Range("F118").Value / 1000000
    
         .Range("C119").FormulaR1C1 = .Range("C119").Value / 1000000
         .Range("D119").FormulaR1C1 = .Range("D119").Value / 1000000
         .Range("E119").FormulaR1C1 = .Range("E119").Value / 1000000
         .Range("F119").FormulaR1C1 = .Range("F119").Value / 1000000
    
         .Range("C120").FormulaR1C1 = .Range("C120").Value / 1000000
         .Range("D120").FormulaR1C1 = .Range("D120").Value / 1000000
         .Range("E120").FormulaR1C1 = .Range("E120").Value / 1000000
         .Range("F120").FormulaR1C1 = .Range("F120").Value / 1000000
    
         .Range("C121").FormulaR1C1 = .Range("C121").Value / 1000000
         .Range("D121").FormulaR1C1 = .Range("D121").Value / 1000000
         .Range("E121").FormulaR1C1 = .Range("E121").Value / 1000000
         .Range("F121").FormulaR1C1 = .Range("F121").Value / 1000000
    
         .Range("C122").FormulaR1C1 = .Range("C122").Value / 1000000
         .Range("D122").FormulaR1C1 = .Range("D122").Value / 1000000
         .Range("E122").FormulaR1C1 = .Range("E122").Value / 1000000
         .Range("F122").FormulaR1C1 = .Range("F122").Value / 1000000
    
         .Range("C123").FormulaR1C1 = .Range("C123").Value / 1000000
         .Range("D123").FormulaR1C1 = .Range("D123").Value / 1000000
         .Range("E123").FormulaR1C1 = .Range("E123").Value / 1000000
         .Range("F123").FormulaR1C1 = .Range("F123").Value / 1000000
    
         .Range("C124").FormulaR1C1 = .Range("C124").Value / 1000000
         .Range("D124").FormulaR1C1 = .Range("D124").Value / 1000000
         .Range("E124").FormulaR1C1 = .Range("E124").Value / 1000000
         .Range("F124").FormulaR1C1 = .Range("F124").Value / 1000000
    
         .Range("C125").FormulaR1C1 = .Range("C125").Value / 1000000
         .Range("D125").FormulaR1C1 = .Range("D125").Value / 1000000
         .Range("E125").FormulaR1C1 = .Range("E125").Value / 1000000
         .Range("F125").FormulaR1C1 = .Range("F125").Value / 1000000
    
         .Range("C126").FormulaR1C1 = .Range("C126").Value / 1000000
         .Range("D126").FormulaR1C1 = .Range("D126").Value / 1000000
         .Range("E126").FormulaR1C1 = .Range("E126").Value / 1000000
         .Range("F126").FormulaR1C1 = .Range("F126").Value / 1000000
    
         .Range("C127").FormulaR1C1 = .Range("C127").Value / 1000000
         .Range("D127").FormulaR1C1 = .Range("D127").Value / 1000000
         .Range("E127").FormulaR1C1 = .Range("E127").Value / 1000000
         .Range("F127").FormulaR1C1 = .Range("F127").Value / 1000000
    
         .Range("C128").FormulaR1C1 = .Range("C128").Value / 1000000
         .Range("D128").FormulaR1C1 = .Range("D128").Value / 1000000
         .Range("E128").FormulaR1C1 = .Range("E128").Value / 1000000
         .Range("F128").FormulaR1C1 = .Range("F128").Value / 1000000
    
         .Range("C129").FormulaR1C1 = .Range("C129").Value / 1000000
         .Range("D129").FormulaR1C1 = .Range("D129").Value / 1000000
         .Range("E129").FormulaR1C1 = .Range("E129").Value / 1000000
         .Range("F129").FormulaR1C1 = .Range("F129").Value / 1000000
    
         .Range("C115:F129").NumberFormat = "0.000"
         
         .Range("B116:F116").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("B116:F116").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
        
    wbkOpen2.Worksheets("Feuil1").Range("A218:D218").Copy wbkThis.Worksheets("Feuil1").Range("B117")
    wbkOpen2.Worksheets("Feuil1").Range("G218").Copy wbkThis.Worksheets("Feuil1").Range("F117")
        
    With wbkThis.Worksheets("Feuil1")
         .Range("C116").FormulaR1C1 = 0
         .Range("D116").FormulaR1C1 = .Range("D115").Value / .Range("D117").Value
         .Range("E116").FormulaR1C1 = .Range("E115").Value / .Range("E117").Value
         .Range("F116").FormulaR1C1 = .Range("F115").Value / .Range("F117").Value
    
         .Range("B117:F117").Delete Shift:=xlUp
         .Range("C116:F116").NumberFormat = "0.00%"
        
         .Range("B114").FormulaR1C1 = "MEJ (en M€) montant max"
         .Range("B116").FormulaR1C1 = "Taux de sinistralité"
         
         .Range("B118:F118").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("B118:F118").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
        
    wbkOpen2.Worksheets("Feuil1").Range("A225:D225").Copy wbkThis.Worksheets("Feuil1").Range("B119")
    wbkOpen2.Worksheets("Feuil1").Range("G225").Copy wbkThis.Worksheets("Feuil1").Range("F119")
        
    With wbkThis.Worksheets("Feuil1")
         .Range("C118").FormulaR1C1 = .Range("C117").Value / .Range("C119").Value
         .Range("D118").FormulaR1C1 = .Range("D117").Value / .Range("D119").Value
         .Range("E118").FormulaR1C1 = 0
         .Range("F118").FormulaR1C1 = 0
    
         .Range("B119:F119").Delete Shift:=xlUp
         .Range("C118:F118").NumberFormat = "0.00%"
        
         .Range("B118").FormulaR1C1 = "Taux de sinistralité"
         
         .Range("B120:F120").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("B120:F120").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
        
    wbkOpen2.Worksheets("Feuil1").Range("A226:D226").Copy wbkThis.Worksheets("Feuil1").Range("B121")
    wbkOpen2.Worksheets("Feuil1").Range("G226").Copy wbkThis.Worksheets("Feuil1").Range("F121")
        
    With wbkThis.Worksheets("Feuil1")
         .Range("C120").FormulaR1C1 = 0
         .Range("D120").FormulaR1C1 = 0
         .Range("E120").FormulaR1C1 = .Range("E119").Value / .Range("E121").Value
         .Range("F120").FormulaR1C1 = 0
    
         .Range("B121:F121").Delete Shift:=xlUp
         .Range("C120:F120").NumberFormat = "0.00%"
        
         .Range("B120").FormulaR1C1 = "Taux de sinistralité"
         
         .Range("B122:F122").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("B122:F122").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
        
    wbkOpen2.Worksheets("Feuil1").Range("A235:D235").Copy wbkThis.Worksheets("Feuil1").Range("B123")
    wbkOpen2.Worksheets("Feuil1").Range("G235").Copy wbkThis.Worksheets("Feuil1").Range("F123")
        
    With wbkThis.Worksheets("Feuil1")
         .Range("C122").FormulaR1C1 = 0
         .Range("D122").FormulaR1C1 = .Range("D121").Value / .Range("D123").Value
         .Range("E122").FormulaR1C1 = 0
         .Range("F122").FormulaR1C1 = 0
    
         .Range("B123:F123").Delete Shift:=xlUp
         .Range("C122:F122").NumberFormat = "0.00%"
        
         .Range("B122").FormulaR1C1 = "Taux de sinistralité"
         
         .Range("B124:F124").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("B124:F124").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
        
    wbkOpen2.Worksheets("Feuil1").Range("A238:D238").Copy wbkThis.Worksheets("Feuil1").Range("B125")
    wbkOpen2.Worksheets("Feuil1").Range("G238").Copy wbkThis.Worksheets("Feuil1").Range("F125")
        
    With wbkThis.Worksheets("Feuil1")
         .Range("C124").FormulaR1C1 = 0
         .Range("D124").FormulaR1C1 = .Range("D123").Value / .Range("D125").Value
         .Range("E124").FormulaR1C1 = .Range("E123").Value / .Range("E125").Value
         .Range("F124").FormulaR1C1 = .Range("F123").Value / .Range("F125").Value
    
         .Range("B125:F125").Delete Shift:=xlUp
         .Range("C124:F124").NumberFormat = "0.00%"
        
         .Range("B124").FormulaR1C1 = "Taux de sinistralité"
         
         .Range("B126:F126").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("B126:F126").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
        
    wbkOpen2.Worksheets("Feuil1").Range("A239:D239").Copy wbkThis.Worksheets("Feuil1").Range("B127")
    wbkOpen2.Worksheets("Feuil1").Range("G239").Copy wbkThis.Worksheets("Feuil1").Range("F127")
        
    With wbkThis.Worksheets("Feuil1")
         .Range("C126").FormulaR1C1 = 0
         .Range("D126").FormulaR1C1 = 0
         .Range("E126").FormulaR1C1 = .Range("E125").Value / .Range("E127").Value
         .Range("F126").FormulaR1C1 = 0
    
         .Range("B127:F127").Delete Shift:=xlUp
         .Range("C126:F126").NumberFormat = "0.00%"
        
         .Range("B126").FormulaR1C1 = "Taux de sinistralité"
         
         .Range("B128:F128").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("B128:F128").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
        
    wbkOpen2.Worksheets("Feuil1").Range("A252:D252").Copy wbkThis.Worksheets("Feuil1").Range("B129")
    wbkOpen2.Worksheets("Feuil1").Range("G252").Copy wbkThis.Worksheets("Feuil1").Range("F129")
        
    With wbkThis.Worksheets("Feuil1")
         .Range("C128").FormulaR1C1 = 0
         .Range("D128").FormulaR1C1 = .Range("D127").Value / .Range("D129").Value
         .Range("E128").FormulaR1C1 = .Range("E127").Value / .Range("E129").Value
         .Range("F128").FormulaR1C1 = 0
    
         .Range("B129:F129").Delete Shift:=xlUp
         .Range("C128:F128").NumberFormat = "0.00%"
        
         .Range("B128").FormulaR1C1 = "Taux de sinistralité"
         
         .Range("B130:F130").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("B130:F130").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
        
    wbkOpen2.Worksheets("Feuil1").Range("A255:D255").Copy wbkThis.Worksheets("Feuil1").Range("B131")
    wbkOpen2.Worksheets("Feuil1").Range("G255").Copy wbkThis.Worksheets("Feuil1").Range("F131")
        
    With wbkThis.Worksheets("Feuil1")
         .Range("C130").FormulaR1C1 = 0
         .Range("D130").FormulaR1C1 = .Range("D129").Value / .Range("D131").Value
         .Range("E130").FormulaR1C1 = .Range("E129").Value / .Range("E131").Value
         .Range("F130").FormulaR1C1 = .Range("F129").Value / .Range("F131").Value
    
         .Range("B131:F131").Delete Shift:=xlUp
         .Range("C130:F130").NumberFormat = "0.00%"
        
         .Range("B130").FormulaR1C1 = "Taux de sinistralité"
         
         .Range("B132:F132").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("B132:F132").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
        
    wbkOpen2.Worksheets("Feuil1").Range("A259:D259").Copy wbkThis.Worksheets("Feuil1").Range("B133")
    wbkOpen2.Worksheets("Feuil1").Range("G259").Copy wbkThis.Worksheets("Feuil1").Range("F133")
        
    With wbkThis.Worksheets("Feuil1")
         .Range("C132").FormulaR1C1 = 0
         .Range("D132").FormulaR1C1 = .Range("D131").Value / .Range("D133").Value
         .Range("E132").FormulaR1C1 = 0
         .Range("F132").FormulaR1C1 = 0
    
         .Range("B133:F133").Delete Shift:=xlUp
         .Range("C132:F132").NumberFormat = "0.00%"
        
         .Range("B132").FormulaR1C1 = "Taux de sinistralité"
         
         .Range("B134:F134").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("B134:F134").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
        
    wbkOpen2.Worksheets("Feuil1").Range("A273:D273").Copy wbkThis.Worksheets("Feuil1").Range("B135")
    wbkOpen2.Worksheets("Feuil1").Range("G273").Copy wbkThis.Worksheets("Feuil1").Range("F135")
        
    With wbkThis.Worksheets("Feuil1")
         .Range("C134").FormulaR1C1 = 0
         .Range("D134").FormulaR1C1 = .Range("D133").Value / .Range("D135").Value
         .Range("E134").FormulaR1C1 = 0
         .Range("F134").FormulaR1C1 = 0
    
         .Range("B135:F135").Delete Shift:=xlUp
         .Range("C134:F134").NumberFormat = "0.00%"
        
         .Range("B134").FormulaR1C1 = "Taux de sinistralité"
         
         .Range("B136:F136").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("B136:F136").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
        
    wbkOpen2.Worksheets("Feuil1").Range("A274:D274").Copy wbkThis.Worksheets("Feuil1").Range("B137")
    wbkOpen2.Worksheets("Feuil1").Range("G274").Copy wbkThis.Worksheets("Feuil1").Range("F137")
        
    With wbkThis.Worksheets("Feuil1")
         .Range("C136").FormulaR1C1 = 0
         .Range("D136").FormulaR1C1 = .Range("D135").Value / .Range("D137").Value
         .Range("E136").FormulaR1C1 = .Range("E135").Value / .Range("E137").Value
         .Range("F136").FormulaR1C1 = .Range("F135").Value / .Range("F137").Value
    
         .Range("B137:F137").Delete Shift:=xlUp
         .Range("C136:F136").NumberFormat = "0.00%"
        
         .Range("B136").FormulaR1C1 = "Taux de sinistralité"
         
         .Range("B138:F138").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("B138:F138").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
        
    wbkOpen2.Worksheets("Feuil1").Range("A280:D280").Copy wbkThis.Worksheets("Feuil1").Range("B139")
    wbkOpen2.Worksheets("Feuil1").Range("G280").Copy wbkThis.Worksheets("Feuil1").Range("F139")
        
    With wbkThis.Worksheets("Feuil1")
         .Range("C138").FormulaR1C1 = 0
         .Range("D138").FormulaR1C1 = 0
         .Range("E138").FormulaR1C1 = .Range("E137").Value / .Range("E139").Value
         .Range("F138").FormulaR1C1 = 0
    
         .Range("B139:F139").Delete Shift:=xlUp
         .Range("C138:F138").NumberFormat = "0.00%"
        
         .Range("B138").FormulaR1C1 = "Taux de sinistralité"
         
         .Range("B140:F140").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("B140:F140").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
        
    wbkOpen2.Worksheets("Feuil1").Range("A282:D282").Copy wbkThis.Worksheets("Feuil1").Range("B141")
    wbkOpen2.Worksheets("Feuil1").Range("G282").Copy wbkThis.Worksheets("Feuil1").Range("F141")
        
    With wbkThis.Worksheets("Feuil1")
         .Range("C140").FormulaR1C1 = 0
         .Range("D140").FormulaR1C1 = .Range("D139").Value / .Range("D141").Value
         .Range("E140").FormulaR1C1 = 0
         .Range("F140").FormulaR1C1 = 0
    
         .Range("B141:F141").Delete Shift:=xlUp
         .Range("C140:F140").NumberFormat = "0.00%"
        
         .Range("B140").FormulaR1C1 = "Taux de sinistralité"
         
         .Range("B142:F142").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("B142:F142").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
        
    wbkOpen2.Worksheets("Feuil1").Range("A284:D284").Copy wbkThis.Worksheets("Feuil1").Range("B143")
    wbkOpen2.Worksheets("Feuil1").Range("G284").Copy wbkThis.Worksheets("Feuil1").Range("F143")
        
    With wbkThis.Worksheets("Feuil1")
         .Range("C142").FormulaR1C1 = 0
         .Range("D142").FormulaR1C1 = 0
         .Range("E142").FormulaR1C1 = 0
         .Range("F142").FormulaR1C1 = .Range("F141").Value / .Range("F143").Value
    
         .Range("B143:F143").Delete Shift:=xlUp
         .Range("C142:F142").NumberFormat = "0.00%"
        
         .Range("B142").FormulaR1C1 = "Taux de sinistralité"
         
         .Range("B144:F144").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("B144:F144").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
        
    wbkOpen2.Worksheets("Feuil1").Range("A287:D287").Copy wbkThis.Worksheets("Feuil1").Range("B145")
    wbkOpen2.Worksheets("Feuil1").Range("G287").Copy wbkThis.Worksheets("Feuil1").Range("F145")
        
    With wbkThis.Worksheets("Feuil1")
         .Range("C144").FormulaR1C1 = 0
         .Range("D144").FormulaR1C1 = .Range("D143").Value / .Range("D145").Value
         .Range("E144").FormulaR1C1 = 0
         .Range("F144").FormulaR1C1 = 0
    
         .Range("B145:F145").Delete Shift:=xlUp
         .Range("C144:F144").NumberFormat = "0.00%"
        
         .Range("B144").FormulaR1C1 = "Taux de sinistralité"
    End With
        
    With wbkThis.Worksheets("Feuil1").Range("B116:F116,B118:F118,B120:F120,B122:F122,B124:F124,B126:F126,B128:F128,B130:F130,B132:F132,B134:F134,B136:F136,B138:F138,B140:F140,B142:F142,B144:F144")
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
