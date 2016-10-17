Sub Octroi_GI()
    
    Dim wbkThis As Workbook
    Dim wbkOpen As Workbook
    
    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\Table_Principale_30-06-16_TdB.xlsm")
    
    wbkOpen.Worksheets("Feuil1").Range("A24:K31").Copy wbkThis.Worksheets("Feuil1").Range("B10")
    wbkOpen.Worksheets("Feuil1").Range("B38").Copy wbkThis.Worksheets("Feuil1").Range("M10")
    wbkOpen.Worksheets("Feuil1").Range("B40:B44").Copy wbkThis.Worksheets("Feuil1").Range("M11")
    wbkOpen.Worksheets("Feuil1").Range("B46:B47").Copy wbkThis.Worksheets("Feuil1").Range("M16")
    
    wbkOpen.Worksheets("Feuil1").Range("A56:K56").Copy
    wbkThis.Worksheets("Feuil1").Range("B12:L12").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("B71").Copy
    wbkThis.Worksheets("Feuil1").Range("M12").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("A57:K57").Copy
    wbkThis.Worksheets("Feuil1").Range("B14:L14").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("B72").Copy
    wbkThis.Worksheets("Feuil1").Range("M14").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("A58:K58").Copy
    wbkThis.Worksheets("Feuil1").Range("B16:L16").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("B73").Copy
    wbkThis.Worksheets("Feuil1").Range("M16").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("A59:K59").Copy
    wbkThis.Worksheets("Feuil1").Range("B18:L18").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("B74").Copy
    wbkThis.Worksheets("Feuil1").Range("M18").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("A60:K60").Copy
    wbkThis.Worksheets("Feuil1").Range("B20:L20").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("B75").Copy
    wbkThis.Worksheets("Feuil1").Range("M20").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("A61:K61").Copy
    wbkThis.Worksheets("Feuil1").Range("B22:L22").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("B77").Copy
    wbkThis.Worksheets("Feuil1").Range("M22").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("A62:K62").Copy
    wbkThis.Worksheets("Feuil1").Range("B24:L24").Insert Shift:=xlDown
    
    wbkOpen.Worksheets("Feuil1").Range("B78").Copy
    wbkThis.Worksheets("Feuil1").Range("M24").Insert Shift:=xlDown
    
    With wbkThis.Worksheets("Feuil1")
    
        .Range("C12").FormulaR1C1 = .Range("C12").Value / 1000000
        .Range("D12").FormulaR1C1 = .Range("D12").Value / 1000000
        .Range("E12").FormulaR1C1 = .Range("E12").Value / 1000000
        .Range("F12").FormulaR1C1 = .Range("F12").Value / 1000000
        .Range("G12").FormulaR1C1 = .Range("G12").Value / 1000000
        .Range("H12").FormulaR1C1 = .Range("H12").Value / 1000000
        .Range("I12").FormulaR1C1 = .Range("I12").Value / 1000000
        .Range("J12").FormulaR1C1 = .Range("J12").Value / 1000000
        .Range("K12").FormulaR1C1 = .Range("K12").Value / 1000000
        .Range("L12").FormulaR1C1 = .Range("L12").Value / 1000000
        .Range("M12").FormulaR1C1 = .Range("M12").Value / 1000000
    
        
        .Range("C14").FormulaR1C1 = .Range("C14").Value / 1000000
        .Range("D14").FormulaR1C1 = .Range("D14").Value / 1000000
        .Range("E14").FormulaR1C1 = .Range("E14").Value / 1000000
        .Range("F14").FormulaR1C1 = .Range("F14").Value / 1000000
        .Range("G14").FormulaR1C1 = .Range("G14").Value / 1000000
        .Range("H14").FormulaR1C1 = .Range("H14").Value / 1000000
        .Range("I14").FormulaR1C1 = .Range("I14").Value / 1000000
        .Range("J14").FormulaR1C1 = .Range("J14").Value / 1000000
        .Range("K14").FormulaR1C1 = .Range("K14").Value / 1000000
        .Range("L14").FormulaR1C1 = .Range("L14").Value / 1000000
        .Range("M14").FormulaR1C1 = .Range("M14").Value / 1000000
    
        .Range("C16").FormulaR1C1 = .Range("C16").Value / 1000000
        .Range("D16").FormulaR1C1 = .Range("D16").Value / 1000000
        .Range("E16").FormulaR1C1 = .Range("E16").Value / 1000000
        .Range("F16").FormulaR1C1 = .Range("F16").Value / 1000000
        .Range("G16").FormulaR1C1 = .Range("G16").Value / 1000000
        .Range("H16").FormulaR1C1 = .Range("H16").Value / 1000000
        .Range("I16").FormulaR1C1 = .Range("I16").Value / 1000000
        .Range("J16").FormulaR1C1 = .Range("J16").Value / 1000000
        .Range("K16").FormulaR1C1 = .Range("K16").Value / 1000000
        .Range("L16").FormulaR1C1 = .Range("L16").Value / 1000000
        .Range("M16").FormulaR1C1 = .Range("M16").Value / 1000000
        
        .Range("C18").FormulaR1C1 = .Range("C18").Value / 1000000
        .Range("D18").FormulaR1C1 = .Range("D18").Value / 1000000
        .Range("E18").FormulaR1C1 = .Range("E18").Value / 1000000
        .Range("F18").FormulaR1C1 = .Range("F18").Value / 1000000
        .Range("G18").FormulaR1C1 = .Range("G18").Value / 1000000
        .Range("H18").FormulaR1C1 = .Range("H18").Value / 1000000
        .Range("I18").FormulaR1C1 = .Range("I18").Value / 1000000
        .Range("J18").FormulaR1C1 = .Range("J18").Value / 1000000
        .Range("K18").FormulaR1C1 = .Range("K18").Value / 1000000
        .Range("L18").FormulaR1C1 = .Range("L18").Value / 1000000
        .Range("M18").FormulaR1C1 = .Range("M18").Value / 1000000
    
        .Range("C20").FormulaR1C1 = .Range("C20").Value / 1000000
        .Range("D20").FormulaR1C1 = .Range("D20").Value / 1000000
        .Range("E20").FormulaR1C1 = .Range("E20").Value / 1000000
        .Range("F20").FormulaR1C1 = .Range("F20").Value / 1000000
        .Range("G20").FormulaR1C1 = .Range("G20").Value / 1000000
        .Range("H20").FormulaR1C1 = .Range("H20").Value / 1000000
        .Range("I20").FormulaR1C1 = .Range("I20").Value / 1000000
        .Range("J20").FormulaR1C1 = .Range("J20").Value / 1000000
        .Range("K20").FormulaR1C1 = .Range("K20").Value / 1000000
        .Range("L20").FormulaR1C1 = .Range("L20").Value / 1000000
        .Range("M20").FormulaR1C1 = .Range("M20").Value / 1000000
    
        .Range("C22").FormulaR1C1 = .Range("C22").Value / 1000000
        .Range("D22").FormulaR1C1 = .Range("D22").Value / 1000000
        .Range("E22").FormulaR1C1 = .Range("E22").Value / 1000000
        .Range("F22").FormulaR1C1 = .Range("F22").Value / 1000000
        .Range("G22").FormulaR1C1 = .Range("G22").Value / 1000000
        .Range("H22").FormulaR1C1 = .Range("H22").Value / 1000000
        .Range("I22").FormulaR1C1 = .Range("I22").Value / 1000000
        .Range("J22").FormulaR1C1 = .Range("J22").Value / 1000000
        .Range("K22").FormulaR1C1 = .Range("K22").Value / 1000000
        .Range("L22").FormulaR1C1 = .Range("L22").Value / 1000000
        .Range("M22").FormulaR1C1 = .Range("M22").Value / 1000000
    
        .Range("C24").FormulaR1C1 = .Range("C24").Value / 1000000
        .Range("D24").FormulaR1C1 = .Range("D24").Value / 1000000
        .Range("E24").FormulaR1C1 = .Range("E24").Value / 1000000
        .Range("F24").FormulaR1C1 = .Range("F24").Value / 1000000
        .Range("G24").FormulaR1C1 = .Range("G24").Value / 1000000
        .Range("H24").FormulaR1C1 = .Range("H24").Value / 1000000
        .Range("I24").FormulaR1C1 = .Range("I24").Value / 1000000
        .Range("J24").FormulaR1C1 = .Range("J24").Value / 1000000
        .Range("K24").FormulaR1C1 = .Range("K24").Value / 1000000
        .Range("L24").FormulaR1C1 = .Range("L24").Value / 1000000
        .Range("M24").FormulaR1C1 = .Range("M24").Value / 1000000
   
        .Range("C12:M12").NumberFormat = "0.00"
        .Range("C14:M14").NumberFormat = "0.00"
        .Range("C16:M16").NumberFormat = "0.00"
        .Range("C18:M18").NumberFormat = "0.00"
        .Range("C20:M20").NumberFormat = "0.00"
        .Range("C22:M22").NumberFormat = "0.00"
        .Range("C24:M24").NumberFormat = "0.00"
    
        .Range("B10").FormulaR1C1 = "Octroi GI (en M€)"
        .Range("K10").FormulaR1C1 = "2016 act."
        .Range("L10").FormulaR1C1 = "Total"
        .Range("M10").FormulaR1C1 = "Encours act."
        .Range("B12").FormulaR1C1 = "Moyenne des GI octroyées"
        .Range("B14").FormulaR1C1 = "Moyenne des GI octroyées"
        .Range("B16").FormulaR1C1 = "Moyenne des GI octroyées"
        .Range("B18").FormulaR1C1 = "Moyenne des GI octroyées"
        .Range("B20").FormulaR1C1 = "Moyenne des GI octroyées"
        .Range("B22").FormulaR1C1 = "Moyenne des GI octroyées"
        .Range("B24").FormulaR1C1 = "Moyenne des GI octroyées"

    End With
    
    wbkOpen.Close False
    
End Sub
