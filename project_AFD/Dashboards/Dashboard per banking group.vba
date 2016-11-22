Public wbkOpen As Workbook
Public rw As Long
Public rowN As Long
Public cl As Long
Public colN As Long
Public n As Long

Sub GrpBq_Octroi_GI_et_GP()
    
    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")
    
    wbkOpen.Worksheets("TCD").Range("AA6:AK8").Copy ThisWorkbook.Worksheets("Feuil1").Range("B4")
    rowN = wbkOpen.Worksheets("TCD").Range("AK8").Row - wbkOpen.Worksheets("TCD").Range("AA6").Row

    For rw = 1 To rowN - 1
        wbkOpen.Worksheets("TCD").Range("AA16:AK16").Offset(rw - 1, 0).Copy
        ThisWorkbook.Worksheets("Feuil1").Range("B4:L4").Offset(2 * rw, 0).Insert Shift:=xlDown
        
        With ThisWorkbook.Worksheets("Feuil1").Range("B4:L4").Offset(2 * rw, 0)
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
    Next rw

    wbkOpen.Worksheets("TCD").Range("AA16:AK16").Offset(rowN - 1, 0).Copy _
    ThisWorkbook.Worksheets("Feuil1").Range("B4:L4").Offset(2 * rowN, 0)

    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B4").FormulaR1C1 = ""
        .Range("B5").FormulaR1C1 = "Octroi(en M€) GI"
        .Range("B6").FormulaR1C1 = "GI Encours(en M€)"
        .Range("B7").FormulaR1C1 = "Octroi(en M€) GP"
        .Range("B8").FormulaR1C1 = "GP Encours(en M€)"
        .Range("K4").FormulaR1C1 = "2016 act."
        .Range("L4").FormulaR1C1 = "Total"
    End With
    
    wbkOpen.Close False
    
End Sub

Sub GrpBq_Octroi_GI()
    
    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")
    
    wbkOpen.Worksheets("TCD").Range("AA25:AK35").Copy ThisWorkbook.Worksheets("Feuil1").Range("B13")
    
    rowN = wbkOpen.Worksheets("TCD").Range("AK35").Row - wbkOpen.Worksheets("TCD").Range("AA25").Row
    colN = wbkOpen.Worksheets("TCD").Range("AK35").Column - wbkOpen.Worksheets("TCD").Range("AA25").Column

    For rw = 1 To rowN - 1
        wbkOpen.Worksheets("TCD").Range("AA80:AK80").Offset(rw - 1, 0).Copy
        ThisWorkbook.Worksheets("Feuil1").Range("B13:L13").Offset(2 * rw, 0).Insert Shift:=xlDown
        
        With ThisWorkbook.Worksheets("Feuil1").Range("B13:L13").Offset(2 * rw, 0)
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
    Next rw

    wbkOpen.Worksheets("TCD").Range("AA80:AK80").Offset(rowN - 1, 0).Copy _
    ThisWorkbook.Worksheets("Feuil1").Range("B13:L13").Offset(2 * rowN, 0)

    For rw = 1 To rowN
        With ThisWorkbook.Worksheets("Feuil1")
            .Range("C13").Offset(2 * rw, 0).FormulaR1C1 = .Range("C13").Offset(2 * rw, 0).Value / 1000000
            .Range("D13").Offset(2 * rw, 0).FormulaR1C1 = .Range("D13").Offset(2 * rw, 0).Value / 1000000
            .Range("E13").Offset(2 * rw, 0).FormulaR1C1 = .Range("E13").Offset(2 * rw, 0).Value / 1000000
            .Range("F13").Offset(2 * rw, 0).FormulaR1C1 = .Range("F13").Offset(2 * rw, 0).Value / 1000000
            .Range("G13").Offset(2 * rw, 0).FormulaR1C1 = .Range("G13").Offset(2 * rw, 0).Value / 1000000
            .Range("H13").Offset(2 * rw, 0).FormulaR1C1 = .Range("H13").Offset(2 * rw, 0).Value / 1000000
            .Range("I13").Offset(2 * rw, 0).FormulaR1C1 = .Range("I13").Offset(2 * rw, 0).Value / 1000000
            .Range("J13").Offset(2 * rw, 0).FormulaR1C1 = .Range("J13").Offset(2 * rw, 0).Value / 1000000
            .Range("K13").Offset(2 * rw, 0).FormulaR1C1 = .Range("K13").Offset(2 * rw, 0).Value / 1000000
            .Range("L13").Offset(2 * rw, 0).FormulaR1C1 = .Range("L13").Offset(2 * rw, 0).Value / 1000000
   
            .Range("C13:L13").Offset(2 * rw, 0).NumberFormat = "0.00"
            .Range("B13").Offset(2 * rw, 0).FormulaR1C1 = "Moyenne des GI octroyées"
        End With
    Next rw

    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B13").FormulaR1C1 = "Octroi GI (en M€)"
        .Range("K13").FormulaR1C1 = "2016 act."
        .Range("B13").Offset(0, colN).FormulaR1C1 = "Total"
    End With

    wbkOpen.Close False
    
End Sub

Sub GrpBq_Octroi_GI_en_nombre()

    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")
    
    wbkOpen.Worksheets("TCD").Range("AA133:AK144").Copy ThisWorkbook.Worksheets("Feuil1").Range("B47")
    
    rowN = wbkOpen.Worksheets("TCD").Range("AK144").Row - wbkOpen.Worksheets("TCD").Range("AA133").Row
    colN = wbkOpen.Worksheets("TCD").Range("AK144").Column - wbkOpen.Worksheets("TCD").Range("AA133").Column
    
    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B47").FormulaR1C1 = "Octroi GI (en nombre)"
        .Range("K47").FormulaR1C1 = "2016 act."
        .Range("B47").Offset(0, colN).FormulaR1C1 = "Total"
        .Range("B47").Offset(rowN, 0).FormulaR1C1 = "Total"
    End With
    
    wbkOpen.Close False

End Sub

Sub GrpBq_Octroi_GP()
    
    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")
    
    wbkOpen.Worksheets("TCD").Range("AA160:AJ169").Copy ThisWorkbook.Worksheets("Feuil1").Range("B72")
    
    rowN = wbkOpen.Worksheets("TCD").Range("AJ169").Row - wbkOpen.Worksheets("TCD").Range("AA160").Row
    colN = wbkOpen.Worksheets("TCD").Range("AJ169").Column - wbkOpen.Worksheets("TCD").Range("AA160").Column

    For rw = 1 To rowN - 1
        wbkOpen.Worksheets("TCD").Range("AA213:AJ213").Offset(rw - 1, 0).Copy
        ThisWorkbook.Worksheets("Feuil1").Range("B72:K72").Offset(2 * rw, 0).Insert Shift:=xlDown
        
        With ThisWorkbook.Worksheets("Feuil1").Range("B72:K72").Offset(2 * rw, 0)
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
    Next rw

    wbkOpen.Worksheets("TCD").Range("AA213:AJ213").Offset(rowN - 1, 0).Copy _
    ThisWorkbook.Worksheets("Feuil1").Range("B72:K72").Offset(2 * rowN, 0)

    For rw = 1 To rowN
        With ThisWorkbook.Worksheets("Feuil1")
            .Range("C72").Offset(2 * rw, 0).FormulaR1C1 = .Range("C72").Offset(2 * rw, 0).Value / 1000000
            .Range("D72").Offset(2 * rw, 0).FormulaR1C1 = .Range("D72").Offset(2 * rw, 0).Value / 1000000
            .Range("E72").Offset(2 * rw, 0).FormulaR1C1 = .Range("E72").Offset(2 * rw, 0).Value / 1000000
            .Range("F72").Offset(2 * rw, 0).FormulaR1C1 = .Range("F72").Offset(2 * rw, 0).Value / 1000000
            .Range("G72").Offset(2 * rw, 0).FormulaR1C1 = .Range("G72").Offset(2 * rw, 0).Value / 1000000
            .Range("H72").Offset(2 * rw, 0).FormulaR1C1 = .Range("H72").Offset(2 * rw, 0).Value / 1000000
            .Range("I72").Offset(2 * rw, 0).FormulaR1C1 = .Range("I72").Offset(2 * rw, 0).Value / 1000000
            .Range("J72").Offset(2 * rw, 0).FormulaR1C1 = .Range("J72").Offset(2 * rw, 0).Value / 1000000
            .Range("K72").Offset(2 * rw, 0).FormulaR1C1 = .Range("K72").Offset(2 * rw, 0).Value / 1000000
   
            .Range("C72:K72").Offset(2 * rw, 0).NumberFormat = "0.00"
            .Range("B72").Offset(2 * rw, 0).FormulaR1C1 = "Moyenne des GI octroyées"
        End With
    Next rw

    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B72").FormulaR1C1 = "Octroi GI (en M€)"
        .Range("J72").FormulaR1C1 = "2016 act."
        .Range("B72").Offset(0, colN).FormulaR1C1 = "Total"
    End With

    wbkOpen.Close False
    
End Sub

Sub GrpBq_Octroi_GP_en_nombre()

    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")
    
    wbkOpen.Worksheets("TCD").Range("AA264:AJ274").Copy ThisWorkbook.Worksheets("Feuil1").Range("B103")
    
    rowN = wbkOpen.Worksheets("TCD").Range("AJ274").Row - wbkOpen.Worksheets("TCD").Range("AA264").Row
    colN = wbkOpen.Worksheets("TCD").Range("AJ274").Column - wbkOpen.Worksheets("TCD").Range("AA264").Column
    
    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B103").FormulaR1C1 = "Octroi GP (en nombre)"
        .Range("J103").FormulaR1C1 = "2016 act."
        .Range("B103").Offset(0, colN).FormulaR1C1 = "Total"
        .Range("B103").Offset(rowN, 0).FormulaR1C1 = "Total"
    End With
    
    wbkOpen.Close False

End Sub
