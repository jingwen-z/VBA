Sub GAR_TdB()
    
    Dim wbkOpen As Workbook
    Dim wbkOpen2 As Workbook
    Dim wbkOpen3 As Workbook
    Dim cl As Long
    Dim colN As Long
    Dim n As Long

    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")
    Set wbkOpen2 = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\MEJ-TCD.xlsm")
    Set wbkOpen3 = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\1- ARIZ suiviReporting Global-TCD.xlsm")
    
    With wbkOpen.Worksheets("TCD_global")
        .Range("A4:K5").Copy ThisWorkbook.Worksheets("Feuil1").Range("B4")
        .Range("A14:K14").Copy ThisWorkbook.Worksheets("Feuil1").Range("B6")
        .Range("A23:K23").Copy ThisWorkbook.Worksheets("Feuil1").Range("B7")
    End With
    
    colN = wbkOpen.Worksheets("TCD_global").Range("K5").Column - wbkOpen.Worksheets("TCD_global").Range("A4").Column
    
    With wbkOpen2.Worksheets("TCD_global")
        .Range("A5:K5").Copy ThisWorkbook.Worksheets("Feuil1").Range("B9")
    End With
    
    With ThisWorkbook.Worksheets("Feuil1").Range("B8")
        
        For cl = 1 To colN
            .Offset(0, cl).FormulaR1C1 = .Offset(1, cl).Value / .Offset(-3, cl).Value
        Next cl
            
        For cl = 1 To colN
            .Offset(0, cl).NumberFormat = "0.00%"
            .Offset(1, 0).Delete Shift:=xlToLeft
        Next cl
        
        .Offset(1, 0).Delete Shift:=xlToLeft

    End With

    wbkOpen2.Worksheets("TCD_global").Range("A14:K14").Copy ThisWorkbook.Worksheets("Feuil1").Range("B10")
    
    With ThisWorkbook.Worksheets("Feuil1").Range("B9")
                
        For cl = 1 To colN
            .Offset(0, cl).FormulaR1C1 = .Offset(1, cl).Value / .Offset(-3, cl).Value
        Next cl
            
        For cl = 1 To colN
            .Offset(0, cl).NumberFormat = "0.00%"
            .Offset(1, 0).Delete Shift:=xlToLeft
        Next cl
        
        .Offset(1, 0).Delete Shift:=xlToLeft
    
    End With

    wbkOpen2.Worksheets("TCD_global").Range("A23:H23").Copy ThisWorkbook.Worksheets("Feuil1").Range("B11")
    With ThisWorkbook.Worksheets("Feuil1").Range("B10")
                        
        For cl = 1 To colN
            .Offset(0, cl).FormulaR1C1 = .Offset(1, cl).Value / .Offset(-4, cl).Value
        Next cl
            
        For cl = 1 To colN
            .Offset(0, cl).NumberFormat = "0.00%"
            .Offset(1, 0).Delete Shift:=xlToLeft
        Next cl
        
        .Offset(1, 0).Delete Shift:=xlToLeft
    
    End With
    
    With wbkOpen.Worksheets("TCD_global")
        .Range("A6:K6").Copy ThisWorkbook.Worksheets("Feuil1").Range("B11")
        .Range("A15:K15").Copy ThisWorkbook.Worksheets("Feuil1").Range("B12")
        .Range("A24:K24").Copy ThisWorkbook.Worksheets("Feuil1").Range("B13")
    End With
    
    With wbkOpen2.Worksheets("TCD_global")
        .Range("A15:K15").Copy ThisWorkbook.Worksheets("Feuil1").Range("B15")
    End With
    
    With wbkOpen3.Worksheets("TCD")
        .Range("B77:G77").Copy ThisWorkbook.Worksheets("Feuil1").Range("D16")
        .Range("I77").Copy ThisWorkbook.Worksheets("Feuil1").Range("L16")
    End With
    

    With ThisWorkbook.Worksheets("Feuil1").Range("B14")
                
        For cl = 1 To colN
            If .Offset(2, cl).Value = 0 Then
                .Offset(0, cl).Value = 0
            Else
                .Offset(0, cl).FormulaR1C1 = .Offset(1, cl).Value / .Offset(2, cl).Value
            End If
        Next cl
            
        For cl = 1 To colN
            .Offset(0, cl).NumberFormat = "0.00%"
            .Offset(1, 0).Delete Shift:=xlToLeft
            .Offset(2, 0).Delete Shift:=xlToLeft
        Next cl
        
        .Offset(1, 0).Delete Shift:=xlToLeft
        .Offset(2, 0).Delete Shift:=xlToLeft
    
    End With

    wbkOpen2.Worksheets("TCD_global").Range("A24:K24").Copy ThisWorkbook.Worksheets("Feuil1").Range("B16")
    
    With wbkOpen3.Worksheets("TCD")
        .Range("B77:G77").Copy ThisWorkbook.Worksheets("Feuil1").Range("D17")
        .Range("I77").Copy ThisWorkbook.Worksheets("Feuil1").Range("L17")
    End With
    
    With ThisWorkbook.Worksheets("Feuil1").Range("B15")
                
        For cl = 1 To colN
            If .Offset(2, cl).Value = 0 Then
                .Offset(0, cl).Value = 0
            Else
                .Offset(0, cl).FormulaR1C1 = .Offset(1, cl).Value / .Offset(2, cl).Value
            End If
        Next cl
            
        For cl = 1 To colN
            .Offset(0, cl).NumberFormat = "0.00%"
            .Offset(1, 0).Delete Shift:=xlToLeft
            .Offset(2, 0).Delete Shift:=xlToLeft
        Next cl
        
        .Offset(1, 0).Delete Shift:=xlToLeft
        .Offset(2, 0).Delete Shift:=xlToLeft
    
    End With
    
    With wbkOpen.Worksheets("TCD_global")
    .Range("A7:K7").Copy ThisWorkbook.Worksheets("Feuil1").Range("B16")
    .Range("A16:K16").Copy ThisWorkbook.Worksheets("Feuil1").Range("B17")
    .Range("A25:K25").Copy ThisWorkbook.Worksheets("Feuil1").Range("B18")
    End With
    
    wbkOpen2.Worksheets("TCD_global").Range("A25:K25").Copy ThisWorkbook.Worksheets("Feuil1").Range("B19")

    wbkOpen.Close False
    wbkOpen2.Close False
    wbkOpen3.Close False
    
    With ThisWorkbook.Worksheets("Feuil1").Range("B4")
        .FormulaR1C1 = "Tableau de bord global"
        .Offset(1, 0).FormulaR1C1 = "GI octroyé en nombre"
        .Offset(2, 0).FormulaR1C1 = "GI montant octroyé (en M€)"
        .Offset(3, 0).FormulaR1C1 = "GI encours restant"
        .Offset(4, 0).FormulaR1C1 = "GI taux de sinistralité en nombre"
        .Offset(5, 0).FormulaR1C1 = "GI taux de sinistralité demandé par la banque"
        .Offset(6, 0).FormulaR1C1 = "GI taux de sinistralité (avec montant d'intemnisation max)"
        .Offset(7, 0).FormulaR1C1 = "GP octroyé en nombre"
        .Offset(8, 0).FormulaR1C1 = "GP montant octroyé (en M€)"
        .Offset(9, 0).FormulaR1C1 = "GP encours restant"
        .Offset(10, 0).FormulaR1C1 = "GP taux de sinistralité demandé par la banque"
        .Offset(11, 0).FormulaR1C1 = "GP taux de sinistralité (avec montant d'intemnisation max)"
        .Offset(12, 0).FormulaR1C1 = "Total nombre octroyé"
        .Offset(13, 0).FormulaR1C1 = "Total montant octroyé (en M€)"
        .Offset(14, 0).FormulaR1C1 = "Total encours (en M€)"
        .Offset(15, 0).FormulaR1C1 = "Total montant d'intemnisation max (en M€)"
        .Offset(0, 10).FormulaR1C1 = "Total"
    End With
    
    With ThisWorkbook.Worksheets("Feuil1")
        For cl = 0 To colN
            For n = 0 To 3
                .Range("B16").Offset(n, cl).Font.Bold = False
                With .Range("B16").Offset(n, cl).Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            Next n
        Next cl

        With Range("B17:L17")
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        End With

        With Range("B19:I19")
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        End With

        With Range("B15:L15")
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 5
                .TintAndShade = 0.399945066682943
                .Weight = xlThin
            End With
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone

        End With

        With Range("B10:L10")
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 5
                .TintAndShade = 0.399945066682943
                .Weight = xlThin
            End With
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        End With
    End With

End Sub
