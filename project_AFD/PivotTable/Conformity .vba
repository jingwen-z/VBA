Sub Conformité()

    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\Conformité_TdB.xlsx")

    wbkOpen.Worksheets("TdB___Conformité").Range("A1:D7").Copy wbkThis.Worksheets("Feuil1").Range("B86")
    
    With wbkThis.Worksheets("Feuil1")
        .Range("B86").FormulaR1C1 = "Conformité"
        
        With .Range("B86:E86").Interior
             .Pattern = xlSolid
             .PatternColorIndex = xlAutomatic
             .ThemeColor = xlThemeColorAccent1
             .TintAndShade = 0.799981688894314
             .PatternTintAndShade = 0
        End With

        With .Range("B86:E86")
             .Borders(xlEdgeTop).LineStyle = xlNone
             .Borders(xlEdgeLeft).LineStyle = xlNone
             .Borders(xlEdgeRight).LineStyle = xlNone
             .Borders(xlDiagonalUp).LineStyle = xlNone
             .Borders(xlDiagonalDown).LineStyle = xlNone
             .Borders(xlInsideVertical).LineStyle = xlNone
             .Borders(xlInsideHorizontal).LineStyle = xlNone
             .Font.Bold = True
        End With

        With .Range("B86:E86").Borders(xlEdgeBottom)
             .LineStyle = xlContinuous
             .ThemeColor = 5
             .TintAndShade = 0.399945066682943
             .Weight = xlThin
        End With
        
    End With

    wbkOpen.Close False

End Sub
