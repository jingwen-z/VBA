Sub Octroi_GP()

    Dim wbkThis As Workbook
    Dim wbkOpen As Workbook
    Dim Rng As Range
    
    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\GPP_31-12-15_TCD.xlsm")
    
    wbkOpen.Worksheets("Feuil1").Range("A6:H9").Copy wbkThis.Worksheets("Feuil1").Range("B43")
    
    wbkOpen.Worksheets("Feuil1").Range("A36:H36").Copy
    wbkThis.Worksheets("Feuil1").Range("B45:I45").Insert Shift:=xlDown
    wbkOpen.Worksheets("Feuil1").Range("B20:B21").Copy wbkThis.Worksheets("Feuil1").Range("J43")
        
    wbkOpen.Worksheets("Feuil1").Range("A37:H37").Copy
    wbkThis.Worksheets("Feuil1").Range("B47:I47").Insert Shift:=xlDown
    wbkOpen.Worksheets("Feuil1").Range("B22").Copy wbkThis.Worksheets("Feuil1").Range("J46")
    
    wbkOpen.Worksheets("Feuil1").Range("A38:H38").Copy
    wbkThis.Worksheets("Feuil1").Range("B49:I49").Insert Shift:=xlDown
    wbkOpen.Worksheets("Feuil1").Range("B23").Copy wbkThis.Worksheets("Feuil1").Range("J48")
        
    With wbkThis.Worksheets("Feuil1")
         .Range("B43").FormulaR1C1 = "Octroi GP (en M€)"
         .Range("I43").FormulaR1C1 = "Total"
         .Range("J43").FormulaR1C1 = "Encours act."
         .Range("B45").FormulaR1C1 = "Taux d'utilisation"
         .Range("B47").FormulaR1C1 = "Taux d'utilisation"
         .Range("B49").FormulaR1C1 = "Taux d'utilisation"
         
         .Range("C45:I45").NumberFormat = "0.00%"
         .Range("C47:I47").NumberFormat = "0.00%"
         .Range("C49:I49").NumberFormat = "0.00%"

    End With
    
    With wbkThis.Worksheets("Feuil1").Range("B45:J45,B47:J47,B49:J49")
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
    
    For Each Rng In Range("C50:H55")
    
        If IsError(Rng.Value) Then
            Rng.Value = 0#
        End If
    
    Next Rng
    
End Sub
