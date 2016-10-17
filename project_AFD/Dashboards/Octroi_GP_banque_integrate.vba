Sub Octroi_GP()

    Dim wbkThis As Workbook
    Dim wbkOpen As Workbook
    Dim Rng As Range
    
    Set wbkThis = ThisWorkbook
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\GPP_31-12-15_TdB.xlsm")
    
    wbkOpen.Worksheets("Feuil1").Range("A6:H9").Copy wbkThis.Worksheets("Feuil1").Range("B49")
    
    wbkOpen.Worksheets("Feuil1").Range("A26:H26").Copy
    wbkThis.Worksheets("Feuil1").Range("B51:I51").Insert Shift:=xlDown
    wbkOpen.Worksheets("Feuil1").Range("B15:B16").Copy wbkThis.Worksheets("Feuil1").Range("J49")
        
    wbkOpen.Worksheets("Feuil1").Range("A27:H27").Copy
    wbkThis.Worksheets("Feuil1").Range("B53:I53").Insert Shift:=xlDown
    wbkOpen.Worksheets("Feuil1").Range("B17").Copy wbkThis.Worksheets("Feuil1").Range("J52")
    
    wbkOpen.Worksheets("Feuil1").Range("A28:H28").Copy
    wbkThis.Worksheets("Feuil1").Range("B55:I55").Insert Shift:=xlDown
    wbkOpen.Worksheets("Feuil1").Range("B18").Copy wbkThis.Worksheets("Feuil1").Range("J54")
        
    wbkThis.Worksheets("Feuil1").Range("B49").FormulaR1C1 = "Octroi GP (en M€)"
    wbkThis.Worksheets("Feuil1").Range("I49").FormulaR1C1 = "Total"
    wbkThis.Worksheets("Feuil1").Range("J49").FormulaR1C1 = "Encours"
    wbkThis.Worksheets("Feuil1").Range("B51").FormulaR1C1 = "Taux d'utilisation"
    wbkThis.Worksheets("Feuil1").Range("B53").FormulaR1C1 = "Taux d'utilisation"
    wbkThis.Worksheets("Feuil1").Range("B55").FormulaR1C1 = "Taux d'utilisation"
    
    wbkOpen.Close False
    
    For Each Rng In Range("C50:H55")
    
        If IsError(Rng.Value) Then
            Rng.Value = 0#
        End If
    
    Next Rng
    
End Sub
