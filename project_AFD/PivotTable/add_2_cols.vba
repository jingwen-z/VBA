Sub add_2_cols()

    Dim wbkMEJ As Workbook
    Dim wbkPrin As Workbook
    Dim shtMEJ As Worksheet

    Set wbkMEJ = ThisWorkbook
    Set wbkPrin = Workbooks.Open(wbkMEJ.Path & "\Table_Principale_30-06-16_TdB.xlsm")

    Set shtMEJ = wbkMEJ.Sheets("MEJ")

    With shtMEJ
         .Columns("W:W").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
         .Columns("W:W").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("W1").FormulaR1C1 = "Nature prêt"
         .Range("X1").FormulaR1C1 = "Secteur détaillé"
         .Range("W2").FormulaR1C1 = _
            "=VLOOKUP(RC[-17],'[Table_Principale_30-06-16_TdB.xlsm]Table_Principale'!C13:C45,33,0)"
         .Range("W2").AutoFill Destination:=Range("W2:W297")
         .Range("X2").FormulaR1C1 = _
            "=VLOOKUP(RC[-18],'[Table_Principale_30-06-16_TdB.xlsm]Table_Principale'!C13:C46,34,0)"
         .Range("X2").AutoFill Destination:=Range("X2:X297")
    End With
    
End Sub
