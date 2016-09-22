Sub setting_format_MEJ()

    Cells.Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
        
    Range("A1", Selection.End(xlToRight)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Font.Bold = True
    End With
    
    ' rename some fields
    Range("AA1").FormulaR1C1 = "Evènement générateur-Date décheance du terme"
    Range("AB1").FormulaR1C1 = _
        "Evènement générateur-Date de l'info du fait générateur par la banque à l'AFD"
    Range("AC1").FormulaR1C1 = "Evènement générateur-Délai respecté"
    Range("AO1").FormulaR1C1 = _
        "Détermination Indemnisation-Perte provisoire calculée par la banque en devise"
    Range("AP1").FormulaR1C1 = _
        "Détermination Indemnisation-Perte provisoire accordée par l'AFD en devise"
    Range("AQ1").FormulaR1C1 = _
        "Détermination Indemnisation-Perte provisoire accordée par l'AFD en €"
    Range("AR1").FormulaR1C1 = _
        "Détermination Indemnisation-Différence sur l'assiette de garantie de la MEJ"
    Range("AS1").FormulaR1C1 = "Détermination Indemnisation-Evaluation des sûretés"
    Range("AT1").FormulaR1C1 = "Détermination Indemnisation-Commentaire"

    ' setting width of columns
    Rows("1:1").RowHeight = 36.75
    Columns("A:A").ColumnWidth = 19
    Columns("B:B").ColumnWidth = 17.29
    Columns("C:C").ColumnWidth = 16
    Columns("D:D").ColumnWidth = 15
    Columns("E:E").ColumnWidth = 11.29
    Columns("F:F").ColumnWidth = 11
    Columns("G:G").ColumnWidth = 51
    Columns("H:H").EntireColumn.AutoFit
    Columns("I:I").ColumnWidth = 16.57
    Columns("J:J").EntireColumn.AutoFit
    Columns("K:K").EntireColumn.AutoFit
    Columns("L:L").ColumnWidth = 42
    Columns("M:M").ColumnWidth = 16.57
    Columns("N:N").EntireColumn.AutoFit
    Columns("O:O").ColumnWidth = 7.57
    Columns("P:P").ColumnWidth = 24.14
    Columns("Q:Q").ColumnWidth = 21.86
    Columns("R:R").ColumnWidth = 23.57
    Columns("S:S").ColumnWidth = 21.71
    Columns("T:T").ColumnWidth = 10.71
    Columns("U:U").ColumnWidth = 10.71
    Columns("V:V").ColumnWidth = 16.29
    Columns("W:W").ColumnWidth = 18
    Columns("X:X").ColumnWidth = 22.71
    Columns("Y:Y").ColumnWidth = 30.57
    Columns("Z:Z").ColumnWidth = 13
    Columns("AA:AA").ColumnWidth = 23.29
    Columns("AB:AB").ColumnWidth = 32.71
    Columns("AC:AC").ColumnWidth = 20.43
    Columns("AD:AD").ColumnWidth = 18.14
    Columns("AE:AE").ColumnWidth = 28.86
    Columns("AF:AF").ColumnWidth = 21.71
    Columns("AG:AG").ColumnWidth = 17.43
    Columns("AH:AH").ColumnWidth = 21.14
    Columns("AI:AI").ColumnWidth = 14.43
    Columns("AJ:AJ").ColumnWidth = 30
    Columns("AK:AK").ColumnWidth = 15
    Columns("AL:AL").ColumnWidth = 27
    Columns("AM:AM").ColumnWidth = 15.71
    Columns("AN:AN").ColumnWidth = 27.29
    Columns("AO:AO").ColumnWidth = 38.86
    Columns("AP:AP").ColumnWidth = 37.57
    Columns("AQ:AQ").ColumnWidth = 34
    Columns("AR:AR").EntireColumn.AutoFit
    Columns("AS:AS").EntireColumn.AutoFit
    Columns("AT:AT").ColumnWidth = 43.86
    Columns("AU:AU").EntireColumn.AutoFit
    Columns("AV:AV").EntireColumn.AutoFit
    Columns("AW:AW").EntireColumn.AutoFit
    Columns("AX:AX").EntireColumn.AutoFit
    Columns("AY:AY").EntireColumn.AutoFit
    Columns("AZ:AZ").ColumnWidth = 29.71
    Columns("BA:BA").ColumnWidth = 17.71
    Columns("BB:BB").EntireColumn.AutoFit
    Columns("BC:BC").ColumnWidth = 13.14
    Columns("BD:BD").ColumnWidth = 29
    Columns("BE:BE").ColumnWidth = 13.43
    Columns("BF:BF").ColumnWidth = 16.57
    Columns("BG:BG").ColumnWidth = 19.29
    Columns("BH:BH").ColumnWidth = 15.57
    Columns("BI:BI").ColumnWidth = 19.29
    Columns("BJ:BJ").ColumnWidth = 21.71
    Columns("BK:BK").ColumnWidth = 12.43
    Columns("BL:BL").EntireColumn.AutoFit
    Columns("BM:BM").ColumnWidth = 26.86
    Columns("BN:BN").ColumnWidth = 16.57
    Columns("BO:BO").ColumnWidth = 44.43
    Columns("BP:BP").ColumnWidth = 15.86
    Columns("BQ:BQ").ColumnWidth = 12
    Columns("BR:BR").ColumnWidth = 15
    Columns("BS:BS").EntireColumn.AutoFit
    Columns("BT:BT").EntireColumn.AutoFit
    Columns("BU:BU").ColumnWidth = 20.86
    Columns("BV:BV").ColumnWidth = 17.29
    Columns("BW:BW").ColumnWidth = 15.31
    Columns("BX:BX").EntireColumn.AutoFit
    Columns("BY:BY").ColumnWidth = 50.71
    Columns("BZ:BZ").ColumnWidth = 17.86
    Columns("CA:CA").ColumnWidth = 62.57

    ' setting color of headers
    Range("A1:C1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    Range("D1:N1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10092543
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range("O1:V1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 14857357
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range("W1:Z1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    Range("AA1:AC1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    Range("AD1,AF1,AG1,AH1,AI1,AK1,AM1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 14408946
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range("AE1,AJ1,AL1,AN1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15395562
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
  
  Range("AO1:AT1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With

    Range("AU1:AY1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13082801
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
  
  Range("AZ1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15395562
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
  
  Range("BA1:BH1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    Range("BI1:BT1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With

    Range("BU1:CA1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With

    ' setting global borders
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    'setting border for headers
    Range("A1", Selection.End(xlToRight)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    ' setting format
    Columns("P:S").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Columns("T:T").NumberFormat = "0.0%"
    Columns("U:U").NumberFormat = "0.00"
    Columns("AO:AS").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Columns("AU:AW").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Columns("BB:BC").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Columns("BL:BN").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Columns("BU:BU").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    
    ' autofilter
    Selection.AutoFilter
    
    ' freeze panes
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
End Sub
