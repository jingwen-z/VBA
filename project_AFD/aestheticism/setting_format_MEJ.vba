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
    Columns("I:I").EntireColumn.AutoFit
    Columns("J:J").EntireColumn.AutoFit
    Columns("K:K").EntireColumn.AutoFit
    Columns("L:L").ColumnWidth = 42
    Columns("M:M").EntireColumn.AutoFit
    Columns("N:N").EntireColumn.AutoFit
    Columns("O:O").ColumnWidth = 7.57
    Columns("P:P").ColumnWidth = 21.86
    Columns("Q:Q").ColumnWidth = 21.86
    Columns("R:R").ColumnWidth = 21.71
    Columns("S:S").ColumnWidth = 21.71
    Columns("T:T").ColumnWidth = 10.71
    Columns("U:U").ColumnWidth = 10.71
    Columns("V:V").ColumnWidth = 16.29
    Columns("W:W").ColumnWidth = 18
    Columns("X:X").ColumnWidth = 20.71
    Columns("Y:Y").ColumnWidth = 30.57
    Columns("Z:Z").ColumnWidth = 13
    Columns("AA:AA").ColumnWidth = 17.14
    Columns("AB:AB").ColumnWidth = 29
    Columns("AC:AC").ColumnWidth = 9.14
    Columns("AD:AD").ColumnWidth = 17
    Columns("AE:AE").ColumnWidth = 28.86
    Columns("AF:AF").ColumnWidth = 10.14
    Columns("AG:AG").ColumnWidth = 14.14
    Columns("AH:AH").EntireColumn.AutoFit
    Columns("AI:AI").ColumnWidth = 12.29
    Columns("AJ:AJ").EntireColumn.AutoFit
    Columns("AK:AK").ColumnWidth = 15
    Columns("AL:AL").EntireColumn.AutoFit
    Columns("AM:AM").EntireColumn.AutoFit
    Columns("AN:AN").ColumnWidth = 27.29
    Columns("AO:AO").EntireColumn.AutoFit
    Columns("AP:AP").EntireColumn.AutoFit
    Columns("AQ:AQ").EntireColumn.AutoFit
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
    Columns("BC:BC").EntireColumn.AutoFit
    Columns("BD:BD").ColumnWidth = 29
    Columns("BE:BE").ColumnWidth = 13.43
    Columns("BF:BF").EntireColumn.AutoFit
    Columns("BG:BG").EntireColumn.AutoFit
    Columns("BH:BH").EntireColumn.AutoFit
    Columns("BI:BI").EntireColumn.AutoFit
    Columns("BJ:BJ").EntireColumn.AutoFit
    Columns("BK:BK").EntireColumn.AutoFit
    Columns("BL:BL").EntireColumn.AutoFit
    Columns("BM:BM").EntireColumn.AutoFit
    Columns("BN:BN").EntireColumn.AutoFit
    Columns("BO:BO").ColumnWidth = 44.43
    Columns("BP:BP").EntireColumn.AutoFit
    Columns("BQ:BQ").ColumnWidth = 12
    Columns("BR:BR").EntireColumn.AutoFit
    Columns("BS:BS").EntireColumn.AutoFit
    Columns("BT:BT").EntireColumn.AutoFit
    Columns("BU:BU").EntireColumn.AutoFit
    Columns("BV:BV").EntireColumn.AutoFit
    Columns("BW:BW").ColumnWidth = 15.31
Columns("BX:BX").EntireColumn.AutoFit
Columns("BY:BY").ColumnWidth = 50.71
Columns("BZ:BZ").EntireColumn.AutoFit
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

    Range("BR2:BX2").Select
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
    Columns("B:B").NumberFormat = "m/d/yyyy"
    Columns("P:S").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Columns("T:T").NumberFormat = "0%"
    Columns("W:W").NumberFormat = "m/d/yyyy"
    Columns("AO:AS").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Columns("AU:AW").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Columns("AX:AY").NumberFormat = "m/d/yyyy"
    Columns("BB:BC").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Columns("BL:BN").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Columns("BS:BT").NumberFormat = "m/d/yyyy"
    Columns("BU:BU").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

    Selection.AutoFilter
    
End Sub
