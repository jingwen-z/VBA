Sub setting_format_AC()

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
    Columns("A:A").ColumnWidth = 12.86
    Columns("B:B").ColumnWidth = 13
    Columns("C:C").ColumnWidth = 16.29
    Columns("D:D").ColumnWidth = 18.86
    Columns("E:E").ColumnWidth = 16.57
    Columns("F:F").ColumnWidth = 14.43
    Columns("G:G").ColumnWidth = 10.29
    Columns("H:H").EntireColumn.AutoFit
    Columns("I:I").ColumnWidth = 11.57
    Columns("K:K").ColumnWidth = 20.29
    Columns("M:M").ColumnWidth = 10.29
    Columns("N:N").ColumnWidth = 6
    Columns("O:O").EntireColumn.AutoFit
    Columns("P:P").EntireColumn.AutoFit
    Columns("Q:Q").EntireColumn.AutoFit
    Columns("R:R").EntireColumn.AutoFit
    Columns("S:S").EntireColumn.AutoFit
    Columns("T:T").EntireColumn.AutoFit
    Columns("U:U").EntireColumn.AutoFit
    Columns("V:V").EntireColumn.AutoFit
    Columns("W:W").EntireColumn.AutoFit
    Columns("X:X").EntireColumn.AutoFit
    Columns("Y:Y").ColumnWidth = 19
    Columns("Z:Z").EntireColumn.AutoFit
    Columns("AA:AA").ColumnWidth = 6.86
    Columns("AB:AB").ColumnWidth = 21.57
    Columns("AC:AC").ColumnWidth = 21.57
    Columns("AD:AD").ColumnWidth = 22.14
    Columns("AE:AE").ColumnWidth = 21.57
    Columns("AF:AF").ColumnWidth = 21.86
    Columns("AG:AG").ColumnWidth = 21.86
    Columns("AH:AH").ColumnWidth = 21.71
    Columns("AI:AI").ColumnWidth = 21.71
    Columns("AL:AL").ColumnWidth = 19
    Columns("AM:AM").ColumnWidth = 14.57
    Columns("AN:AN").ColumnWidth = 17.57
    Columns("AO:AO").ColumnWidth = 22.14
    Columns("AP:AP").ColumnWidth = 14.57
    Columns("AQ:AQ").ColumnWidth = 14.57
    Columns("AR:AR").ColumnWidth = 17.57
    Columns("AS:AS").ColumnWidth = 22.14
    Columns("AT:AT").ColumnWidth = 14.57
    Columns("AU:AU").ColumnWidth = 14.71
    Columns("AV:AV").ColumnWidth = 18.43
    Columns("AW:AW").ColumnWidth = 18.71
    Columns("AX:AX").ColumnWidth = 15.29
    Columns("AY:AY").ColumnWidth = 20.43
    Columns("AZ:AZ").ColumnWidth = 43.71
    Columns("BA:BA").ColumnWidth = 45.14
    Columns("BB:BB").ColumnWidth = 48
    Columns("BC:BC").ColumnWidth = 24.29
    Columns("BD:BD").ColumnWidth = 13.14
    Columns("BE:BE").ColumnWidth = 18.86
    Columns("BF:BF").ColumnWidth = 19.43
    Columns("BG:BG").ColumnWidth = 19.29
    Columns("BH:BH").ColumnWidth = 14.43
    Columns("BI:BI").ColumnWidth = 35.71
    Columns("BJ:BJ").ColumnWidth = 56

    ' setting color of headers
    Range("A1:E1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With

    Range("F1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13311
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range("G1:K1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6750207
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    Range("L1:Z1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    Range("AA1:AL1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    Range("AM1:AP1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    Range("AQ1:AT1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With

    Range("AU1:AZ1").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range("BA1:BB1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13082801
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range("BC1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range("BD1:BH1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    Range("BI1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
        
    Range("BJ1").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    ' setting color of fond
    Range("AU1:BB1").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
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
    Columns("C:C").NumberFormat = "m/d/yyyy"
    Columns("AB:AI").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Columns("AJ:AJ").NumberFormat = "0%"
    Columns("AK:AK").NumberFormat = "0.00"
    Columns("AM:AX").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Columns("BC:BC").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Columns("BI:BI").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

    ' autofilter
    Selection.AutoFilter
    
    ' freeze panes
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    ' hiding several fields
    Columns("D:H").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:K").Select
    Selection.EntireColumn.Hidden = True
    Columns("P:R").Select
    Selection.EntireColumn.Hidden = True
    
End Sub
