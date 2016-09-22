Sub setting_format_Table_Principale()

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

    ' setting height of row
    Rows("1:1").RowHeight = 36.75
    
    ' setting width of columns
    Columns("A:A").ColumnWidth = 12.86
    Columns("B:B").ColumnWidth = 13
    Columns("C:C").ColumnWidth = 16.29
    Columns("D:D").ColumnWidth = 18.86
    Columns("E:E").ColumnWidth = 16.57
    Columns("F:F").ColumnWidth = 14.43
    Columns("G:G").ColumnWidth = 10.29
    Columns("H:H").ColumnWidth = 20.57
    Columns("I:I").ColumnWidth = 11.57
    Columns("K:K").ColumnWidth = 20.29
    Columns("M:M").ColumnWidth = 10.29
    Columns("N:N").ColumnWidth = 6
    Columns("O:O").EntireColumn.AutoFit
    Columns("P:P").ColumnWidth = 15.29
    Columns("Q:Q").ColumnWidth = 15.57
    Columns("R:R").ColumnWidth = 16.43
    Columns("S:S").EntireColumn.AutoFit
    Columns("T:T").EntireColumn.AutoFit
    Columns("U:U").EntireColumn.AutoFit
    Columns("V:V").ColumnWidth = 15.86
    Columns("W:W").ColumnWidth = 24.29
    Columns("X:X").ColumnWidth = 16
    Columns("Y:Y").ColumnWidth = 19
    Columns("Z:Z").ColumnWidth = 17.57
    Columns("AA:AA").ColumnWidth = 6.86
    Columns("AB:AB").ColumnWidth = 22.86
    Columns("AC:AC").ColumnWidth = 21.71
    Columns("AD:AD").ColumnWidth = 23
    Columns("AE:AE").ColumnWidth = 21.57
    Columns("AF:AF").ColumnWidth = 22.71
    Columns("AG:AG").ColumnWidth = 21.86
    Columns("AH:AH").ColumnWidth = 22.57
    Columns("AI:AI").ColumnWidth = 21.71
    Columns("AL:AL").ColumnWidth = 19
    Columns("AM:AM").ColumnWidth = 12.86
    Columns("AN:AN").ColumnWidth = 16.57
    Columns("AO:AO").ColumnWidth = 20.29
    Columns("AP:AP").ColumnWidth = 20.71
    Columns("AQ:AQ").ColumnWidth = 17.14
    Columns("AR:AR").ColumnWidth = 19.14
    Columns("AS:AS").ColumnWidth = 13.29
    Columns("AT:AT").ColumnWidth = 24.29
    Columns("AU:AU").ColumnWidth = 22.14
    Columns("AV:AV").ColumnWidth = 18.14
    Columns("AW:AW").ColumnWidth = 31.14
    Columns("AX:AX").ColumnWidth = 54.71
    Columns("AY:AY").ColumnWidth = 15.29
    Columns("AZ:AZ").ColumnWidth = 11.29
    Columns("BA:BA").ColumnWidth = 18.14
    Columns("BB:BB").ColumnWidth = 12.43

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
    
    Range("AM1:AQ1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    Range("AR1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    Range("AS1:AW1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With

    Range("AX1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range("AY1:BB1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13082801
        .TintAndShade = 0
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
    Columns("C:C").NumberFormat = "m/d/yyyy"
    Columns("AB:AI").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Columns("AJ:AJ").NumberFormat = "0.0%"
    Columns("AK:AK").NumberFormat = "0.00"
    Columns("AL:AL").NumberFormat = "0.0%"
    Columns("AR:AR").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

    ' autofilter
    Selection.AutoFilter

    ' freeze panes
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
End Sub
