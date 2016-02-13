'analyse the historical prices of GOOG (just for VBA exercise)
'I download the data of the period between 27/03/2014 and 29/01/2016 from Yahoo Finance

'First, I will find the max "Close" among all data
Sub max_close_overall()
    
    Dim RowN As Long
    Dim LastRow As Range
    Dim AllClose As Range
    Dim MaxClose As Range
    
    If Cells(Rows.Count, "E").Value = "" Then
        Set LastRow = Cells(Rows.Count, "E").End(xlUp)
    Else
        Set LastRow = Cells(Rows.Count, "E")
    End If
    
    RowN = LastRow.Row
    Debug.Print RowN
    
    Set AllClose = Feuil1.Range(Cells(2, "E"), Cells(RowN, "E"))
    
    overallmax = Application.WorksheetFunction.Max(AllClose)
    '776.599976
    
    Set MaxClose = AllClose.Find(overallmax, Range("E2"), xlValues, xlWhole)
    Debug.Print MaxClose.Row
    '23
    
    maxclosedate = Cells(MaxClose.Row, "A").Text
    Debug.Print maxclosedate
    '2015/12/29

    Range(Cells(MaxClose.Row, "A"), Cells(MaxClose.Row, "G")).Interior.Color = RGB(255, 255, 0)

    MsgBox prompt:="The overall max ""Close"" value is " & overallmax & " which is on " & maxclosedate & ".", _
            Buttons:=vbOKOnly
            
    'Cells.Interior.ColorIndex = 0
    'clean the interior color immediately after closing the msgbox
            
End Sub

'Then, I will try to find out the max "Close" value for each year
Sub max_close_14()

    Dim RowN14 As Long

    'clean the interior color
    Cells.Interior.ColorIndex = 0

    Range("A1").AutoFilter Field:=1, Criteria2:=Array(0, "12/31/2014"), Operator:=xlFilterValues
    
    RowN14 = Range("A1").CurrentRegion.Rows.Count
    'Debug.Print RowN14
    '466
    
    maxClose14 = Cells(RowN14, "E").Value
    'Debug.Print maxClose14
    '596.082692
    
    maxClose14date = Cells(RowN14, "A").Text
    'Debug.Print maxClose14date
    '2014/9/19
    
    Range(Cells(RowN14, "A"), Cells(RowN14, "G")).Interior.Color = RGB(255, 255, 0)
    
    MsgBox prompt:="In 2014, the max value of ""Close"" is " & maxClose14 & " which is on " & maxClose14date & ".", _
            Buttons:=vbOKOnly

    Selection.AutoFilter
    
End Sub

