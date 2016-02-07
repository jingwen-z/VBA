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

    Set MaxClose = AllClose.Find(overallmax, Range("E2"), xlValues, xlWhole)
    Debug.Print MaxClose.Row
    
    maxclosedate = Cells(MaxClose.Row, "A").Text
    Debug.Print maxclosedate
 
    
    MsgBox prompt:="The overall max ""Close"" value is " & overallmax & " which is on " & maxclosedate & ".", _
            Buttons:=vbOKOnly
            
End Sub
