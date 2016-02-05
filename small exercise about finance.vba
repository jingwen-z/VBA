'analyse the historical prices of GOOG (just for VBA exercise)
'I download the data of the period between 27/03/2014 and 29/01/2016 from Yahoo Finance

'First, I will find the max "Close" among all data
Sub max_close_overall()
    
    Dim allclose As Range
    
    Set allclose = Feuil1.Range("E2:E466")  'HOW TO IMPROVE THIS???
    
    overallmax = Application.WorksheetFunction.Max(allclose)
    
    MsgBox prompt:="The overall max ""Close"" value is " & overallmax & ".", _
            Buttons:=vbOKOnly
            
End Sub
