Sub hyperlinks_the_whole_row_of_target_nconcours()

    Dim wbkT1 As Workbook
    Dim wbkOpenT2 As Workbook
    Dim shtT1D2 As Worksheet
    Dim shtT2D2 As Worksheet
    Dim mR As Variant
    Dim slctRng As String
    Dim RowN As Long
    Dim nR As Long
    
    Set wbkT1 = ThisWorkbook
    
    ' open workbook "test2"
    Set wbkOpenT2 = Workbooks.Open("P:\BDDs\après ETL\copie\test2.xlsx")
    
    ' define worksheets
    Set shtT1D2 = wbkT1.Sheets("t1_d2")
    Set shtT2D2 = wbkOpenT2.Sheets("t2_d2")
    
    shtT2D2.Activate
    shtT2D2.Columns(1).Select
    
    
    ' get the last row's number in worksheet "shtT1D2"
    RowN = shtT1D2.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' remove all hyperlinks
    shtT1D2.Hyperlinks.Delete
    

    ' go through all N concours in test1
    For nR = 2 To RowN
        ' locate the row of target N concours
        mR = Application.Match(shtT1D2.Cells(nR, 1).Value, shtT2D2.Columns(1), 0)
        
            ' for every cell that is not empty,
            ' search through all N concours in test2
            If IsError(mR) Then
                ' write nothing in the cell
                shtT1D2.Cells(nR, 6).Value = ""
            Else
                ' the cells that should be chose
                slctRng = "A" & mR & ":E" & mR
                ' active wbkT1
                Windows(wbkT1.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtT2D2
                shtT1D2.Hyperlinks.Add Anchor:=Cells(nR, 6), _
                Address:="P:\BDDs\après ETL\copie\test2.xlsx", _
                SubAddress:="t2_d2!" & slctRng, _
                TextToDisplay:="cliquez ici"
            End If
    Next nR

End Sub
