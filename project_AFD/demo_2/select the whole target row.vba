Sub select_the_whole_row_of_target_nconcours()

    Dim wbkT1 As Workbook
    Dim wbkOpenT2 As Workbook
    Dim nR As Long
    
    Set wbkT1 = ThisWorkbook
    
    ' open workbook "test2"
    Set wbkOpenT2 = Workbooks.Open("P:\BDDs\après ETL\copie\test2.xlsx")
    
    ' locate the row of target N concours
    nR = Application.Match(wbkT1.Sheets("t1_d1").Cells(2, 1).Value, wbkOpenT2.Sheets("t2_d1").Columns(1), 0)
    Debug.Print (nR)
    
    Windows(wbkOpenT2.Name).Activate
    ActiveWorkbook.Sheets("t2_d1").Activate
    ActiveSheet.Cells(nR, 1).EntireRow.Select

End Sub
