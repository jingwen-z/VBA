Sub select_the_target_row_special_case()

    Dim wbkT1 As Workbook
    Dim wbkOpenT2 As Workbook
    Dim nR As Long
    Dim bridge As String ' value of two lookup_value
    
    Set wbkT1 = ThisWorkbook
    bridge = wbkT1.Sheets("t1_d1").Cells(2, 1).Value & wbkT1.Sheets("t1_d1").Cells(2, 2).Value
    
    ' open workbook "test2"
    Set wbkOpenT2 = Workbooks.Open("P:\BDDs\après ETL\copie\test2.xlsx")
    
    ' locate the row of target N concours
    nR = Application.Match(bridge, wbkOpenT2.Sheets("t2_d1").Columns(1), 0)
    Debug.Print (nR)
    Debug.Print (bridge)
    
    
    Windows(wbkOpenT2.Name).Activate
    ActiveWorkbook.Sheets("t2_d1").Activate
    ActiveSheet.Cells(nR, 1).EntireRow.Select

End Sub
