Sub display_result_in_a_new_wkb()
'
' return multiple columns from a VLOOKUP in a new workbook
'   
    Dim wbkOpenT2 As Workbook
    Dim wbkRslt As Workbook
    
    ' open workbook "test2"
    Set wbkOpenT2 = Workbooks.Open("P:\BDDs\après ETL\copie\test2.xlsx")
    
    ' create a new workbook for reading result
    Set wbkRslt = Workbooks.Add
    
    ' copy the fields' names from "test2"
    wbkOpenT2.Sheets("t2_d1").Range("A1:E1").Copy
    
    ' paste to new workbook
    wbkRslt.Sheets(1).Range("A1:E1").PasteSpecial
        
    ' select the cells (cells equal to the number of columns that you wish to
    ' fetch) where you wish to populate the VLOOKUP results
    wbkRslt.Sheets(1).Range("A2:E2").Select
    
    ' being to do VLOOKUP
    Selection.FormulaArray = _
        "=VLOOKUP([test1.xlsm]t1_d1!R2C1,[test2.xlsx]t2_d1!R2C1:R6C5,{1,2,3,4,5},0)"
End Sub
