Sub data_transformation_GPP()

    Workbooks.Open Filename:= _
        "blablabla..." _
        , UpdateLinks:=0
    Range("A1:FS3").Copy
    ' "blablabla..." is the full path of workbook
    
    Windows("GPP_copie.xlsm").Activate
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Rows("4:4").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp

    'close the first opened workbook
    Windows("blabla.xxx").Close False
    ' "blabla.xxx" is the name of workbook

End Sub
