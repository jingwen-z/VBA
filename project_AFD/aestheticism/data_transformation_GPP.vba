Sub data_transformation_GPP()

    Workbooks.Open Filename:= _
        "S:\EBC\GAR\3 - MIDDLE OFFICE\2015\Suivi encours GPP\Au 31-12-2015\1- ARIZ suiviReporting Global 31 12 2015.xlsm" _
        , UpdateLinks:=0
    Range("A1:FS3").Copy
    
    Windows("GPP_copie.xlsm").Activate
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Rows("4:4").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp

    'wbkOpenAC.Close False

End Sub
