Sub hyperlinks_GPP()

    Dim wbkPrin As Workbook
    Dim wbkOpenGPP As Workbook
    Dim shtPrin As Worksheet
    Dim shtGPP As Worksheet
    Dim mR As Variant
    Dim slctRng As String
    Dim RowN As Long
    Dim nR As Long
    
    Set wbkPrin = ThisWorkbook
    
    ' open workbook "GPP_31-12-15_copie"
    ' the address is variable
    Set wbkOpenGPP = Workbooks.Open("P:\BDDs\après ETL\copie\GPP_31-12-15_copie.xlsm")
    
    ' define worksheets
    Set shtPrin = wbkPrin.Sheets("Table_Principale")
    Set shtGPP = wbkOpenGPP.Sheets("GPP")
    
    ' get the last row's number in worksheet "shtPrin"
    RowN = shtPrin.Cells(Rows.Count, 13).End(xlUp).Row
    
    ' remove all hyperlinks of column 57
    shtPrin.Columns(57).Hyperlinks.Delete

    ' go through all N concours in Table_Principale
    For nR = 2 To RowN
        ' locate the row of target N concours
        mR = Application.Match(shtPrin.Cells(nR, 13).Value, shtGPP.Columns(3), 0)
        
            ' for every cell that is not empty,
            ' search through all N concours in GPP_31-12-15_copie
            If IsError(mR) Then
                ' write nothing in the cell
                ' the column number is variable
                shtPrin.Cells(nR, 57).Value = ""
            Else
                ' the cells that should be chose
                slctRng = "A" & mR & ":FS" & mR
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 57), _
                Address:=wbkOpenGPP.Path & "\GPP_31-12-15_copie.xlsm", _
                SubAddress:="GPP!" & slctRng, _
                TextToDisplay:="cliquez ici"
            End If
    Next nR

    wbkOpenGPP.Close False

End Sub

