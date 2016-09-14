Sub hyperlinks_GI_douteux()

    Dim wbkPrin As Workbook
    Dim wbkOpenGI As Workbook
    Dim shtPrin As Worksheet
    Dim shtGI As Worksheet
    Dim mR As Variant
    Dim slctRng As String
    Dim RowN As Long
    Dim nR As Long
    
    Set wbkPrin = ThisWorkbook
    
    ' open workbook "GI_douteux_copie"
    ' the address is variable
    Set wbkOpenGI = Workbooks.Open("P:\BDDs\après ETL\copie\GI_douteux_copie.xlsm")
    
    ' define worksheets
    Set shtPrin = wbkPrin.Sheets("Table_Principale")
    Set shtGI = wbkOpenGI.Sheets("GI")
    
    ' get the last row's number in worksheet "shtPrin"
    RowN = shtPrin.Cells(Rows.Count, 13).End(xlUp).Row
    
    ' remove all hyperlinks of column 56
    shtPrin.Columns(56).Hyperlinks.Delete

    ' go through all N concours in Table_Principale
    For nR = 2 To RowN
        ' locate the row of target N concours
        ' the column number is variable
        mR = Application.Match(shtPrin.Cells(nR, 13).Value, shtGI.Columns(6), 0)
        
            ' for every cell that is not empty,
            ' search through all N concours in GI_douteux_copie
            If IsError(mR) Then
                ' write nothing in the cell
                ' the column number is variable
                shtPrin.Cells(nR, 56).Value = ""
            Else
                ' the cells that should be chose
                slctRng = "A" & mR & ":AD" & mR
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGI
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 56), _
                Address:=wbkOpenGI.Path & "\GI_douteux_copie.xlsm", _
                SubAddress:="GI!" & slctRng, _
                TextToDisplay:="cliquez ici"
            End If
    Next nR

    wbkOpenGI.Close False
    
End Sub

