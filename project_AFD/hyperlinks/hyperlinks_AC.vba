Sub hyperlinks_AC()

    Dim wbkPrin As Workbook
    Dim wbkOpenAC As Workbook
    Dim shtPrin As Worksheet
    Dim shtAC As Worksheet
    Dim mR As Variant
    Dim slctRng As String
    Dim RowN As Long
    Dim nR As Long
    
    Set wbkPrin = ThisWorkbook
    
    ' open workbook "Arrêté_Comptable_copie"
    ' the address is variable
    Set wbkOpenAC = Workbooks.Open("P:\BDDs\après ETL\copie\Arrêté_Comptable_copie.xlsm")
    
    ' define worksheets
    Set shtPrin = wbkPrin.Sheets("Table_Principale")
    Set shtAC = wbkOpenAC.Sheets("Arrêté_Comptable")
    
    ' get the last row's number in worksheet "shtPrin"
    RowN = shtPrin.Cells(Rows.Count, 13).End(xlUp).Row
    
    ' remove all hyperlinks of column 55
    shtPrin.Columns(55).Hyperlinks.Delete

    ' go through all N concours in Table_Principale
    For nR = 2 To RowN
        ' locate the row of target N concours
        mR = Application.Match(shtPrin.Cells(nR, 13).Value, shtAC.Columns(13), 0)
        
            ' for every cell that is not empty,
            ' search through all N concours in Arrêté_Comptable_copie
            If IsError(mR) Then
                ' write nothing in the cell
                ' the column number is variable
                shtPrin.Cells(nR, 55).Value = ""
            Else
                ' the cells that should be chose
                slctRng = "A" & mR & ":BI" & mR
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtAC
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 55), _
                Address:=wbkOpenAC.Path & "\Arrêté_Comptable_copie.xlsm", _
                SubAddress:="Arrêté_Comptable!" & slctRng, _
                TextToDisplay:="cliquez ici"
            End If
    Next nR
    
    wbkOpenAC.Close False

End Sub
