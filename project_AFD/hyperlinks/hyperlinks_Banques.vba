Sub hyperlinks_Banques()

    Dim wbkPrin As Workbook
    Dim wbkOpenBq As Workbook
    Dim shtPrin As Worksheet
    Dim shtBq As Worksheet
    Dim mR As Variant
    Dim slctRng As String
    Dim RowN As Long
    Dim nR As Long
    
    Set wbkPrin = ThisWorkbook
    
    ' open workbook "Banques_copie"
    ' the address is variable
    Set wbkOpenBq = Workbooks.Open("P:\BDDs\après ETL\copie\Banques_copie.xlsm")
    
    ' define worksheets
    Set shtPrin = wbkPrin.Sheets("Table_Principale")
    Set shtBq = wbkOpenBq.Sheets("Banques")
    
    ' get the last row's number in worksheet "shtPrin"
    RowN = shtPrin.Cells(Rows.Count, 19).End(xlUp).Row
    
    ' remove all hyperlinks of column 59
    shtPrin.Columns(59).Hyperlinks.Delete

    ' go through all N concours in Table_Principale
    For nR = 2 To RowN
        ' locate the row of target N concours
        mR = Application.Match(shtPrin.Cells(nR, 19).Value, shtBq.Columns(2), 0)
        
            ' for every cell that is not empty,
            ' search through all N concours in Banques_copie
            If IsError(mR) Then
                ' write nothing in the cell
                ' the column number is variable
                shtPrin.Cells(nR, 59).Value = ""
            Else
                ' the cells that should be chose
                slctRng = "A" & mR & ":V" & mR
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtBq
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 59), _
                Address:=wbkOpenBq.Path & "\Banques_copie.xlsm", _
                SubAddress:="Banques!" & slctRng, _
                TextToDisplay:="cliquez ici"
            End If
    Next nR

    wbkOpenBq.Close False

End Sub
