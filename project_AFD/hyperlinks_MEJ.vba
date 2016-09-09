Sub hyperlinks_MEJ()

    Dim wbkPrin As Workbook
    Dim wbkOpenMEJ As Workbook
    Dim shtPrin As Worksheet
    Dim shtMEJ As Worksheet
    Dim mR As Variant
    Dim bridge As String ' value of two lookup_value
    Dim slctRng As String
    Dim RowN As Long
    Dim nR As Long
    
    Set wbkPrin = ThisWorkbook
    
    ' open workbook "MEJ_copie"
    ' the address is variable
    Set wbkOpenMEJ = Workbooks.Open("P:\BDDs\après ETL\copie\MEJ_copie.xlsx")
    
    ' define worksheets
    Set shtPrin = wbkPrin.Sheets("Table_Principale")
    Set shtMEJ = wbkOpenMEJ.Sheets("MEJ")
    
    ' get the last row's number in worksheet "shtPrin"
    RowN = shtPrin.Cells(Rows.Count, 13).End(xlUp).Row
    
    ' remove all hyperlinks of column 59
    shtPrin.Columns(59).Hyperlinks.Delete

    ' go through all N concours in Table_Principale
    For nR = 2 To RowN
        ' create bridge which contains two columns
        bridge = shtPrin.Cells(nR, 13).Value & "_" & shtPrin.Cells(nR, 21).Value
        Debug.Print (bridge)
        
        ' locate the row of target N concours
        mR = Application.Match(bridge, shtMEJ.Columns(7), 0)
        
            ' for every cell that is not empty,
            ' search through all N concours in MEJ_copie
            If IsError(mR) Then
                ' write nothing in the cell
                ' the column number is variable
                shtPrin.Cells(nR, 59).Value = ""
            Else
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & mR
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtMEJ
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 59), _
                Address:=wbkOpenMEJ.Path & "\MEJ_copie.xlsx", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            End If
    Next nR

    wbkOpenMEJ.Close False

End Sub
