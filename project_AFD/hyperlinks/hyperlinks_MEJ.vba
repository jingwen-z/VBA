Sub hyperlinks_MEJ()

    Dim wbkPrin As Workbook
    Dim wbkOpenMEJ As Workbook
    Dim shtPrin As Worksheet
    Dim shtMEJ As Worksheet
    Dim mR As Variant
    Dim slctRng As String
    Dim RowN As Long
    Dim nR As Long
    
    Set wbkPrin = ThisWorkbook
    
    ' open workbook "MEJ_30-06-16_copie"
    ' the address is variable
    Set wbkOpenMEJ = Workbooks.Open("P:\BDDs\après ETL\copie\MEJ_30-06-16_copie.xlsm")
    
    ' define worksheets
    Set shtPrin = wbkPrin.Sheets("Table_Principale")
    Set shtMEJ = wbkOpenMEJ.Sheets("MEJ")
    
    ' get the last row's number in worksheet "shtPrin"
    RowN = shtPrin.Cells(Rows.Count, 13).End(xlUp).Row
    
    ' remove all hyperlinks of column 60
    shtPrin.Columns(60).Hyperlinks.Delete

    ' go through all N concours in Table_Principale
    For nR = 2 To RowN
        ' locate the row of target N° concours
        mR = Application.Match(shtPrin.Cells(nR, 13).Value, shtMEJ.Columns(6), 0)
        
            ' for every cell that is not empty,
            ' search through all N concours in MEJ_30-06-16_copie
            If IsError(mR) Then
                ' write nothing in the cell
                ' the column number is variable
                shtPrin.Cells(nR, 60).Value = ""
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CBF120201" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 2)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CCI132201" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 7)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CCM122001" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 15)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CCM122101" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 12)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CCM123201" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 16)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CCM128501" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 14)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CCM128601" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 24)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CCM132301" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 32)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CCM132303" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 3)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CCM133401" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 4)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CGA116002" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 1)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CGA120101" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 4)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CMG134001" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 1)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CMG134902" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 10)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CMG136302" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 1)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CMG143502" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 9)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CNA101701" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 1)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CSN130401" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 7)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CSN139901" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 1)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CSN140801" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 4)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CTD114202" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 1)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CTD116401" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 3)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CUG103101" Then
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & (mR + 2)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            
            Else
                ' the cells that should be chose
                slctRng = "A" & mR & ":CA" & mR
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtMEJ
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 60), _
                Address:=wbkOpenMEJ.Path & "\MEJ_30-06-16_copie.xlsm", _
                SubAddress:="MEJ!" & slctRng, _
                TextToDisplay:="cliquez ici"
            End If
    Next nR

    wbkOpenMEJ.Close False

End Sub
