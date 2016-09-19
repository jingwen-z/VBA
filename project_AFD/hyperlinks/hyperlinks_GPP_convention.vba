Sub hyperlinks_GPP_convention()

    Dim wbkPrin As Workbook
    Dim wbkOpenGPP As Workbook
    Dim shtPrin As Worksheet
    Dim shtGPP As Worksheet
    Dim mR As Variant
    Dim shtName As String
    Dim RowN As Long
    Dim nR As Long
    
    Set wbkPrin = ThisWorkbook
    
    ' open workbook "conventions"
    ' the address is variable
    Set wbkOpenGPP = Workbooks.Open("P:\BDDs\copie des bases\Copie de 1- ARIZ suiviReporting Global 31 12 2015_conventions.xlsm")
    
    ' define worksheets
    Set shtPrin = wbkPrin.Sheets("Table_Principale")
    Set shtGPP = wbkOpenGPP.Sheets("BDD GPP")
    
    ' get the last row's number in worksheet "shtPrin"
    RowN = shtPrin.Cells(Rows.Count, 13).End(xlUp).Row
    
    ' remove all hyperlinks of column 58
    shtPrin.Columns(58).Hyperlinks.Delete

    ' go through all N concours in Table_Principale
    For nR = 2 To RowN
        ' locate the row of target N concours
        mR = Application.Match(shtPrin.Cells(nR, 13).Value, shtGPP.Columns(3), 0)
        
            ' for every cell that is not empty,
            ' search through all N concours in worksheet "BDD GPP"
            If IsError(mR) Then
                ' write nothing in the cell
                ' the column number is variable
                shtPrin.Cells(nR, 58).Value = ""
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CRW102302" Then
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 58), _
                Address:=wbkOpenGPP.Path & "\Copie de 1- ARIZ suiviReporting Global 31 12 2015_conventions.xlsm", _
                SubAddress:="'BANK OF KIGALI - CRW102302'!A1", _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CMU105201" Then
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 58), _
                Address:=wbkOpenGPP.Path & "\Copie de 1- ARIZ suiviReporting Global 31 12 2015_conventions.xlsm", _
                SubAddress:="'BANK ONE - CMU105201'!A1", _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CMG143501" Then
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 58), _
                Address:=wbkOpenGPP.Path & "\Copie de 1- ARIZ suiviReporting Global 31 12 2015_conventions.xlsm", _
                SubAddress:="'BFV-SG - CMG143501'!A1", _
                TextToDisplay:="cliquez ici"
            
            ElseIf shtPrin.Cells(nR, 13).Value = "CGA120101" Then
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 58), _
                Address:=wbkOpenGPP.Path & "\Copie de 1- ARIZ suiviReporting Global 31 12 2015_conventions.xlsm", _
                SubAddress:="'BGFI - CGA120101'!A1", _
                TextToDisplay:="cliquez ici"
                
            ElseIf shtPrin.Cells(nR, 13).Value = "CBF121301" Then
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 58), _
                Address:=wbkOpenGPP.Path & "\Copie de 1- ARIZ suiviReporting Global 31 12 2015_conventions.xlsm", _
                SubAddress:="'BICIAB - CBF121301'!A1", _
                TextToDisplay:="cliquez ici"
                
            ElseIf shtPrin.Cells(nR, 13).Value = "CCI130901" Then
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 58), _
                Address:=wbkOpenGPP.Path & "\Copie de 1- ARIZ suiviReporting Global 31 12 2015_conventions.xlsm", _
                SubAddress:="'BICICI - CCI130901'!A1", _
                TextToDisplay:="cliquez ici"
                
            ElseIf shtPrin.Cells(nR, 13).Value = "CKE105601" Then
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 58), _
                Address:=wbkOpenGPP.Path & "\Copie de 1- ARIZ suiviReporting Global 31 12 2015_conventions.xlsm", _
                SubAddress:="'COOPERATIVE BANK - CKE105601'!A1", _
                TextToDisplay:="cliquez ici"
                
            ElseIf shtPrin.Cells(nR, 13).Value = "CKE105701" Then
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 58), _
                Address:=wbkOpenGPP.Path & "\Copie de 1- ARIZ suiviReporting Global 31 12 2015_conventions.xlsm", _
                SubAddress:="'COOPERATIVE BANK - CKE105701'!A1", _
                TextToDisplay:="cliquez ici"
                
            ElseIf shtPrin.Cells(nR, 13).Value = "CBF120102" Then
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 58), _
                Address:=wbkOpenGPP.Path & "\Copie de 1- ARIZ suiviReporting Global 31 12 2015_conventions.xlsm", _
                SubAddress:="'CORIS - CBF120102'!A1", _
                TextToDisplay:="cliquez ici"
                
            ElseIf shtPrin.Cells(nR, 13).Value = "CBF119002" Then
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 58), _
                Address:=wbkOpenGPP.Path & "\Copie de 1- ARIZ suiviReporting Global 31 12 2015_conventions.xlsm", _
                SubAddress:="'BOA - CBF119002'!A1", _
                TextToDisplay:="cliquez ici"
                
            ElseIf shtPrin.Cells(nR, 13).Value = "CZZ176201" Then
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 58), _
                Address:=wbkOpenGPP.Path & "\Copie de 1- ARIZ suiviReporting Global 31 12 2015_conventions.xlsm", _
                SubAddress:="'GARRIGUE - CZZ176201'!A1", _
                TextToDisplay:="cliquez ici"
                
            ElseIf shtPrin.Cells(nR, 13).Value = "CGA117901" Then
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 58), _
                Address:=wbkOpenGPP.Path & "\Copie de 1- ARIZ suiviReporting Global 31 12 2015_conventions.xlsm", _
                SubAddress:="'ORABANK - CGA117901'!A1", _
                TextToDisplay:="cliquez ici"
                
            ElseIf shtPrin.Cells(nR, 13).Value = "CGH111401" Then
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 58), _
                Address:=wbkOpenGPP.Path & "\Copie de 1- ARIZ suiviReporting Global 31 12 2015_conventions.xlsm", _
                SubAddress:="'SG-SSB - CGH111401'!A1", _
                TextToDisplay:="cliquez ici"
                
            ElseIf shtPrin.Cells(nR, 13).Value = "CGH114101" Then
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 58), _
                Address:=wbkOpenGPP.Path & "\Copie de 1- ARIZ suiviReporting Global 31 12 2015_conventions.xlsm", _
                SubAddress:="'SG-SSB - CGH114101'!A1", _
                TextToDisplay:="cliquez ici"
                                
            Else
                ' the worksheet that should be chosen
                shtName = "'" & shtGPP.Cells(mR, 2).Value & " - " & shtGPP.Cells(mR, 3).Value & "'"
                Debug.Print (shtName)
                
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add Anchor:=Cells(nR, 58), _
                Address:=wbkOpenGPP.Path & "\Copie de 1- ARIZ suiviReporting Global 31 12 2015_conventions.xlsm", _
                SubAddress:=shtName & "!A1", _
                TextToDisplay:="cliquez ici"
            End If
    Next nR

    wbkOpenGPP.Close False

End Sub
