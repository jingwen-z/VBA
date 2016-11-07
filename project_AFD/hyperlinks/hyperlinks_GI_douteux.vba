Sub hyperlinks_GI_douteux()

    Dim wbkPrin As Workbook
    Dim wbkGI As Workbook
    Dim shtPrin As Worksheet
    Dim shtGI As Worksheet
    Dim matchedRow  As Variant
    Dim rowN As Long
    Dim rw As Long
    
    Set wbkPrin = ThisWorkbook
    Set wbkGI = Workbooks.Open("full path of workbook")

    Set shtPrin = wbkPrin.Sheets("Base de données")
    Set shtGI = wbkGI.Sheets("Provisions_GI_au_30_09_2016")

    shtPrin.Columns(56).Hyperlinks.Delete

    rowN = shtPrin.Cells(Rows.Count, 13).End(xlUp).Row

    For rw = 4 To rowN
        matchedRow = Application.Match(shtPrin.Cells(rw, 13).Value, shtGI.Columns(13), 0)

            If IsError(matchedRow) Then
                shtPrin.Cells(rw, 56).Value = ""
            Else
                Windows(wbkPrin.Name).Activate
                
                shtPrin.Hyperlinks.Add _
                Anchor:=Cells(rw, 56), _
                Address:=wbkGI.Path & "name of workbook", _
                SubAddress:=shtGI.Name & "!" & "A" & matchedRow & ":AY" & matchedRow, _
                TextToDisplay:="cliquez ici"            
            End If
    Next rw

    wbkGI.Close False
    
End Sub
