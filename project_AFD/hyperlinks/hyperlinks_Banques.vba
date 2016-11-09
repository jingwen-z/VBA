Sub hyperlinks_Banques()

    Dim wbkPrin As Workbook
    Dim wbkBanque As Workbook
    Dim shtPrin As Worksheet
    Dim shtBanque As Worksheet
    Dim matchedRow As Variant
    Dim rowN As Long
    Dim rw As Long
    
    Set wbkPrin = ThisWorkbook
    Set wbkBanque = Workbooks.Open("full path of workbook")

    Set shtPrin = wbkPrin.Sheets("name of worksheet")
    Set shtBanque = wbkBanque.Sheets("name of worksheet")

    shtPrin.Columns(58).Hyperlinks.Delete

    rowN = shtPrin.Cells(Rows.Count, 13).End(xlUp).Row    

    For rw = 3 To rowN
        matchedRow = Application.Match(shtPrin.Cells(rw, 20).Value, shtBanque.Columns(3), 0)
            If IsError(matchedRow) Then
                shtPrin.Cells(rw, 58).Value = ""
            Else
                Windows(wbkPrin.Name).Activate

                shtPrin.Hyperlinks.Add _
                Anchor:=Cells(rw, 58), _
                Address:=wbkBanque.Path & "\name of workbook", _
                SubAddress:=shtBanque.Name & "!" & "A" & matchedRow & ":U" & matchedRow, _
                TextToDisplay:="cliquez ici"
            End If
    Next rw
    
    wbkBanque.Close False

End Sub
