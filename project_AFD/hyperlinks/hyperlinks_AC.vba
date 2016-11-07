Sub hyperlinks_AC()

    Dim wbkPrin As Workbook
    Dim wbkAC As Workbook
    Dim shtPrin As Worksheet
    Dim shtAC As Worksheet
    Dim matchedRow As Variant
    Dim rowN As Long
    Dim rw As Long
    
    Set wbkPrin = ThisWorkbook
    Set wbkAC = Workbooks.Open("full path of workbook")

    Set shtPrin = wbkPrin.Sheets("name of worksheet")
    Set shtAC = wbkAC.Sheets("name of worksheet")
    
    shtPrin.Columns(55).Hyperlinks.Delete

    rowN = shtPrin.Cells(Rows.Count, 13).End(xlUp).Row

    For rw = 3 To rowN
        matchedRow = Application.Match(shtPrin.Cells(rw, 13).Value, shtAC.Columns(13), 0)
            If IsError(matchedRow) Then
                shtPrin.Cells(rw, 55).Value = ""
            Else
                Windows(wbkPrin.Name).Activate

                shtPrin.Hyperlinks.Add _
                Anchor:=Cells(rw, 55), _
                Address:=wbkAC.Path & "\name of workbook", _
                SubAddress:=shtAC.Name & "!" & "A" & matchedRow & ":BI" & matchedRow, _
                TextToDisplay:="cliquez ici"
            End If
    Next rw
    
    wbkAC.Close False

End Sub
