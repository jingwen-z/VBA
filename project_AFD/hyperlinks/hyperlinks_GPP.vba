Sub hyperlinks_GPP()

    Dim wbkPrin As Workbook
    Dim wbkGPP As Workbook
    Dim shtPrin As Worksheet
    Dim shtGPP As Worksheet
    Dim matchedRow  As Variant
    Dim rowN As Long
    Dim rw As Long
    
    Set wbkPrin = ThisWorkbook
    
    ' open workbook "GPP"
    ' the address is variable
    Set wbkGPP = Workbooks.Open("S:\EBC\GAR\3 - MIDDLE OFFICE\BASES DE DONNEES\1- ARIZ suiviReporting Global 30 06 2016.xlsm")
    
    ' define worksheets
    Set shtPrin = wbkPrin.Sheets("Base de données")
    Set shtGPP = wbkGPP.Sheets("BDD_GPP")
    
    ' get the last row's number in worksheet "shtPrin"
    rowN = shtPrin.Cells(Rows.Count, 13).End(xlUp).Row
    
    ' remove all hyperlinks of column 57
    shtPrin.Columns(57).Hyperlinks.Delete

    ' go through all N concours in BDD Principale
    For rw = 4 To rowN
        ' locate the row of target N concours
        ' the column number is variable
        matchedRow = Application.Match(shtPrin.Cells(rw, 13).Value, shtGPP.Columns(3), 0)
        
            ' for every cell that is not empty,
            ' search through all N concours in BDD_GPP
            If IsError(matchedRow) Then
                ' write nothing in the cell
                ' the column number is variable
                shtPrin.Cells(rw, 57).Value = ""
            Else
                ' active wbkPrin
                Windows(wbkPrin.Name).Activate
                ' create a hyperlink in the same row as
                ' the corresponding N concours for shtGPP
                ' the column number is variable
                shtPrin.Hyperlinks.Add _
                Anchor:=Cells(rw, 57), _
                Address:=wbkGPP.Path & "\1- ARIZ suiviReporting Global 30 06 2016.xlsm", _
                SubAddress:=shtGPP.Name & "!" & "A" & matchedRow & ":FS" & matchedRow, _
                TextToDisplay:="cliquez ici"
            End If
    Next rw

    wbkGPP.Close False
    
End Sub
