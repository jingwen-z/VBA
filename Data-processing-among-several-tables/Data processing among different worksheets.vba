'Chapiter 4 Data processing among several tables
'Section 1 Data processing among different worksheets  

'Question 61 How to generate specified quantity of worksheets according to the template  
Sub Generating_worksheets_according_to_the_template()  
    Dim rowN As Long  
    Dim shtOld As Worksheet  
    Dim shtNew As Worksheet  
    Dim shtTemplate As Worksheet  
      
    'setting template worksheet  
    Set shtTemplate = Sheet2  
    'setting data worksheet  
    Set shtOld = Sheet1  
      
    'go through all data  
    For rowN = 2 To 11  
        'create a duplicate of template worksheet, and the put it after the last worksheet  
        shtTemplate.Copy After:=Worksheets(Worksheets.Count)  
        'variable assignment,then pointing it to the new worksheet  
        Set shtNew = Worksheets(Worksheets.Count)  
        'changing  name of worksheet  
        shtNew.Name = shtOld.Cells(rowN, "A").Value  
        'reading data and assignment  
        shtNew.Cells(3, "A").Value = shtOld.Cells(rowN, "A").Value  
        shtNew.Cells(9, "A").Value = shtOld.Cells(rowN, "B").Value  
        shtNew.Cells(9, "E").Value = shtOld.Cells(rowN, "C").Value  
    Next  
      
End Sub  


'Question 62 cross-worksheet data query
'variable's public declaration for storing the searching initial cell
'public declaration is for all steps
'In this case, if we enter "excel",we can get client whose name with "excel"; if not, we can only get the first client "Shijiazhuang..."
Dim rngFind As Range
Sub cross_worksheet_query()
    'variable declaration
    Dim shtData As Worksheet
    Dim shtQuery As Worksheet
    Dim sKey As String
    Dim i As Long
    'getting worksheet variable
    Set shtQuery = Sheet1
    Set shtData = Sheet4
    'searching key words
    sKey = shtQuery.Cells(5, "B")
    
    With shtData
    
        'if the searching initial cell is not defined, then setting cell"A1" as the initial one
        If rngFind Is Nothing Then
            Set rngFind = .Range("A1")
        End If
        'looking for the cells
        Set rngFind = .Cells.Find(sKey, rngFind, lookat:=xlPart)
        'if we do not find it
        If rngFind Is Nothing Then
            'deleting the data in query worksheet
            For i = 9 To 12
                shtQuery.Cells(i, "B") = ""
            Next
    
        'if we find it
        Else
            'fill in the information in query sheet
            For i = 9 To 12
                shtQuery.Cells(i, "B") = .Cells(rngFind.Row, i - 8)
            Next
            'setting next initial cell as the end of the current cell
            Set rngFind = Intersect(rngFind.EntireRow, .Columns("D:D"))
    
        End If
    
    End With
End Sub

'Question 63 cross-worksheet data entry
Sub cross_worksheet_data_entry()
    Dim lastRow As Long
    Dim lstData As ListObject
    Dim rngTitle As Range
    
    'setting ListObject variable
    Set lstData = Sheet3.ListObjects(1)
    'focus on lstData
    With lstData
        'when Cell Area exists
        If (Not .DataBodyRange Is Nothing) Then
            lastRow = .DataBodyRange.Rows.Count
        'otherwise
        Else
            .ListRows.Add
            lastRow = 0
        End If
    End With
    'for testing
    Debug.Print lastRow
    
    'go through all titles in "data entry"
    For Each rngTitle In Union(Sheet2.Range("A4:A10"), Sheet2.Range("C7:C10"))
        'assinging the value of "data entry" to the relevant columns in "data sheet"
        lstData.ListColumns(rngTitle.Value).DataBodyRange(lastRow).Offset(1, 0).Value = rngTitle.Offset(0, 1).Value
    Next rngTitle
End Sub

'Question 64 creating batch of hyperlinks according to key words
Sub creating_batch_of_hyperlinks()
    Dim Sht1 As Worksheet   'balance sheet
    Dim Sht2 As Worksheet   'report item description
    Dim RngAll As Range     'ranges for hyperlink in balance sheet
    Dim Rng1 As Range
    Dim Rng2 As Range
    
    'setting worksheets
    Set Sht1 = Sheet5
    Set Sht2 = Sheet6
    
    'delete all hyperlinks
    Sht1.Hyperlinks.Delete
    Sht2.Hyperlinks.Delete
    Sht2.Columns(2).Clear
    
    
    'variable assignment
    'range area where we need to create hyperlinks
    Set RngAll = Union(Sht1.Range("A3:A13"), Sht1.Range("D3:D13"))
    'go through all ranges
    For Each Rng1 In RngAll
        'when there is content in the cell
        If Rng1 <> "" Then
            'try to find the same content in the first column of "report item description"
            Set Rng2 = Sht2.Range("A:A").Find(Rng1.Value, Lookat:=xlWhole)
            'if we can find it, then create hyperlinks
            If Not Rng2 Is Nothing Then
                'creating hyperlink
                Sht1.Hyperlinks.Add Rng1, "", Sht2.Name & Rng2.Address
                Sht2.Hyperlinks.Add Rng2.Offset(0, 1), "", Sht1.Name & Rng1.Address, "", "return"
                'setting format
                Rng1.Font.Size = 9
            End If
        End If
    Next

End Sub
