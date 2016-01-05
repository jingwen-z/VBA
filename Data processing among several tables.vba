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
