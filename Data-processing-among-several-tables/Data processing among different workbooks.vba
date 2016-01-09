'Section 2 Data processing among different workbooks

'Question 65 How to save several worksheets as different single workbooks, respectively?
Sub save_sheets_as_single_book()
    Dim fileName As String          'name of file
    Dim filePath As String          'path of file
    Dim fileFullName As String      'full name of file,including path and name
    Dim sht As Worksheet            'worksheet
    Dim wbkOld As Workbook          'the workbook at present
    Dim wbkNew As Workbook          'new workbook
    
    'set the variable "the workbook at present"
    Set wbkOld = ThisWorkbook
    'get the path of file
    filePath = wbkOld.Path
    
    'go through all worksheets
    For Each sht In wbkOld.Worksheets
        
        'get the name of file - ATTENTION: DO NOT FORGET ".xlsx"
        fileName = sht.Name & ".xlsx"
        'get the full name of file
        fileFullName = filePath & Application.PathSeparator & fileName
        'add workbook
        Set wbkNew = Workbooks.Add
        'copy "the worksheet at present" into the new one
        sht.Copy Before:=wbkNew.Worksheets(1)
        'to see if the file exists or not. If it exists, then we delete it
        If Dir(fileFullName) <> "" Then
            'delete the file
            Kill fileFullName
        End If
        'save the new workbook
        wbkNew.SaveAs fileFullName
        'close the new workbook
        wbkNew.Close
        
    Next sht
    
End Sub
