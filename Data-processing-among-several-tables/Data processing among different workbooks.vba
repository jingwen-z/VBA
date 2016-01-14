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

'Question 66 How to read data from other  workbook?
Sub reading_data_across_workbook()
    Dim wbkThis As Workbook   'the workbook at present
    Dim wbkOpen As Workbook   'the workbook that we will open
    
    'set object Viewer
    Set wbkThis = ThisWorkbook
    'set the workbook that we should open
    Set wbkOpen = Workbooks.Open(wbkThis.Path & "\database.xlsx")
    'copy the first worksheet's content of opened workbook into the workbook at present
    wbkOpen.Worksheets(1).Cells.Copy wbkThis.Worksheets(1).Range("A1")
    'close the opened worksheet and save nothing
    wbkOpen.Close False
    
    'name the worksheet in ThisWorkbook
    wbkThis.Worksheets(1).Name = "read data across workbook"
End Sub

'Question 67 How to import worksheets from different workbooks?
Sub batch_importing_worksheets()
    Dim fileName As String        'name of file
    Dim filePath As String        'path of file
    Dim wbkThis As Workbook       'the workbook at present
    Dim wbkOpen As Workbook       'workbook to be opened
    Dim shtNew As Worksheet       'new data worksheet
    Dim shtData As Worksheet      'original data worksheet
    
    'stop screenupdating
    Application.ScreenUpdating = False
    
    'set the workbook at present
    Set wbkThis = ThisWorkbook
    
    'stop alert display
    Application.DisplayAlerts = False
    
        'delete all the worksheets except for Sheet1 in the workbook at present
        For Each shtNew In wbkThis.Worksheets
            If shtNew.Name <> Sheet1.Name Then
                shtNew.Delete
            End If
        Next
    
    'start alert display
    Application.DisplayAlerts = True
    
    'get the path of file at present
    filePath = wbkThis.Path
    'look for .xlsx workbooks
    fileName = Dir(filePath & "\*.xlsx")
    'check fileName
    'Debug.Print fileName
    
    'loop if the result is not empty
    Do While fileName <> ""
        
        'open workbook
        Set wbkOpen = Workbooks.Open(filePath & "\" & fileName)
        'set original data worksheet
        Set shtData = wbkOpen.Worksheets(1)
        'insert worksheet into the workbook at present
        Set shtNew = wbkThis.Worksheets.Add(after:=wbkThis.Worksheets(wbkThis.Worksheets.Count))
        'change the name of worksheet, and rename it with the name of workbook
        shtNew.Name = Left(fileName, 4)
        'copy original data into new worksheet
        shtData.Cells.Copy shtNew.Range("A1")
        
        'close the workbook
        wbkOpen.Close False
        'continu to find the next workbook
        fileName = Dir
        
    Loop
    
    'start screenupdating
    Application.ScreenUpdating = True
    
End Sub
