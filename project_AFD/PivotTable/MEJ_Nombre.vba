Sub MEJ_Nombre()

    Dim shtData As Worksheet
    Dim shtSum As Worksheet
    Dim pvCache As PivotCache
    Dim pvTable As PivotTable

    Set shtData = Worksheets("MEJ")
    Set shtSum = Worksheets("Feuil1")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:="MEJ!R1C1:R297C83")
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AK6"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Type de garantie")
             .Orientation = xlRowField
             .Position = 1
        End With
        
        With .PivotFields("Type de garantie")
             .PivotItems("AG").Visible = False
        End With
        
        With .PivotFields("Année d'autorisation")
             .Orientation = xlColumnField
             .Position = 1
        End With
                
        With .PivotFields("Année d'autorisation")
             .PivotItems("1998").Visible = False
             .PivotItems("1999").Visible = False
             .PivotItems("2001").Visible = False
             .PivotItems("2004").Visible = False
             .PivotItems("2005").Visible = False
             .PivotItems("2006").Visible = False
             .PivotItems("2007").Visible = False
        End With
            
        .AddDataField .PivotFields("Total indemnisation en €"), _
                           "Nombre de demande", xlCount

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "COTE D'IVOIRE"
    
End Sub
