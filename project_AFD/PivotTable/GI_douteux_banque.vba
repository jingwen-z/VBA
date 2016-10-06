Public shtData As Worksheet
Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable

Sub GI_douteux_banque()
    
    Set shtData = Worksheets("GI")
    Set shtSum = Worksheets("Feuil1")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=shtData.Range("A1").CurrentRegion)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A5"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
             .Orientation = xlRowField
             .Position = 1
        End With
            
        .AddDataField .PivotFields("Autorisation nette Montant du prêt en €"), _
                           "Montant des prêts(en €)", xlSum
        .AddDataField .PivotFields("Encours de risque au 31/03/2016 en €"), _
                           "Encours(en €)", xlSum
        .AddDataField .PivotFields("Provision au 31/03/2016 en €"), _
                           "Provision(en €)", xlSum

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "COTE D'IVOIRE"
    
End Sub
