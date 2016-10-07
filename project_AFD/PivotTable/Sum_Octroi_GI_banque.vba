Public shtData As Worksheet
Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable

Sub Sum_Octroi_GI_banque()
    
    Set shtData = Worksheets("Table_Principale")
    Set shtSum = Worksheets("Feuil1")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=shtData.Range("A1").CurrentRegion)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A23"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
             .Orientation = xlRowField
             .Position = 1
        End With
    
        With .PivotFields("Année d'octroi")
             .Orientation = xlColumnField
             .Position = 1
        End With
                
        With .PivotFields("Année d'octroi")
             .PivotItems("1997").Visible = False
             .PivotItems("1998").Visible = False
             .PivotItems("1999").Visible = False
             .PivotItems("2000").Visible = False
             .PivotItems("2001").Visible = False
             .PivotItems("2002").Visible = False
             .PivotItems("2003").Visible = False
             .PivotItems("2004").Visible = False
             .PivotItems("2005").Visible = False
             .PivotItems("2006").Visible = False
             .PivotItems("2007").Visible = False
        End With
            
        .CalculatedFields.Add "Octroi GI(en M€)", _
        "= 'Autorisation nette Montant garanti en €'/1000000", True
            
        With .PivotFields("Octroi GI(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "COTE D'IVOIRE"
    
    pvTable.PivotFields("AG/GI/SP/FP").ClearAllFilters
    pvTable.PivotFields("AG/GI/SP/FP").CurrentPage = "GI"
    
End Sub

Sub Sum_Encours_GI_banque()
    
    Set shtData = Worksheets("Table_Principale")
    Set shtSum = Worksheets("Feuil1")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=shtData.Range("A1").CurrentRegion)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A38"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .Orientation = xlPageField
             .Position = 1
        End With
                
        With .PivotFields("Bénéficiaire Primaire")
             .Orientation = xlRowField
             .Position = 1
        End With
                
        .CalculatedFields.Add "Encours actuel(en M€)", _
        "= 'Encours de risque DBO au 31/03/2016'/1000000", True
            
        With .PivotFields("Encours actuel(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "COTE D'IVOIRE"
    
    pvTable.PivotFields("AG/GI/SP/FP").ClearAllFilters
    pvTable.PivotFields("AG/GI/SP/FP").CurrentPage = "GI"
    
End Sub
