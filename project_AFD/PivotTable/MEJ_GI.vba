Public shtData As Worksheet
Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable

Sub MEJ_montant_engagement_garanti_GI()
    
    Set shtData = Worksheets("MEJ")
    Set shtSum = Worksheets("Feuil1")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:="MEJ!R1C1:R297C80")
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A5"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Type de garantie")
             .Orientation = xlPageField
             .Position = 1
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
            
        .CalculatedFields.Add "montant d'engagement garanti(en M€)", _
        "= 'Autorisation nette Montant garanti En €'/1000000", True
            
        With .PivotFields("montant d'engagement garanti(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "COTE D'IVOIRE"
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "AI"
    
End Sub

Sub MEJ_taux_de_sinistralité1_GI()
    
    Set shtData = Worksheets("MEJ")
    Set shtSum = Worksheets("Feuil1")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:="MEJ!R1C1:R297C80")
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A14"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Type de garantie")
             .Orientation = xlPageField
             .Position = 1
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
            
        .CalculatedFields.Add "taux de sinistralité 1", _
        "= 'Autorisation nette Montant garanti En €'/Autorisation nette Montant du prêt En €", True
            
        With .PivotFields("taux de sinistralité 1")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"

        End With
    
    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "COTE D'IVOIRE"
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "AI"
    
End Sub

Sub MEJ_montant_indemnisation_max_GI()
    
    Set shtData = Worksheets("MEJ")
    Set shtSum = Worksheets("Feuil1")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:="MEJ!R1C1:R297C80")
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A22"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Type de garantie")
             .Orientation = xlPageField
             .Position = 1
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
            
        .CalculatedFields.Add "montant d'indemnisation max(en M€)", _
        "= 'Max indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation max(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "COTE D'IVOIRE"
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "AI"
    
End Sub

Sub MEJ_taux_de_sinistralité2_GI()
    
    Set shtData = Worksheets("MEJ")
    Set shtSum = Worksheets("Feuil1")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:="MEJ!R1C1:R297C80")
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A30"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Type de garantie")
             .Orientation = xlPageField
             .Position = 1
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
            
        .CalculatedFields.Add "taux de sinistralité 2", _
        "= 'Max indemnisation en €'/Autorisation nette Montant du prêt En €", True
            
        With .PivotFields("taux de sinistralité 2")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"

        End With
    
    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "COTE D'IVOIRE"
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "AI"
    
End Sub

Sub MEJ_montant_indemnisation_réel_GI()
    
    Set shtData = Worksheets("MEJ")
    Set shtSum = Worksheets("Feuil1")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:="MEJ!R1C1:R297C80")
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A38"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Type de garantie")
             .Orientation = xlPageField
             .Position = 1
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
            
        .CalculatedFields.Add "montant d'indemnisation réel(en M€)", _
        "= 'Total indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation réel(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "COTE D'IVOIRE"
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "AI"
    
End Sub

Sub MEJ_taux_de_sinistralité3_GI()
    
    Set shtData = Worksheets("MEJ")
    Set shtSum = Worksheets("Feuil1")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:="MEJ!R1C1:R297C80")
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A46"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Type de garantie")
             .Orientation = xlPageField
             .Position = 1
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
            
        .CalculatedFields.Add "taux de sinistralité 3", _
        "= 'Total indemnisation en €'/Autorisation nette Montant du prêt En €", True
            
        With .PivotFields("taux de sinistralité 3")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"

        End With
    
    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "COTE D'IVOIRE"
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "AI"
    
End Sub
