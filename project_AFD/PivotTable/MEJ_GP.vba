Sub MEJ_montant_indemnisation_réel_GP()
    
    Set shtData = Worksheets("MEJ")
    Set shtSum = Worksheets("Feuil1")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:="MEJ!R1C1:R297C80")
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("J6"))
    
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
                            
        .CalculatedFields.Add "montant d'indemnisation réel(en M€) SP", _
        "= 'Total indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation réel(en M€) SP")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "COTE D'IVOIRE"
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "SP"
    
End Sub

Sub MEJ_taux_de_sinistralité_GP()
    
    Set shtData = Worksheets("MEJ")
    Set shtSum = Worksheets("Feuil1")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:="MEJ!R1C1:R297C80")
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("J14"))
    
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
            
        .CalculatedFields.Add "taux de sinistralité GP", _
        "= 'Total indemnisation en €'/Autorisation nette Montant du prêt En €", True
            
        With .PivotFields("taux de sinistralité GP")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With
    
    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "COTE D'IVOIRE"
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "SP"
    
End Sub

