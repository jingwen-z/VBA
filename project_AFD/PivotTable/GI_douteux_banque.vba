Sub GI_douteux_banque()

    Dim shtData As Worksheet
    Dim shtSum As Worksheet
    Dim pvCache As PivotCache
    Dim pvTable As PivotTable

    Set shtData = Worksheets("GI")
    Set shtSum = Worksheets("Feuil1")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=shtData.Range("A1").CurrentRegion)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A6"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With

        With .PivotFields("Indicateur sain/douteux détaillé au 30/09/15")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Indicateur sain/douteux détaillé au 30/09/15")
             .PivotItems("0").Visible = False
             .PivotItems("Garantie échue").Visible = False
             .PivotItems("Pas encore décaissé").Visible = False
             .PivotItems("Prêt échu").Visible = False
             .PivotItems("S").Visible = False
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
             .Orientation = xlRowField
             .Position = 1
        End With
            
        .CalculatedFields.Add "Montant garanti(en M€)", _
        "= 'Autorisation nette Montant garanti en €'/1000000", True
            
        With .PivotFields("Montant garanti(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.000"
        End With

        .CalculatedFields.Add "Encours(en M€)", _
        "= 'Encours de risque au 31/03/2016 en €'/1000000", True
            
        With .PivotFields("Encours(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

        .CalculatedFields.Add "Provision(en M€)", _
        "= 'Provision au 31/03/2016 en €'/1000000", True
            
        With .PivotFields("Provision(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "COTE D'IVOIRE"
    
End Sub
