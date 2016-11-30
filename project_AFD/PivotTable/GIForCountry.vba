Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable
Public Const pays = "SENEGAL"
Public Const source = "Provisions_GI_au_30_09_2016!R3C1:R920C51"

Sub GI_provision_banque()
    
    Set shtSum = Worksheets("TCD")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A6"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Indicateur sain/douteux détaillé au 30/09/16")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Indicateur sain/douteux détaillé au 30/09/16")
             .PivotItems("Garantie échue").Visible = False
             .PivotItems("Prêt non décaissé").Visible = False
             .PivotItems("S").Visible = False
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
             .Orientation = xlRowField
             .Position = 1
        End With
            
        .CalculatedFields.Add "Montant garanti(en M€)", _
        "= 'Montant garanti en €2'/1000000", True
            
        With .PivotFields("Montant garanti(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.000"
        End With

        .CalculatedFields.Add "Encours(en M€)", _
        "= 'Encours de risque DBO au 30/06/2016                                     (maj 05/08/2016)'/1000000", True
            
        With .PivotFields("Encours(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.000"
        End With

        .CalculatedFields.Add "Provision(en M€)", _
        "= 'Provision au 30/09/2016 en €'/1000000", True
            
        With .PivotFields("Provision(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.000"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
End Sub
