Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable
Public Const pays = "XXX"
Public Const source = "xxx!RxCx"

Sub Pays_Octroi_GP_banque_Somme()
    
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
        
        With .PivotFields("Banque")
             .Orientation = xlRowField
             .Position = 1
        End With
        
        With .PivotFields("Année d'autorisation")
             .Orientation = xlColumnField
             .Position = 1
        End With
                          
        .CalculatedFields.Add "Octroi GP(en M€)", _
        "= Montant d'enveloppe en EUR/1000000", True
            
        With .PivotFields("Octroi GP(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
End Sub

Sub Pays_Encours_GP_banque_Somme()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A24"))
    
    With pvTable
        
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Banque")
             .Orientation = xlRowField
             .Position = 1
        End With
        
        With .PivotFields("Année d'autorisation")
             .Orientation = xlColumnField
             .Position = 1
        End With
        
        .CalculatedFields.Add "Encours au 31/12/2015(en M€)", _
        "= 'Encours de Garanties Sous-Participées en Euro11'/1000000", True
            
        With .PivotFields("Encours au 31/12/2015(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
End Sub

Sub Pays_GP_taux_utilisation_banque()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A42"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Banque")
             .Orientation = xlRowField
             .Position = 1
        End With
        
        With .PivotFields("Année d'autorisation")
             .Orientation = xlColumnField
             .Position = 1
        End With
                          
        .CalculatedFields.Add "Taux utilisation", _
       "='Montant d''engagement initial en euro' /'Montant d''enveloppe en EUR'", True
            
        With .PivotFields("Taux utilisation")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
End Sub

Sub Pays_GP_nombre()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A60"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
                
        With .PivotFields("Année d'autorisation")
             .Orientation = xlColumnField
             .Position = 1
        End With
                          
        .AddDataField .PivotFields("Montant d'enveloppe en EUR"), _
                           "GP(en nombre)", xlCount
    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
End Sub

Sub Pays_GP_montant_enveloppe()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A68"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
                
        With .PivotFields("Année d'autorisation")
             .Orientation = xlColumnField
             .Position = 1
        End With
                          
        .CalculatedFields.Add "Octroi GP(en M€)", _
        "= Montant d'enveloppe en EUR/1000000", True
            
        With .PivotFields("Octroi GP(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With
    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
End Sub
