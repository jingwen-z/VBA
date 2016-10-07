Public shtData As Worksheet
Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable

Sub Somme_Octroi_GP_banque()
    
    Set shtData = Worksheets("GPP")
    Set shtSum = Worksheets("Feuil1")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:="GPP!R1C1:R81C176")
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A5"))
    
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
    pvTable.PivotFields("Pays").CurrentPage = "Côte d'Ivoire"
    
End Sub


Sub Somme_Encours_GP_banque()
    
    Set shtData = Worksheets("GPP")
    Set shtSum = Worksheets("Feuil1")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:="GPP!R1C1:R81C176")
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A15"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Banque")
             .Orientation = xlRowField
             .Position = 1
        End With
        
        .CalculatedFields.Add "Encours actuel(en M€)", _
        "= 'Encours de Garanties Sous-Participées en Euro11'/1000000", True
            
        With .PivotFields("Encours actuel(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "Côte d'Ivoire"
    
End Sub

Sub GP_taux_utilisation_banque()
    
    Set shtData = Worksheets("GPP")
    Set shtSum = Worksheets("Feuil1")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:="GPP!R1C1:R81C176")
                
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
                          
        .CalculatedFields.Add "Taux d'utilisation", _
        "= Montant d'engagement initial en euro/Montant d'enveloppe en EUR", True
            
        With .PivotFields("Taux d'utilisation")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "Côte d'Ivoire"
    
End Sub
