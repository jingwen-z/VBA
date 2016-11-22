Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable
Public Const groupe_bancaire = "SOCIETE GENERALE"
Public Const source = "Base de données!R2C1:R1629C54"

Sub GrpBq_Octroi_GI_GP_sum()

    Set shtSum = Worksheets("TCD")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AA5"))
    
    With pvTable
            
        With .PivotFields("Groupe Bancaire")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .Orientation = xlRowField
             .Position = 1
        End With
    
        With .PivotFields("Année d'octroi")
             .Orientation = xlColumnField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .PivotItems("AG").Visible = False
             .PivotItems("FP").Visible = False
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
        
        .CalculatedFields.Add "Octroi (en M€) GI et GP", _
        "= 'Montant garanti en €2'/1000000", True
            
        With .PivotFields("Octroi (en M€) GI et GP")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Groupe Bancaire").ClearAllFilters
    pvTable.PivotFields("Groupe Bancaire").CurrentPage = groupe_bancaire

End Sub

Sub GrpBq_Encours_GI_GP_sum()
    
    Set shtSum = Worksheets("TCD")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AA14"))
    
    With pvTable
        
        With .PivotFields("Groupe Bancaire")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .Orientation = xlRowField
             .Position = 1
        End With
    
        With .PivotFields("Année d'octroi")
             .Orientation = xlColumnField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .PivotItems("AG").Visible = False
             .PivotItems("FP").Visible = False
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
                    
        .CalculatedFields.Add "Encours 30/06/2016(en M€) GI et GP", _
        "= 'Encours de risque DBO au 30/06/2016                                     (maj 05/08/2016)'/1000000", True
            
        With .PivotFields("Encours 30/06/2016(en M€) GI et GP")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With
    
    End With
    
    pvTable.PivotFields("Groupe Bancaire").ClearAllFilters
    pvTable.PivotFields("Groupe Bancaire").CurrentPage = groupe_bancaire
    
End Sub

Sub GrpBq_Octroi_GI_pays_sum()
    
    Set shtSum = Worksheets("TCD")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AA24"))
    
    With pvTable
        
        With .PivotFields("Groupe Bancaire")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Pays")
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
        "= 'Montant garanti en €2'/1000000", True
            
        With .PivotFields("Octroi GI(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With
        
    End With
    
    pvTable.PivotFields("Groupe Bancaire").ClearAllFilters
    pvTable.PivotFields("Groupe Bancaire").CurrentPage = groupe_bancaire
    
    pvTable.PivotFields("AG/GI/SP/FP").ClearAllFilters
    pvTable.PivotFields("AG/GI/SP/FP").CurrentPage = "GI"
    
End Sub

Sub GrpBq_Encours_GI_pays_sum()
    
    Set shtSum = Worksheets("TCD")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AA51"))
    
    With pvTable
        
        With .PivotFields("Groupe Bancaire")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Pays")
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
                
        .CalculatedFields.Add "Encours 30/06/2016(en M€)", _
        "= 'Encours de risque DBO au 30/06/2016                                     (maj 05/08/2016)'/1000000", True
            
        With .PivotFields("Encours 30/06/2016(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With
                
    End With
    
    pvTable.PivotFields("Groupe Bancaire").ClearAllFilters
    pvTable.PivotFields("Groupe Bancaire").CurrentPage = groupe_bancaire
    
    pvTable.PivotFields("AG/GI/SP/FP").ClearAllFilters
    pvTable.PivotFields("AG/GI/SP/FP").CurrentPage = "GI"
    
End Sub

Sub GrpBq_Octroi_GI_pays_moyenne()
    
    Set shtSum = Worksheets("TCD")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AA78"))
    
    With pvTable
        
        With .PivotFields("Groupe Bancaire")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Pays")
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
            
        .AddDataField .PivotFields("Montant garanti en €2"), _
                           "Moyenne Octroi GI(en €)", xlAverage

    End With
    
    pvTable.PivotFields("Groupe Bancaire").ClearAllFilters
    pvTable.PivotFields("Groupe Bancaire").CurrentPage = groupe_bancaire
    
    pvTable.PivotFields("AG/GI/SP/FP").ClearAllFilters
    pvTable.PivotFields("AG/GI/SP/FP").CurrentPage = "GI"
    
End Sub

Sub GrpBq_Encours_GI_pays_moyenne()
    
    Set shtSum = Worksheets("TCD")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AA105"))
    
    With pvTable
        
        With .PivotFields("Groupe Bancaire")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Pays")
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
            
        .AddDataField .PivotFields("Encours de risque DBO au 30/06/2016                                     (maj 05/08/2016)"), _
                           "Moyenne Encours 30/06/2016(en €)", xlAverage

    End With
    
    pvTable.PivotFields("Groupe Bancaire").ClearAllFilters
    pvTable.PivotFields("Groupe Bancaire").CurrentPage = groupe_bancaire
    
    pvTable.PivotFields("AG/GI/SP/FP").ClearAllFilters
    pvTable.PivotFields("AG/GI/SP/FP").CurrentPage = "GI"
    
End Sub

Sub GrpBq_Octroi_GI_pays_nombre()
    
    Set shtSum = Worksheets("TCD")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AA132"))
    
    With pvTable
        
        With .PivotFields("Groupe Bancaire")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Pays")
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
            
        .AddDataField .PivotFields("Montant garanti en €2"), _
                           "Octroi GI(en nombre)", xlCount

    End With
    
    pvTable.PivotFields("Groupe Bancaire").ClearAllFilters
    pvTable.PivotFields("Groupe Bancaire").CurrentPage = groupe_bancaire
    
    pvTable.PivotFields("AG/GI/SP/FP").ClearAllFilters
    pvTable.PivotFields("AG/GI/SP/FP").CurrentPage = "GI"
    
End Sub

Sub GrpBq_Octroi_GP_pays_sum()
    
    Set shtSum = Worksheets("TCD")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AA159"))
    
    With pvTable
        
        With .PivotFields("Groupe Bancaire")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Pays")
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
            
        .CalculatedFields.Add "Octroi GP(en M€)", _
        "= 'Montant garanti en €2'/1000000", True
            
        With .PivotFields("Octroi GP(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With
            
    End With
    
    pvTable.PivotFields("Groupe Bancaire").ClearAllFilters
    pvTable.PivotFields("Groupe Bancaire").CurrentPage = groupe_bancaire
    
    pvTable.PivotFields("AG/GI/SP/FP").ClearAllFilters
    pvTable.PivotFields("AG/GI/SP/FP").CurrentPage = "SP"
    
End Sub

Sub GrpBq_Encours_GP_pays_sum()
    
    Set shtSum = Worksheets("TCD")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AA185"))
    
    With pvTable
        
        With .PivotFields("Groupe Bancaire")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Pays")
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
                
        .CalculatedFields.Add "Encours 30/06/2016(en M€)", _
        "= 'Encours de risque DBO au 30/06/2016                                     (maj 05/08/2016)'/1000000", True
            
        With .PivotFields("Encours 30/06/2016(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With
                
    End With
    
    pvTable.PivotFields("Groupe Bancaire").ClearAllFilters
    pvTable.PivotFields("Groupe Bancaire").CurrentPage = groupe_bancaire
    
    pvTable.PivotFields("AG/GI/SP/FP").ClearAllFilters
    pvTable.PivotFields("AG/GI/SP/FP").CurrentPage = "SP"
    
End Sub

Sub GrpBq_Octroi_GP_pays_moyenne()
    
    Set shtSum = Worksheets("TCD")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AA211"))
    
    With pvTable
        
        With .PivotFields("Groupe Bancaire")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Pays")
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
            
        .AddDataField .PivotFields("Montant garanti en €2"), _
                           "Moyenne Octroi GP(en €)", xlAverage

    End With
    
    pvTable.PivotFields("Groupe Bancaire").ClearAllFilters
    pvTable.PivotFields("Groupe Bancaire").CurrentPage = groupe_bancaire
    
    pvTable.PivotFields("AG/GI/SP/FP").ClearAllFilters
    pvTable.PivotFields("AG/GI/SP/FP").CurrentPage = "SP"
    
End Sub

Sub GrpBq_Encours_GP_pays_moyenne()
    
    Set shtSum = Worksheets("TCD")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AA237"))
    
    With pvTable
        
        With .PivotFields("Groupe Bancaire")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Pays")
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
            
        .AddDataField .PivotFields("Encours de risque DBO au 30/06/2016                                     (maj 05/08/2016)"), _
                           "Moyenne Encours 30/06/2016(en €)", xlAverage

    End With
    
    pvTable.PivotFields("Groupe Bancaire").ClearAllFilters
    pvTable.PivotFields("Groupe Bancaire").CurrentPage = groupe_bancaire
    
    pvTable.PivotFields("AG/GI/SP/FP").ClearAllFilters
    pvTable.PivotFields("AG/GI/SP/FP").CurrentPage = "SP"
    
End Sub

Sub GrpBq_Octroi_GP_pays_nombre()
    
    Set shtSum = Worksheets("TCD")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AA263"))
    
    With pvTable
        With .PivotFields("Groupe Bancaire")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Pays")
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
            
        .AddDataField .PivotFields("Montant garanti en €2"), _
                           "Octroi GP(en nombre)", xlCount

    End With
    
    pvTable.PivotFields("Groupe Bancaire").ClearAllFilters
    pvTable.PivotFields("Groupe Bancaire").CurrentPage = groupe_bancaire
    
    pvTable.PivotFields("AG/GI/SP/FP").ClearAllFilters
    pvTable.PivotFields("AG/GI/SP/FP").CurrentPage = "SP"
    
End Sub
