Public shtData As Worksheet
Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable

Sub MEJ_montant_max_nature()
    
    Set shtData = Worksheets("MEJ")
    Set shtSum = Worksheets("Feuil1")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:="MEJ!R1C1:R297C83")
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AH23"))
    
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
                
        With .PivotFields("Nature prêt")
             .Orientation = xlRowField
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
        
        With .PivotFields("Nature prêt")
             .PivotItems("0").Visible = False
             .PivotItems("#N/A").Visible = False
        End With

        .AddDataField .PivotFields("Total indemnisation en €"), _
                           "MEJ (en M€) montant max(GI)", xlMax

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "COTE D'IVOIRE"
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "AI"
    
End Sub

Sub Taux_de_sinistralité_max_nature()

    Set shtData = Worksheets("MEJ")
    Set shtSum = Worksheets("Feuil1")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:="MEJ!R1C1:R297C83")
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AH33"))
    
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
                
        With .PivotFields("Nature prêt")
             .Orientation = xlRowField
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
        
        With .PivotFields("Nature prêt")
             .PivotItems("0").Visible = False
             .PivotItems("#N/A").Visible = False
        End With

        .AddDataField .PivotFields("Taux de sinistralité réel"), _
                           "Taux de sinistralité max", xlMax

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "COTE D'IVOIRE"
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "AI"
    
End Sub
