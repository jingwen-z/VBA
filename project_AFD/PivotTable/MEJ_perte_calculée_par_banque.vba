Sub MEJ_perte_calculée_par_banque_GI()
    
    Dim shtData As Worksheet
    Dim shtSum As Worksheet
    Dim pvCache As PivotCache
    Dim pvTable As PivotTable

    Set shtData = Worksheets("MEJ")
    Set shtSum = Worksheets("Feuil1")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:="MEJ!R1C1:R297C84")
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A33"))
    
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
            
        .CalculatedFields.Add "perte provisoire calculée par la banque(en M€)", _
        "= 'DI-Perte provisoire calculée par la banque en euro'/1000000", True
            
        With .PivotFields("perte provisoire calculée par la banque(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "COTE D'IVOIRE"
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "AI"
    
End Sub
