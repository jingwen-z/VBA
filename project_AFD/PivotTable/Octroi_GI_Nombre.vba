Sub Octroi_GI_Nombre()
    
    Dim shtData As Worksheet
    Dim shtSum As Worksheet
    Dim pvCache As PivotCache
    Dim pvTable As PivotTable
    
    Set shtData = Worksheets("Table_Principale")
    Set shtSum = Worksheets("Feuil1")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:="Table_Principale!R1C1:R1563C54")
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A134"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .Orientation = xlRowField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .PivotItems("AG").Visible = False
             .PivotItems("SP").Visible = False
             .PivotItems("FP").Visible = False
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
            
        .AddDataField .PivotFields("Autorisation nette Montant garanti en €"), _
                           "Octroi GI(en nombre)", xlCount

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = "COTE D'IVOIRE"

End Sub
