Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable
Public Const source = "Base de données!R2C1:R1629C54"

Sub Global_Octroi_GI_GP_nombre()
    
    Set shtSum = Worksheets("TCD_global")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A3"))
    
    With pvTable
        With .PivotFields("AG/GI/SP/FP")
             .Orientation = xlRowField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .PivotItems("AG").Visible = False
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
            
        .AddDataField .PivotFields("Montant garanti en €2"), _
                           "Octroi GI et GP(en nombre)", xlCount

    End With

End Sub

Sub Global_Octroi_GI_GP_montant()
    
    Set shtSum = Worksheets("TCD_global")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A12"))
    
    With pvTable
        With .PivotFields("AG/GI/SP/FP")
             .Orientation = xlRowField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .PivotItems("AG").Visible = False
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
            
        .CalculatedFields.Add "Octroi GI et GP(en M€)", _
        "= 'Montant garanti en €2'/1000000", True
            
        With .PivotFields("Octroi GI et GP(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With

End Sub

Sub Global_Encours_GI_GP_montant()
    
    Set shtSum = Worksheets("TCD_global")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A21"))
    
    With pvTable
        With .PivotFields("AG/GI/SP/FP")
             .Orientation = xlRowField
             .Position = 1
        End With
        
        With .PivotFields("AG/GI/SP/FP")
             .PivotItems("AG").Visible = False
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
            
        .CalculatedFields.Add "Encours restant GI et GP(en M€)", _
        "= 'Encours de risque DBO au 30/06/2016                                     (maj 05/08/2016)'/1000000", True     ' justifier le nom de champ
            
        With .PivotFields("Encours restant GI et GP(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With
    End With

End Sub
