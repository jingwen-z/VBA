' Module 11
Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable
Public Const source = "MEJ_Globale!R3C1:R335C82"

Sub Global_MEJ_nombre()
    
    Set shtSum = Worksheets("TCD_global")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A3"))
    
    With pvTable
        With .PivotFields("Type de garantie")
             .Orientation = xlRowField
             .Position = 1
        End With
        
        With .PivotFields("Type de garantie")
             .PivotItems("AG").Visible = False
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
            
        .AddDataField .PivotFields("Total indemnisation en €"), _
                           "Nombre de demande GI", xlCount

    End With
End Sub

Sub Global_MEJ_perte_calculée_par_banque()
    
    Set shtSum = Worksheets("TCD_global")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A12"))
    
    With pvTable
        With .PivotFields("Type de garantie")
             .Orientation = xlRowField
             .Position = 1
        End With
        
        With .PivotFields("Type de garantie")
             .PivotItems("AG").Visible = False
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
        "= 'Perte provisoire calculée par la banque en euro'/1000000", True
            
        With .PivotFields("perte provisoire calculée par la banque(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
End Sub

Sub Global_MEJ_montant_indemnisation_max()

    Set shtSum = Worksheets("TCD_global")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A21"))
    
    With pvTable
        With .PivotFields("Type de garantie")
             .Orientation = xlRowField
             .Position = 1
        End With
        
        With .PivotFields("Type de garantie")
             .PivotItems("AG").Visible = False
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
            
        .CalculatedFields.Add "montant d'indemnisation max(en M€)", _
        "= 'Max indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation max(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With

End Sub
