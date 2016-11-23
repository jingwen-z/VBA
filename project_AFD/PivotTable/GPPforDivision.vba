Sub Global_GP_montant_enveloppe()

    Dim shtSum As Worksheet
    Dim pvCache As PivotCache
    Dim pvTable As PivotTable
    
    Set shtSum = Worksheets("TCD")
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A75"))
    
    With pvTable
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
    
End Sub
