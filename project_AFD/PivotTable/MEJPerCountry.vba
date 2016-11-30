' Module 1
Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable
Public Const pays = "SENEGAL"
Public Const source = "MEJ_Globale!R3C1:R335C82"

Sub Pays_MEJ_montant_engagement_garanti_GI()
    
    Set shtSum = ThisWorkbook.Worksheets("TCD")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A12"))
    
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
            
        .CalculatedFields.Add "montant d'engagement garanti(en M€)", _
        "= 'En €2'/1000000", True
            
        With .PivotFields("montant d'engagement garanti(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "AI"
    
End Sub

Sub Pays_MEJ_montant_indemnisation_max_GI()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A21"))
    
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
            
        .CalculatedFields.Add "montant d'indemnisation max(en M€)", _
        "= 'Max indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation max(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "AI"
    
End Sub

Sub Pays_MEJ_montant_indemnisation_réel_GI()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A30"))
    
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
            
        .CalculatedFields.Add "montant d'indemnisation réel(en M€)", _
        "= 'Total indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation réel(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "AI"
    
End Sub

Sub Pays_MEJ_perte_calculée_par_banque_GI()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("A39"))
    
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
        "= 'Perte provisoire calculée par la banque en euro'/1000000", True
            
        With .PivotFields("perte provisoire calculée par la banque(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "AI"
    
End Sub

' Module 2
Public wbkPrin As Workbook
Public shtMEJ As Worksheet
Public rowN As Long

Sub add_cols_nature_secteur()

    Dim rw As Long
    
    Set wbkPrin = Workbooks.Open(ThisWorkbook.Path & "\BDD Principale-TCD.xlsm")
    Set shtMEJ = ThisWorkbook.Sheets("MEJ_Globale")
    
    With shtMEJ
         .Columns("V:V").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
         .Columns("V:V").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("V3").FormulaR1C1 = "Nature prêt"
         .Range("W3").FormulaR1C1 = "Secteur détaillé"
    End With
    
    rowN = shtMEJ.Cells(Rows.Count, 6).End(xlUp).Row
    
    For rw = 4 To rowN
         shtMEJ.Cells(rw, 22).FormulaR1C1 = _
            "=VLOOKUP(RC[-16],'[BDD Principale-TCD.xlsm]Base de données'!C13:C45,33,0)"
         shtMEJ.Cells(rw, 23).FormulaR1C1 = _
            "=VLOOKUP(RC[-17],'[BDD Principale-TCD.xlsm]Base de données'!C13:C46,34,0)"
    Next rw
    
    wbkPrin.Close False
    
End Sub

Sub add_perte_par_banque_en_euro()

    Dim Rng As Range

    Set shtMEJ = ThisWorkbook.Sheets("MEJ_Globale")

    With shtMEJ
         .Columns("AQ:AQ").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("AQ3").FormulaR1C1 = "Perte provisoire calculée par la banque en euro"
    End With

    rowN = shtMEJ.Cells(Rows.Count, 6).End(xlUp).Row

    For Each Rng In shtMEJ.Range("AQ4:AQ" & rowN)
    
        If Rng.Offset(0, -1).Value = 0 Then
            Rng.Value = 0
        Else
            Rng.Value = "=RC[-1]/RC[-23]"
        End If
    Next Rng

End Sub

Sub add_max_indemnisation()

    Set shtMEJ = ThisWorkbook.Sheets("MEJ_Globale")
    
    rowN = shtMEJ.Cells(Rows.Count, 6).End(xlUp).Row

    With shtMEJ
         .Columns("BX:BX").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
         .Range("BX3").FormulaR1C1 = "Max indemnisation en €"

         .Range("BX4:BX" & rowN).FormulaR1C1 = _
            "=IF(RC[-8]<>"""",RC[-25]+RC[-8],RC[-25]*2)"
    End With

End Sub

' Module 3
Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable
Public Const pays = "SENEGAL"
Public Const source = "MEJ_Globale!R3C1:R335C82"

Sub Pays_MEJ_montant_engagement_garanti_GP()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("R12"))
    
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
            
        .CalculatedFields.Add "montant d'engagement garanti(en M€)", _
        "= 'En €2'/1000000", True
            
        With .PivotFields("montant d'engagement garanti(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "SP"
    
End Sub

Sub Pays_MEJ_montant_indemnisation_max_GP()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("R21"))
    
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
            
        .CalculatedFields.Add "montant d'indemnisation max(en M€)", _
        "= 'Max indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation max(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "SP"
    
End Sub

Sub Pays_MEJ_montant_indemnisation_réel_GP()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("R30"))
    
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
                            
        .CalculatedFields.Add "montant d'indemnisation réel(en M€) SP", _
        "= 'Total indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation réel(en M€) SP")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "SP"
    
End Sub

Sub Pays_MEJ_perte_calculée_par_banque_GP()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("R39"))
    
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
        "= 'Perte provisoire calculée par la banque en euro'/1000000", True
            
        With .PivotFields("perte provisoire calculée par la banque(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "SP"
    
End Sub

' Module 4
Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable
Public Const pays = "SENEGAL"
Public Const source = "MEJ_Globale!R3C1:R335C82"

Sub Pays_MEJ_montant_engagement_garanti_ALIOS()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE12"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
                            
        .CalculatedFields.Add "montant d'engagement garanti(en M€)", _
        "= 'En €2'/1000000", True
            
        With .PivotFields("montant d'engagement garanti(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With
        
    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
        
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "ALIOS"
        
End Sub

Sub Pays_MEJ_montant_indemnisation_max_ALIOS()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE21"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
            
        .CalculatedFields.Add "montant d'indemnisation max(en M€)", _
        "= 'Max indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation max(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "ALIOS"
    
End Sub

Sub Pays_MEJ_montant_indemnisation_réel_ALIOS()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE30"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
            
        .CalculatedFields.Add "montant d'indemnisation réel(en M€)", _
        "= 'Total indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation réel(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "ALIOS"
    
End Sub

Sub Pays_MEJ_perte_calculée_ALIOS()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE39"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
        "= 'Perte provisoire calculée par la banque en euro'/1000000", True
            
        With .PivotFields("perte provisoire calculée par la banque(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "ALIOS"
    
End Sub

' Module 5
Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable
Public Const pays = "SENEGAL"
Public Const source = "MEJ_Globale!R3C1:R335C82"

Sub Pays_MEJ_montant_engagement_garanti_BICIS()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE51"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
                            
        .CalculatedFields.Add "montant d'engagement garanti(en M€)", _
        "= 'En €2'/1000000", True
            
        With .PivotFields("montant d'engagement garanti(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With
        
    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
        
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "BICIS"
        
End Sub

Sub Pays_MEJ_montant_indemnisation_max_BICIS()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE60"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
            
        .CalculatedFields.Add "montant d'indemnisation max(en M€)", _
        "= 'Max indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation max(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "BICIS"
    
End Sub

Sub Pays_MEJ_montant_indemnisation_réel_BICIS()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE69"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
            
        .CalculatedFields.Add "montant d'indemnisation réel(en M€)", _
        "= 'Total indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation réel(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "BICIS"
    
End Sub

Sub Pays_MEJ_perte_calculée_BICIS()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE78"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
        "= 'Perte provisoire calculée par la banque en euro'/1000000", True
            
        With .PivotFields("perte provisoire calculée par la banque(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "BICIS"
    
End Sub

' Module 6
Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable
Public Const pays = "SENEGAL"
Public Const source = "MEJ_Globale!R3C1:R335C82"

Sub Pays_MEJ_montant_engagement_garanti_BOASénégal()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE90"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
                            
        .CalculatedFields.Add "montant d'engagement garanti(en M€)", _
        "= 'En €2'/1000000", True
            
        With .PivotFields("montant d'engagement garanti(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With
        
    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
        
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "BOA Sénégal"
        
End Sub

Sub Pays_MEJ_montant_indemnisation_max_BOASénégal()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE99"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
            
        .CalculatedFields.Add "montant d'indemnisation max(en M€)", _
        "= 'Max indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation max(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "BOA Sénégal"
    
End Sub

Sub Pays_MEJ_montant_indemnisation_réel_BOASénégal()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE108"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
            
        .CalculatedFields.Add "montant d'indemnisation réel(en M€)", _
        "= 'Total indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation réel(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "BOA Sénégal"
    
End Sub

Sub Pays_MEJ_perte_calculée_BOASénégal()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE117"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
        "= 'Perte provisoire calculée par la banque en euro'/1000000", True
            
        With .PivotFields("perte provisoire calculée par la banque(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "BOA Sénégal"
    
End Sub

' Module 7
Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable
Public Const pays = "SENEGAL"
Public Const source = "MEJ_Globale!R3C1:R335C82"

Sub Pays_MEJ_montant_engagement_garanti_CBAO()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE129"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
                            
        .CalculatedFields.Add "montant d'engagement garanti(en M€)", _
        "= 'En €2'/1000000", True
            
        With .PivotFields("montant d'engagement garanti(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With
        
    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
        
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "CBAO"
        
End Sub

Sub Pays_MEJ_montant_indemnisation_max_CBAO()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE138"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
            
        .CalculatedFields.Add "montant d'indemnisation max(en M€)", _
        "= 'Max indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation max(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "CBAO"
    
End Sub

Sub Pays_MEJ_montant_indemnisation_réel_CBAO()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE147"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
            
        .CalculatedFields.Add "montant d'indemnisation réel(en M€)", _
        "= 'Total indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation réel(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "CBAO"
    
End Sub

Sub Pays_MEJ_perte_calculée_CBAO()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE156"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
        "= 'Perte provisoire calculée par la banque en euro'/1000000", True
            
        With .PivotFields("perte provisoire calculée par la banque(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "CBAO"
    
End Sub

' Module 8
Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable
Public Const pays = "SENEGAL"
Public Const source = "MEJ_Globale!R3C1:R335C82"

Sub Pays_MEJ_montant_engagement_garanti_CréditduSénégal()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE168"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
                            
        .CalculatedFields.Add "montant d'engagement garanti(en M€)", _
        "= 'En €2'/1000000", True
            
        With .PivotFields("montant d'engagement garanti(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With
        
    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
        
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "CREDIT DU SENEGAL"
        
End Sub

Sub Pays_MEJ_montant_indemnisation_max_CréditduSénégal()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE177"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
            
        .CalculatedFields.Add "montant d'indemnisation max(en M€)", _
        "= 'Max indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation max(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "CREDIT DU SENEGAL"
    
End Sub

Sub Pays_MEJ_montant_indemnisation_réel_CréditduSénégal()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE186"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
            
        .CalculatedFields.Add "montant d'indemnisation réel(en M€)", _
        "= 'Total indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation réel(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "CREDIT DU SENEGAL"
    
End Sub

Sub Pays_MEJ_perte_calculée_CréditduSénégal()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE195"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
        "= 'Perte provisoire calculée par la banque en euro'/1000000", True
            
        With .PivotFields("perte provisoire calculée par la banque(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "CREDIT DU SENEGAL"
    
End Sub

' Module 9
Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable
Public Const pays = "SENEGAL"
Public Const source = "MEJ_Globale!R3C1:R335C82"

Sub Pays_MEJ_montant_engagement_garanti_SGBS()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE207"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
                            
        .CalculatedFields.Add "montant d'engagement garanti(en M€)", _
        "= 'En €2'/1000000", True
            
        With .PivotFields("montant d'engagement garanti(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With
        
    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
        
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "SGBS"
        
End Sub

Sub Pays_MEJ_montant_indemnisation_max_SGBS()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE216"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
            
        .CalculatedFields.Add "montant d'indemnisation max(en M€)", _
        "= 'Max indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation max(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "SGBS"
    
End Sub

Sub Pays_MEJ_montant_indemnisation_réel_SGBS()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE225"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
            
        .CalculatedFields.Add "montant d'indemnisation réel(en M€)", _
        "= 'Total indemnisation en €'/1000000", True
            
        With .PivotFields("montant d'indemnisation réel(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "SGBS"
    
End Sub

Sub Pays_MEJ_perte_calculée_SGBS()
    
    Set shtSum = Worksheets("TCD")

    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AE234"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Bénéficiaire Primaire")
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
        "= 'Perte provisoire calculée par la banque en euro'/1000000", True
            
        With .PivotFields("perte provisoire calculée par la banque(en M€)")
             .Orientation = xlDataField
             .NumberFormat = "#,##0.00"
        End With

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Bénéficiaire Primaire").ClearAllFilters
    pvTable.PivotFields("Bénéficiaire Primaire").CurrentPage = "SGBS"
    
End Sub

' Module 10
Public shtSum As Worksheet
Public pvCache As PivotCache
Public pvTable As PivotTable
Public Const pays = "SENEGAL"
Public Const source = "MEJ_Globale!R3C1:R335C82"

Sub Pays_MEJ_Nombre_GI()
    
    Set shtSum = Worksheets("TCD")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("AQ11"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
        
        With .PivotFields("Type de garantie")
             .Orientation = xlRowField
             .Position = 1
        End With
        
        With .PivotFields("Type de garantie")
             .PivotItems("AG").Visible = False
             .PivotItems("SP").Visible = False
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
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
End Sub

Sub Pays_MEJ_montant_max_grpBancaire()

    Set shtSum = Worksheets("TCD")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("BH11"))
    
    With pvTable
        With .PivotFields("Pays")
             .Orientation = xlPageField
             .Position = 1
        End With
                
        With .PivotFields("Année d'autorisation")
             .Orientation = xlColumnField
             .Position = 1
        End With
                
        With .PivotFields("Groupe Bancaire")
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
        
        .AddDataField .PivotFields("Total indemnisation en €"), _
                           "MEJ (en €) montant max", xlMax

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
End Sub

Sub Pays_MEJ_montant_max_nature()

    Set shtSum = Worksheets("TCD")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("BH33"))
    
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
                           "MEJ (en €) montant max(GI)", xlMax

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "AI"
    
End Sub

Sub Pays_MEJ_montant_max_secteur()

    Set shtSum = Worksheets("TCD")
    
    Set pvCache = ThisWorkbook.PivotCaches.Create( _
                SourceType:=xlDatabase, _
                SourceData:=source)
                
    Set pvTable = pvCache.CreatePivotTable(shtSum.Range("BH52"))
    
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
                
        With .PivotFields("Secteur détaillé")
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
        
        With .PivotFields("Secteur détaillé")
             .PivotItems("0").Visible = False
             .PivotItems("#N/A").Visible = False
        End With

        .AddDataField .PivotFields("Total indemnisation en €"), _
                           "MEJ (en M€) montant max(GI)", xlMax

    End With
    
    pvTable.PivotFields("Pays").ClearAllFilters
    pvTable.PivotFields("Pays").CurrentPage = pays
    
    pvTable.PivotFields("Type de garantie").ClearAllFilters
    pvTable.PivotFields("Type de garantie").CurrentPage = "AI"
    
End Sub
