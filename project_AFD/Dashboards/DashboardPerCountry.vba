Public wbkOpen As Workbook
Public rw As Long
Public rowN As Long
Public cl As Long
Public colN As Long
Public n As Long

Sub Octroi_GI_et_GP()
    
    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")
    
    wbkOpen.Worksheets("TCD").Range("A6:K8").Copy ThisWorkbook.Worksheets("Feuil1").Range("B4")
    rowN = wbkOpen.Worksheets("TCD").Range("K8").Row - wbkOpen.Worksheets("TCD").Range("A6").Row
    
    For rw = 1 To rowN - 1
        wbkOpen.Worksheets("TCD").Range("A16:K16").Offset(rw - 1, 0).Copy
        ThisWorkbook.Worksheets("Feuil1").Range("B4:L4").Offset(2 * rw, 0).Insert Shift:=xlDown
        
        With ThisWorkbook.Worksheets("Feuil1").Range("B4:L4").Offset(2 * rw, 0)
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 5
                .TintAndShade = 0.399914548173467
                .Weight = xlThin
            End With
         
        End With
    Next rw
    
    wbkOpen.Worksheets("TCD").Range("A16:K16").Offset(rowN - 1, 0).Copy _
    ThisWorkbook.Worksheets("Feuil1").Range("B4:L4").Offset(2 * rowN, 0)
    
    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B4").FormulaR1C1 = ""
        .Range("B5").FormulaR1C1 = "Octroi(en M€) GI"
        .Range("B6").FormulaR1C1 = "GI Encours(en M€)"
        .Range("B7").FormulaR1C1 = "Octroi(en M€) GP"
        .Range("B8").FormulaR1C1 = "GP Encours(en M€)"
        .Range("K4").FormulaR1C1 = "2016 act."
        .Range("L4").FormulaR1C1 = "Total"
    End With
    
    wbkOpen.Close False
    
End Sub

Sub Octroi_GI()
    
    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")
    
    wbkOpen.Worksheets("TCD").Range("A25:J34").Copy ThisWorkbook.Worksheets("Feuil1").Range("B13")
    
    rowN = wbkOpen.Worksheets("TCD").Range("J34").Row - wbkOpen.Worksheets("TCD").Range("A25").Row

    For rw = 1 To rowN - 1
        wbkOpen.Worksheets("TCD").Range("A78:J78").Offset(rw - 1, 0).Copy
        ThisWorkbook.Worksheets("Feuil1").Range("B13:K13").Offset(2 * rw, 0).Insert Shift:=xlDown
        
        With ThisWorkbook.Worksheets("Feuil1").Range("B13:K13").Offset(2 * rw, 0)
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 5
                .TintAndShade = 0.399914548173467
                .Weight = xlThin
            End With
         
        End With
    Next rw

    wbkOpen.Worksheets("TCD").Range("A78:J78").Offset(rowN - 1, 0).Copy _
    ThisWorkbook.Worksheets("Feuil1").Range("B13:K13").Offset(2 * rowN, 0)

    For rw = 1 To rowN
        With ThisWorkbook.Worksheets("Feuil1")
            .Range("C13").Offset(2 * rw, 0).FormulaR1C1 = .Range("C13").Offset(2 * rw, 0).Value / 1000000
            .Range("D13").Offset(2 * rw, 0).FormulaR1C1 = .Range("D13").Offset(2 * rw, 0).Value / 1000000
            .Range("E13").Offset(2 * rw, 0).FormulaR1C1 = .Range("E13").Offset(2 * rw, 0).Value / 1000000
            .Range("F13").Offset(2 * rw, 0).FormulaR1C1 = .Range("F13").Offset(2 * rw, 0).Value / 1000000
            .Range("G13").Offset(2 * rw, 0).FormulaR1C1 = .Range("G13").Offset(2 * rw, 0).Value / 1000000
            .Range("H13").Offset(2 * rw, 0).FormulaR1C1 = .Range("H13").Offset(2 * rw, 0).Value / 1000000
            .Range("I13").Offset(2 * rw, 0).FormulaR1C1 = .Range("I13").Offset(2 * rw, 0).Value / 1000000
            .Range("J13").Offset(2 * rw, 0).FormulaR1C1 = .Range("J13").Offset(2 * rw, 0).Value / 1000000
            .Range("K13").Offset(2 * rw, 0).FormulaR1C1 = .Range("K13").Offset(2 * rw, 0).Value / 1000000
   
            .Range("C13:M13").Offset(2 * rw, 0).NumberFormat = "0.00"
            .Range("B13").Offset(2 * rw, 0).FormulaR1C1 = "Moyenne des GI octroyées"
        End With
    Next rw
    
    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B13").FormulaR1C1 = "Octroi GI (en M€)"
        .Range("K13").FormulaR1C1 = "2016 act."
    End With

    wbkOpen.Close False
    
End Sub

Sub Octroi_GI_en_nombre()

    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")
    
    wbkOpen.Worksheets("TCD").Range("A129:K139").Copy ThisWorkbook.Worksheets("Feuil1").Range("B43")
    
    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B43").FormulaR1C1 = "Octroi GI (en nombre)"
        .Range("B53").FormulaR1C1 = "Total"
        .Range("K43").FormulaR1C1 = "2016 act."
        .Range("L43").FormulaR1C1 = "Total"
    End With
    
    wbkOpen.Close False

End Sub

Sub GI_douteux()
    
    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\GI-provisions-TCD.xlsm")
    
    wbkOpen.Worksheets("TCD").Range("A7:D14").Copy ThisWorkbook.Worksheets("Feuil1").Range("B65")
    
    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B65").FormulaR1C1 = "GI_douteux (en M€)"
        .Range("C65").FormulaR1C1 = "montant des garantie douteux"
        .Range("D65").FormulaR1C1 = "encours"
        .Range("E65").FormulaR1C1 = "provision"
    End With
    
    wbkOpen.Close False
    
End Sub

Sub Octroi_GP()
    
    Dim Rng As Range
    
    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\1- ARIZ suiviReporting Global-TCD.xlsm")
    
    wbkOpen.Worksheets("TCD").Range("A7:F9").Copy ThisWorkbook.Worksheets("Feuil1").Range("B84")
    rowN = wbkOpen.Worksheets("TCD").Range("F9").Row - wbkOpen.Worksheets("TCD").Range("A7").Row

    For rw = 1 To rowN - 1
        wbkOpen.Worksheets("TCD").Range("A44:F44").Offset(rw - 1, 0).Copy
        ThisWorkbook.Worksheets("Feuil1").Range("B84:G84").Offset(2 * rw, 0).Insert Shift:=xlDown
        
        With ThisWorkbook.Worksheets("Feuil1").Range("B84:G84").Offset(2 * rw, 0)
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 5
                .TintAndShade = 0.399914548173467
                .Weight = xlThin
            End With
         
        End With
    Next rw

    wbkOpen.Worksheets("TCD").Range("A44:F44").Offset(rowN - 1, 0).Copy _
    ThisWorkbook.Worksheets("Feuil1").Range("B84:G84").Offset(2 * rowN, 0)

    For rw = 1 To rowN

        With ThisWorkbook.Worksheets("Feuil1")
            .Range("C84:G84").Offset(2 * rw, 0).NumberFormat = "0.00%"
            .Range("B84").Offset(2 * rw, 0).FormulaR1C1 = "Taux d'utilisation max"
        End With
        
    Next rw

    With ThisWorkbook.Worksheets("Feuil1")
         .Range("B84").FormulaR1C1 = "Octroi GP (en M€)"
         .Range("G84").FormulaR1C1 = "Total"
    End With
        
    wbkOpen.Close False
    
    For Each Rng In Range("B84:G88")
    
        If IsError(Rng.Value) Then
            Rng.Value = 0#
        End If
    
    Next Rng
    
End Sub

Sub MEJ_Nombre_GI()

    Dim wbkOpen2 As Workbook

    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\MEJ-TCD.xlsm")
    Set wbkOpen2 = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")
    
    wbkOpen.Worksheets("TCD").Range("AQ12:AY13").Copy ThisWorkbook.Worksheets("Feuil1").Range("B99")
    wbkOpen2.Worksheets("TCD").Range("A155:H155").Copy ThisWorkbook.Worksheets("Feuil1").Range("B99").Offset(3, 0)
    
    colN = wbkOpen.Worksheets("TCD").Range("AY13").Column - wbkOpen.Worksheets("TCD").Range("AQ12").Column

    wbkOpen.Close False
    wbkOpen2.Close False
    
    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B99").Offset(3, colN).FormulaR1C1 = "=SUM(RC[-7]:RC[-1])"
        
        For cl = 1 To colN
            .Range("B99").Offset(2, cl).FormulaR1C1 = .Range("B99").Offset(1, cl).Value / .Range("B99").Offset(3, cl).Value
        Next cl
        
        .Range("B99").FormulaR1C1 = "MEJ (en nombre)GI"
        .Range("B100").FormulaR1C1 = "nb. de demande"
        .Range("B101").FormulaR1C1 = "Taux de sinistralité en nombre"
        .Range("B99").Offset(0, colN).FormulaR1C1 = "Total"
        .Range("B99").Offset(3, 0).Delete Shift:=xlUp
    
        For cl = 1 To colN
            .Range("B99").Offset(2, cl).NumberFormat = "0.00%"
            .Range("B99").Offset(3, cl).Delete Shift:=xlUp
        Next cl
    End With
    
End Sub

Sub MEJ_GI()

    Dim wbkOpen2 As Workbook

    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\MEJ-TCD.xlsm")
    Set wbkOpen2 = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")

    wbkOpen.Worksheets("TCD").Range("A13:I14").Copy ThisWorkbook.Worksheets("Feuil1").Range("B116")
    wbkOpen2.Worksheets("TCD").Range("B7:H7").Copy ThisWorkbook.Worksheets("Feuil1").Range("B116").Offset(3, 1)
    
    colN = wbkOpen.Worksheets("TCD").Range("I14").Column - wbkOpen.Worksheets("TCD").Range("A13").Column

    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B116").Offset(3, colN).FormulaR1C1 = "=SUM(RC[-7]:RC[-1])"
        
        For cl = 1 To colN
            .Range("B116").Offset(2, cl).FormulaR1C1 = .Range("B116").Offset(1, cl).Value / .Range("B116").Offset(3, cl).Value
        Next cl

        For cl = 1 To colN
            .Range("B116").Offset(1, cl).Font.Bold = False
            .Range("B116").Offset(2, cl).NumberFormat = "0.00%"
            .Range("B116").Offset(3, 1).Delete Shift:=xlToLeft
        
            With .Range("B116").Offset(1, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
    End With
    
    wbkOpen.Worksheets("TCD").Range("A41:I41").Copy ThisWorkbook.Worksheets("Feuil1").Range("B119")
    wbkOpen2.Worksheets("TCD").Range("B7:H7").Copy ThisWorkbook.Worksheets("Feuil1").Range("B119").Offset(2, 1)
    
    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B119").Offset(2, colN).FormulaR1C1 = "=SUM(RC[-7]:RC[-1])"
        
        For cl = 1 To colN
            .Range("B119").Offset(1, cl).FormulaR1C1 = .Range("B119").Offset(0, cl).Value / .Range("B119").Offset(2, cl).Value
        Next cl

        For cl = 1 To colN
            .Range("B119").Offset(0, cl).Font.Bold = False
            .Range("B119").Offset(1, cl).NumberFormat = "0.00%"
            .Range("B119").Offset(2, 1).Delete Shift:=xlToLeft
        
            With .Range("B119").Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
    End With
    
    wbkOpen.Worksheets("TCD").Range("A23:I23").Copy ThisWorkbook.Worksheets("Feuil1").Range("B121")
    wbkOpen2.Worksheets("TCD").Range("B7:H7").Copy ThisWorkbook.Worksheets("Feuil1").Range("B121").Offset(2, 1)
    
    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B121").Offset(2, colN).FormulaR1C1 = "=SUM(RC[-7]:RC[-1])"
        
        For cl = 1 To colN
            .Range("B121").Offset(1, cl).FormulaR1C1 = .Range("B121").Offset(0, cl).Value / .Range("B121").Offset(2, cl).Value
        Next cl

        For cl = 1 To colN
            .Range("B121").Offset(0, cl).Font.Bold = False
            .Range("B121").Offset(1, cl).NumberFormat = "0.00%"
            .Range("B121").Offset(2, 1).Delete Shift:=xlToLeft
        
            With .Range("B121").Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
    End With
    
    wbkOpen.Worksheets("TCD").Range("A32:I32").Copy ThisWorkbook.Worksheets("Feuil1").Range("B123")
    wbkOpen2.Worksheets("TCD").Range("B7:H7").Copy ThisWorkbook.Worksheets("Feuil1").Range("B123").Offset(2, 1)
    
    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B123").Offset(2, colN).FormulaR1C1 = "=SUM(RC[-7]:RC[-1])"
        
        For cl = 1 To colN
            .Range("B123").Offset(1, cl).FormulaR1C1 = .Range("B123").Offset(0, cl).Value / .Range("B123").Offset(2, cl).Value
        Next cl

        For cl = 1 To colN
            .Range("B123").Offset(0, cl).Font.Bold = False
            .Range("B123").Offset(1, cl).NumberFormat = "0.00%"
            .Range("B123").Offset(2, 1).Delete Shift:=xlToLeft
        
            With .Range("B123").Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
    End With

    wbkOpen.Close False
    wbkOpen2.Close False

    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B116").FormulaR1C1 = "MEJ (en M€) GI"
        .Range("B116").Offset(1, 0).FormulaR1C1 = "montant d'engagement garanti"
        .Range("B116").Offset(2, 0).FormulaR1C1 = "Taux de sinistralité 1"
        .Range("B116").Offset(3, 0).FormulaR1C1 = "perte provisoire calculée par la banque"
        .Range("B116").Offset(4, 0).FormulaR1C1 = "Taux de sinistralité 2"
        .Range("B116").Offset(5, 0).FormulaR1C1 = "montant d'indemnisation max"
        .Range("B116").Offset(6, 0).FormulaR1C1 = "Taux de sinistralité 3"
        .Range("B116").Offset(7, 0).FormulaR1C1 = "montant d'indemnisation réel"
        .Range("B116").Offset(8, 0).FormulaR1C1 = "Taux de sinistralité 4"
        .Range("B116").Offset(0, colN).FormulaR1C1 = "Avant 2016"
        
        For n = 1 To 4
            .Range("B116").Offset(2 * n - 1, 0).Font.Bold = False
            
            With .Range("B116").Offset(2 * n - 1, 0).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With  
        Next n
    End With

End Sub

Sub MEJ_GP()

    Dim wbkOpen2 As Workbook

    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\MEJ-TCD.xlsm")
    Set wbkOpen2 = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\1- ARIZ suiviReporting Global-TCD.xlsm")

    wbkOpen.Worksheets("TCD").Range("R13:V14").Copy ThisWorkbook.Worksheets("Feuil1").Range("B129")
    wbkOpen2.Worksheets("TCD").Range("B70:F70").Copy ThisWorkbook.Worksheets("Feuil1").Range("B129").Offset(9, 2)
    
    colN = wbkOpen2.Worksheets("TCD").Range("F70").Column - wbkOpen.Worksheets("TCD").Range("B70").Column + 2

    With ThisWorkbook.Worksheets("Feuil1")

        For cl = 1 To colN
            
            If .Range("B129").Offset(9, cl).Value = 0 Then
                .Range("B129").Offset(2, cl).Value = 0
            Else
                .Range("B129").Offset(2, cl).FormulaR1C1 = .Range("B129").Offset(1, cl).Value / .Range("B129").Offset(9, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Range("B129").Offset(1, cl).Font.Bold = False
            .Range("B129").Offset(2, cl).NumberFormat = "0.00%"
        
            With .Range("B129").Offset(1, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("R41:V41").Copy ThisWorkbook.Worksheets("Feuil1").Range("B132")
    
    With ThisWorkbook.Worksheets("Feuil1")

        For cl = 1 To colN
            
            If .Range("B132").Offset(6, cl).Value = 0 Then
                .Range("B132").Offset(1, cl).Value = 0
            Else
                .Range("B132").Offset(1, cl).FormulaR1C1 = .Range("B132").Offset(0, cl).Value / .Range("B132").Offset(6, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Range("B132").Offset(0, cl).Font.Bold = False
            .Range("B132").Offset(1, cl).NumberFormat = "0.00%"
        
            With .Range("B132").Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("R23:V23").Copy ThisWorkbook.Worksheets("Feuil1").Range("B134")
    
    With ThisWorkbook.Worksheets("Feuil1")

        For cl = 1 To colN
            
            If .Range("B134").Offset(4, cl).Value = 0 Then
                .Range("B134").Offset(1, cl).Value = 0
            Else
                .Range("B134").Offset(1, cl).FormulaR1C1 = .Range("B134").Offset(0, cl).Value / .Range("B134").Offset(4, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Range("B134").Offset(0, cl).Font.Bold = False
            .Range("B134").Offset(1, cl).NumberFormat = "0.00%"
        
            With .Range("B134").Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("R32:V32").Copy ThisWorkbook.Worksheets("Feuil1").Range("B136")
    
    With ThisWorkbook.Worksheets("Feuil1")

        For cl = 1 To colN
            
            If .Range("B136").Offset(2, cl).Value = 0 Then
                .Range("B136").Offset(1, cl).Value = 0
            Else
                .Range("B136").Offset(1, cl).FormulaR1C1 = .Range("B136").Offset(0, cl).Value / .Range("B136").Offset(2, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Range("B136").Offset(0, cl).Font.Bold = False
            .Range("B136").Offset(1, cl).NumberFormat = "0.00%"
            .Range("B136").Offset(2, 1).Delete Shift:=xlToLeft
        
            With .Range("B136").Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With

    wbkOpen.Close False
    wbkOpen2.Close False

    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B129").FormulaR1C1 = "MEJ (en M€) GP"
        .Range("B129").Offset(1, 0).FormulaR1C1 = "montant d'engagement garanti"
        .Range("B129").Offset(2, 0).FormulaR1C1 = "Taux de sinistralité 1"
        .Range("B129").Offset(3, 0).FormulaR1C1 = "perte provisoire calculée par la banque"
        .Range("B129").Offset(4, 0).FormulaR1C1 = "Taux de sinistralité 2"
        .Range("B129").Offset(5, 0).FormulaR1C1 = "montant d'indemnisation max"
        .Range("B129").Offset(6, 0).FormulaR1C1 = "Taux de sinistralité 3"
        .Range("B129").Offset(7, 0).FormulaR1C1 = "montant d'indemnisation réel"
        .Range("B129").Offset(8, 0).FormulaR1C1 = "Taux de sinistralité 4"
        
        .Range("B129").Offset(0, 4).FormulaR1C1 = "2013"
        .Range("B129").Offset(0, 5).FormulaR1C1 = "2014"
        .Range("B129").Offset(0, colN).FormulaR1C1 = "Total"

        .Range("G129:H129").Font.Bold = True
        
        With .Range("G129:H129").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
        With .Range("G129:H129,G131:H131,G133:H133,G135:H135,G137:H137")
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 5
                .TintAndShade = 0.399945066682943
                .Weight = xlThin
            End With
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        End With

    End With
    
    With ThisWorkbook.Worksheets("Feuil1")
        
        For n = 1 To 4
            .Range("B129").Offset(2 * n - 1, 0).Font.Bold = False
            
            With .Range("B129").Offset(2 * n - 1, 0).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next n
        
    End With

End Sub

Sub MEJ_ALIOS()

    Dim wbkOpen2 As Workbook

    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\MEJ-TCD.xlsm")
    Set wbkOpen2 = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")

    wbkOpen.Worksheets("TCD").Range("AE13:AG14").Copy ThisWorkbook.Worksheets("Feuil1").Range("B142")
    wbkOpen2.Worksheets("TCD").Range("F165").Copy ThisWorkbook.Worksheets("Feuil1").Range("B142").Offset(9, 1)
    wbkOpen2.Worksheets("TCD").Range("H165").Copy ThisWorkbook.Worksheets("Feuil1").Range("B142").Offset(9, 2)
    
    colN = wbkOpen.Worksheets("TCD").Range("AG14").Column - wbkOpen.Worksheets("TCD").Range("AE13").Column

    With ThisWorkbook.Worksheets("Feuil1")

        For cl = 1 To colN
            
            If .Range("B142").Offset(9, cl).Value = 0 Then
                .Range("B142").Offset(2, cl).Value = 0
            Else
                .Range("B142").Offset(2, cl).FormulaR1C1 = .Range("B142").Offset(1, cl).Value / .Range("B142").Offset(9, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Range("B142").Offset(1, cl).Font.Bold = False
            .Range("B142").Offset(2, cl).NumberFormat = "0.00%"
        
            With .Range("B142").Offset(1, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("AE41:AG41").Copy ThisWorkbook.Worksheets("Feuil1").Range("B145")
    
    With ThisWorkbook.Worksheets("Feuil1")

        For cl = 1 To colN
            
            If .Range("B145").Offset(6, cl).Value = 0 Then
                .Range("B145").Offset(1, cl).Value = 0
            Else
                .Range("B145").Offset(1, cl).FormulaR1C1 = .Range("B145").Offset(0, cl).Value / .Range("B145").Offset(6, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Range("B145").Offset(0, cl).Font.Bold = False
            .Range("B145").Offset(1, cl).NumberFormat = "0.00%"
        
            With .Range("B145").Offset(0, cl).Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("AE23:AG23").Copy ThisWorkbook.Worksheets("Feuil1").Range("B147")
    
    With ThisWorkbook.Worksheets("Feuil1")

        For cl = 1 To colN
            
            If .Range("B147").Offset(4, cl).Value = 0 Then
                .Range("B147").Offset(1, cl).Value = 0
            Else
                .Range("B147").Offset(1, cl).FormulaR1C1 = .Range("B147").Offset(0, cl).Value / .Range("B147").Offset(4, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Range("B147").Offset(0, cl).Font.Bold = False
            .Range("B147").Offset(1, cl).NumberFormat = "0.00%"
        
            With .Range("B147").Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("AE32:AG32").Copy ThisWorkbook.Worksheets("Feuil1").Range("B149")
    
    With ThisWorkbook.Worksheets("Feuil1")

        For cl = 1 To colN
            
            If .Range("B149").Offset(2, cl).Value = 0 Then
                .Range("B149").Offset(1, cl).Value = 0
            Else
                .Range("B149").Offset(1, cl).FormulaR1C1 = .Range("B149").Offset(0, cl).Value / .Range("B149").Offset(2, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Range("B149").Offset(0, cl).Font.Bold = False
            .Range("B149").Offset(1, cl).NumberFormat = "0.00%"
            .Range("B149").Offset(2, 1).Delete Shift:=xlToLeft
        
            With .Range("B149").Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With

    wbkOpen.Close False
    wbkOpen2.Close False
    
    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B142").FormulaR1C1 = "MEJ (en M€) ALIOS"
        .Range("B142").Offset(1, 0).FormulaR1C1 = "montant d'engagement garanti"
        .Range("B142").Offset(2, 0).FormulaR1C1 = "Taux de sinistralité 1"
        .Range("B142").Offset(3, 0).FormulaR1C1 = "perte provisoire calculée par la banque"
        .Range("B142").Offset(4, 0).FormulaR1C1 = "Taux de sinistralité 2"
        .Range("B142").Offset(5, 0).FormulaR1C1 = "montant d'indemnisation max"
        .Range("B142").Offset(6, 0).FormulaR1C1 = "Taux de sinistralité 3"
        .Range("B142").Offset(7, 0).FormulaR1C1 = "montant d'indemnisation réel"
        .Range("B142").Offset(8, 0).FormulaR1C1 = "Taux de sinistralité 4"
        
        For n = 1 To 4
            .Range("B142").Offset(2 * n - 1, 0).Font.Bold = False
            
            With .Range("B142").Offset(2 * n - 1, 0).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next n
        
    End With
    
End Sub

Sub MEJ_BICIS()

    Dim wbkOpen2 As Workbook

    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\MEJ-TCD.xlsm")
    Set wbkOpen2 = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")

    wbkOpen.Worksheets("TCD").Range("AE52:AH53").Copy ThisWorkbook.Worksheets("Feuil1").Range("B155")
    wbkOpen2.Worksheets("TCD").Range("B166").Copy ThisWorkbook.Worksheets("Feuil1").Range("B155").Offset(9, 1)
    wbkOpen2.Worksheets("TCD").Range("D166").Copy ThisWorkbook.Worksheets("Feuil1").Range("B155").Offset(9, 2)
    wbkOpen2.Worksheets("TCD").Range("F166").Copy ThisWorkbook.Worksheets("Feuil1").Range("B155").Offset(9, 3)
    
    colN = wbkOpen.Worksheets("TCD").Range("AH53").Column - wbkOpen.Worksheets("TCD").Range("AE52").Column

    With ThisWorkbook.Worksheets("Feuil1")

        For cl = 1 To colN
            
            If .Range("B155").Offset(9, cl).Value = 0 Then
                .Range("B155").Offset(2, cl).Value = 0
            Else
                .Range("B155").Offset(2, cl).FormulaR1C1 = .Range("B155").Offset(1, cl).Value / .Range("B155").Offset(9, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Range("B155").Offset(1, cl).Font.Bold = False
            .Range("B155").Offset(2, cl).NumberFormat = "0.00%"
        
            With .Range("B155").Offset(1, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("AE80:AH80").Copy ThisWorkbook.Worksheets("Feuil1").Range("B158")
    
    With ThisWorkbook.Worksheets("Feuil1")

        For cl = 1 To colN
            
            If .Range("B158").Offset(6, cl).Value = 0 Then
                .Range("B158").Offset(1, cl).Value = 0
            Else
                .Range("B158").Offset(1, cl).FormulaR1C1 = .Range("B158").Offset(0, cl).Value / .Range("B158").Offset(6, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Range("B158").Offset(0, cl).Font.Bold = False
            .Range("B158").Offset(1, cl).NumberFormat = "0.00%"
        
            With .Range("B158").Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("AE62:AH62").Copy ThisWorkbook.Worksheets("Feuil1").Range("B160")
    
    With ThisWorkbook.Worksheets("Feuil1")

        For cl = 1 To colN
            
            If .Range("B160").Offset(4, cl).Value = 0 Then
                .Range("B160").Offset(1, cl).Value = 0
            Else
                .Range("B160").Offset(1, cl).FormulaR1C1 = .Range("B160").Offset(0, cl).Value / .Range("B160").Offset(4, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Range("B160").Offset(0, cl).Font.Bold = False
            .Range("B160").Offset(1, cl).NumberFormat = "0.00%"
        
            With .Range("B160").Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("AE71:AH71").Copy ThisWorkbook.Worksheets("Feuil1").Range("B162")
    
    With ThisWorkbook.Worksheets("Feuil1")

        For cl = 1 To colN
            
            If .Range("B162").Offset(2, cl).Value = 0 Then
                .Range("B162").Offset(1, cl).Value = 0
            Else
                .Range("B162").Offset(1, cl).FormulaR1C1 = .Range("B162").Offset(0, cl).Value / .Range("B162").Offset(2, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Range("B162").Offset(0, cl).Font.Bold = False
            .Range("B162").Offset(1, cl).NumberFormat = "0.00%"
            .Range("B162").Offset(2, 1).Delete Shift:=xlToLeft
        
            With .Range("B162").Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With

    wbkOpen.Close False
    wbkOpen2.Close False
    
    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B155").FormulaR1C1 = "MEJ (en M€) BICIS"
        .Range("B155").Offset(1, 0).FormulaR1C1 = "montant d'engagement garanti"
        .Range("B155").Offset(2, 0).FormulaR1C1 = "Taux de sinistralité 1"
        .Range("B155").Offset(3, 0).FormulaR1C1 = "perte provisoire calculée par la banque"
        .Range("B155").Offset(4, 0).FormulaR1C1 = "Taux de sinistralité 2"
        .Range("B155").Offset(5, 0).FormulaR1C1 = "montant d'indemnisation max"
        .Range("B155").Offset(6, 0).FormulaR1C1 = "Taux de sinistralité 3"
        .Range("B155").Offset(7, 0).FormulaR1C1 = "montant d'indemnisation réel"
        .Range("B155").Offset(8, 0).FormulaR1C1 = "Taux de sinistralité 4"
        
        For n = 1 To 4
            .Range("B155").Offset(2 * n - 1, 0).Font.Bold = False
            
            With .Range("B155").Offset(2 * n - 1, 0).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next n
        
    End With
    
End Sub

Sub MEJ_BOASénégal()

    Dim wbkOpen2 As Workbook

    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\MEJ-TCD.xlsm")
    Set wbkOpen2 = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")

    wbkOpen.Worksheets("TCD").Range("AE91:AG92").Copy ThisWorkbook.Worksheets("Feuil1").Range("B168")
    wbkOpen2.Worksheets("TCD").Range("C168:D168").Copy ThisWorkbook.Worksheets("Feuil1").Range("B168").Offset(9, 1)
    
    colN = wbkOpen.Worksheets("TCD").Range("AG92").Column - wbkOpen.Worksheets("TCD").Range("AE91").Column

    With ThisWorkbook.Worksheets("Feuil1").Range("B168")

        For cl = 1 To colN
            
            If .Offset(9, cl).Value = 0 Then
                .Offset(2, cl).Value = 0
            Else
                .Offset(2, cl).FormulaR1C1 = .Offset(1, cl).Value / .Offset(9, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Offset(1, cl).Font.Bold = False
            .Offset(2, cl).NumberFormat = "0.00%"
        
            With .Offset(1, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("AE119:AG119").Copy ThisWorkbook.Worksheets("Feuil1").Range("B171")
    
    With ThisWorkbook.Worksheets("Feuil1").Range("B171")

        For cl = 1 To colN
            
            If .Offset(6, cl).Value = 0 Then
                .Offset(1, cl).Value = 0
            Else
                .Offset(1, cl).FormulaR1C1 = .Offset(0, cl).Value / .Offset(6, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Offset(0, cl).Font.Bold = False
            .Offset(1, cl).NumberFormat = "0.00%"
        
            With .Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("AE101:AG101").Copy ThisWorkbook.Worksheets("Feuil1").Range("B173")
    
    With ThisWorkbook.Worksheets("Feuil1").Range("B173")

        For cl = 1 To colN
            
            If .Offset(4, cl).Value = 0 Then
                .Offset(1, cl).Value = 0
            Else
                .Offset(1, cl).FormulaR1C1 = .Offset(0, cl).Value / .Offset(4, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Offset(0, cl).Font.Bold = False
            .Offset(1, cl).NumberFormat = "0.00%"
        
            With .Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("AE110:AG110").Copy ThisWorkbook.Worksheets("Feuil1").Range("B175")
    
    With ThisWorkbook.Worksheets("Feuil1").Range("B175")

        For cl = 1 To colN
            
            If .Offset(2, cl).Value = 0 Then
                .Offset(1, cl).Value = 0
            Else
                .Offset(1, cl).FormulaR1C1 = .Offset(0, cl).Value / .Offset(2, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Offset(0, cl).Font.Bold = False
            .Offset(1, cl).NumberFormat = "0.00%"
            .Offset(2, 1).Delete Shift:=xlToLeft
        
            With .Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With

    wbkOpen.Close False
    wbkOpen2.Close False
    
    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B168").FormulaR1C1 = "MEJ (en M€) BOA Sénégal"
        .Range("B168").Offset(1, 0).FormulaR1C1 = "montant d'engagement garanti"
        .Range("B168").Offset(2, 0).FormulaR1C1 = "Taux de sinistralité 1"
        .Range("B168").Offset(3, 0).FormulaR1C1 = "perte provisoire calculée par la banque"
        .Range("B168").Offset(4, 0).FormulaR1C1 = "Taux de sinistralité 2"
        .Range("B168").Offset(5, 0).FormulaR1C1 = "montant d'indemnisation max"
        .Range("B168").Offset(6, 0).FormulaR1C1 = "Taux de sinistralité 3"
        .Range("B168").Offset(7, 0).FormulaR1C1 = "montant d'indemnisation réel"
        .Range("B168").Offset(8, 0).FormulaR1C1 = "Taux de sinistralité 4"
        
        For n = 1 To 4
            .Range("B168").Offset(2 * n - 1, 0).Font.Bold = False
            
            With .Range("B168").Offset(2 * n - 1, 0).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next n
        
    End With
    
End Sub

Sub MEJ_CBAO()

    Dim wbkOpen2 As Workbook

    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\MEJ-TCD.xlsm")
    Set wbkOpen2 = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")

    wbkOpen.Worksheets("TCD").Range("AE130:AF131").Copy ThisWorkbook.Worksheets("Feuil1").Range("B181")
    wbkOpen2.Worksheets("TCD").Range("B169").Copy ThisWorkbook.Worksheets("Feuil1").Range("B181").Offset(9, 1)
    
    colN = wbkOpen.Worksheets("TCD").Range("AF131").Column - wbkOpen.Worksheets("TCD").Range("AE130").Column

    With ThisWorkbook.Worksheets("Feuil1").Range("B181")

        For cl = 1 To colN
            
            If .Offset(9, cl).Value = 0 Then
                .Offset(2, cl).Value = 0
            Else
                .Offset(2, cl).FormulaR1C1 = .Offset(1, cl).Value / .Offset(9, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Offset(1, cl).Font.Bold = False
            .Offset(2, cl).NumberFormat = "0.00%"
        
            With .Offset(1, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("AE158:AF158").Copy ThisWorkbook.Worksheets("Feuil1").Range("B184")
    
    With ThisWorkbook.Worksheets("Feuil1").Range("B184")

        For cl = 1 To colN
            
            If .Offset(6, cl).Value = 0 Then
                .Offset(1, cl).Value = 0
            Else
                .Offset(1, cl).FormulaR1C1 = .Offset(0, cl).Value / .Offset(6, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Offset(0, cl).Font.Bold = False
            .Offset(1, cl).NumberFormat = "0.00%"
        
            With .Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("AE140:AF140").Copy ThisWorkbook.Worksheets("Feuil1").Range("B186")
    
    With ThisWorkbook.Worksheets("Feuil1").Range("B186")

        For cl = 1 To colN
            
            If .Offset(4, cl).Value = 0 Then
                .Offset(1, cl).Value = 0
            Else
                .Offset(1, cl).FormulaR1C1 = .Offset(0, cl).Value / .Offset(4, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Offset(0, cl).Font.Bold = False
            .Offset(1, cl).NumberFormat = "0.00%"
        
            With .Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("AE149:AF149").Copy ThisWorkbook.Worksheets("Feuil1").Range("B188")
    
    With ThisWorkbook.Worksheets("Feuil1").Range("B188")

        For cl = 1 To colN
            
            If .Offset(2, cl).Value = 0 Then
                .Offset(1, cl).Value = 0
            Else
                .Offset(1, cl).FormulaR1C1 = .Offset(0, cl).Value / .Offset(2, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Offset(0, cl).Font.Bold = False
            .Offset(1, cl).NumberFormat = "0.00%"
            .Offset(2, 1).Delete Shift:=xlToLeft
        
            With .Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With

    wbkOpen.Close False
    wbkOpen2.Close False
    
    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B181").FormulaR1C1 = "MEJ (en M€) CBAO"
        .Range("B181").Offset(1, 0).FormulaR1C1 = "montant d'engagement garanti"
        .Range("B181").Offset(2, 0).FormulaR1C1 = "Taux de sinistralité 1"
        .Range("B181").Offset(3, 0).FormulaR1C1 = "perte provisoire calculée par la banque"
        .Range("B181").Offset(4, 0).FormulaR1C1 = "Taux de sinistralité 2"
        .Range("B181").Offset(5, 0).FormulaR1C1 = "montant d'indemnisation max"
        .Range("B181").Offset(6, 0).FormulaR1C1 = "Taux de sinistralité 3"
        .Range("B181").Offset(7, 0).FormulaR1C1 = "montant d'indemnisation réel"
        .Range("B181").Offset(8, 0).FormulaR1C1 = "Taux de sinistralité 4"
        
        For n = 1 To 4
            .Range("B181").Offset(2 * n - 1, 0).Font.Bold = False
            
            With .Range("B181").Offset(2 * n - 1, 0).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next n
        
    End With
    
End Sub

Sub MEJ_CREDIT_DU_SENEGAL()

    Dim wbkOpen2 As Workbook

    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\MEJ-TCD.xlsm")
    Set wbkOpen2 = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")

    wbkOpen.Worksheets("TCD").Range("AE169:AH170").Copy ThisWorkbook.Worksheets("Feuil1").Range("B194")
    wbkOpen2.Worksheets("TCD").Range("B171:C171").Copy ThisWorkbook.Worksheets("Feuil1").Range("B194").Offset(9, 1)
    wbkOpen2.Worksheets("TCD").Range("F171").Copy ThisWorkbook.Worksheets("Feuil1").Range("B194").Offset(9, 3)
    
    colN = wbkOpen.Worksheets("TCD").Range("AH170").Column - wbkOpen.Worksheets("TCD").Range("AE169").Column

    With ThisWorkbook.Worksheets("Feuil1").Range("B194")

        For cl = 1 To colN
            
            If .Offset(9, cl).Value = 0 Then
                .Offset(2, cl).Value = 0
            Else
                .Offset(2, cl).FormulaR1C1 = .Offset(1, cl).Value / .Offset(9, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Offset(1, cl).Font.Bold = False
            .Offset(2, cl).NumberFormat = "0.00%"
        
            With .Offset(1, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("AE197:AH197").Copy ThisWorkbook.Worksheets("Feuil1").Range("B197")
    
    With ThisWorkbook.Worksheets("Feuil1").Range("B197")

        For cl = 1 To colN
            
            If .Offset(6, cl).Value = 0 Then
                .Offset(1, cl).Value = 0
            Else
                .Offset(1, cl).FormulaR1C1 = .Offset(0, cl).Value / .Offset(6, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Offset(0, cl).Font.Bold = False
            .Offset(1, cl).NumberFormat = "0.00%"
        
            With .Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("AE179:AH179").Copy ThisWorkbook.Worksheets("Feuil1").Range("B199")
    
    With ThisWorkbook.Worksheets("Feuil1").Range("B199")

        For cl = 1 To colN
            
            If .Offset(4, cl).Value = 0 Then
                .Offset(1, cl).Value = 0
            Else
                .Offset(1, cl).FormulaR1C1 = .Offset(0, cl).Value / .Offset(4, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Offset(0, cl).Font.Bold = False
            .Offset(1, cl).NumberFormat = "0.00%"
        
            With .Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("AE188:AH188").Copy ThisWorkbook.Worksheets("Feuil1").Range("B201")
    
    With ThisWorkbook.Worksheets("Feuil1").Range("B201")

        For cl = 1 To colN
            
            If .Offset(2, cl).Value = 0 Then
                .Offset(1, cl).Value = 0
            Else
                .Offset(1, cl).FormulaR1C1 = .Offset(0, cl).Value / .Offset(2, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Offset(0, cl).Font.Bold = False
            .Offset(1, cl).NumberFormat = "0.00%"
            .Offset(2, 1).Delete Shift:=xlToLeft
        
            With .Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With

    wbkOpen.Close False
    wbkOpen2.Close False
    
    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B194").FormulaR1C1 = "MEJ (en M€) CREDIT DU SENEGAL"
        .Range("B194").Offset(1, 0).FormulaR1C1 = "montant d'engagement garanti"
        .Range("B194").Offset(2, 0).FormulaR1C1 = "Taux de sinistralité 1"
        .Range("B194").Offset(3, 0).FormulaR1C1 = "perte provisoire calculée par la banque"
        .Range("B194").Offset(4, 0).FormulaR1C1 = "Taux de sinistralité 2"
        .Range("B194").Offset(5, 0).FormulaR1C1 = "montant d'indemnisation max"
        .Range("B194").Offset(6, 0).FormulaR1C1 = "Taux de sinistralité 3"
        .Range("B194").Offset(7, 0).FormulaR1C1 = "montant d'indemnisation réel"
        .Range("B194").Offset(8, 0).FormulaR1C1 = "Taux de sinistralité 4"
        
        For n = 1 To 4
            .Range("B194").Offset(2 * n - 1, 0).Font.Bold = False
            
            With .Range("B194").Offset(2 * n - 1, 0).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next n
        
    End With
    
End Sub

Sub MEJ_SGBS()

    Dim wbkOpen2 As Workbook

    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\MEJ-TCD.xlsm")
    Set wbkOpen2 = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")

    wbkOpen.Worksheets("TCD").Range("AE208:AK209").Copy ThisWorkbook.Worksheets("Feuil1").Range("B207")
    wbkOpen2.Worksheets("TCD").Range("B174:G174").Copy ThisWorkbook.Worksheets("Feuil1").Range("B207").Offset(9, 1)
    
    colN = wbkOpen.Worksheets("TCD").Range("AK209").Column - wbkOpen.Worksheets("TCD").Range("AE208").Column

    With ThisWorkbook.Worksheets("Feuil1").Range("B207")

        For cl = 1 To colN
            
            If .Offset(9, cl).Value = 0 Then
                .Offset(2, cl).Value = 0
            Else
                .Offset(2, cl).FormulaR1C1 = .Offset(1, cl).Value / .Offset(9, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Offset(1, cl).Font.Bold = False
            .Offset(2, cl).NumberFormat = "0.00%"
        
            With .Offset(1, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("AE236:AK236").Copy ThisWorkbook.Worksheets("Feuil1").Range("B210")
    
    With ThisWorkbook.Worksheets("Feuil1").Range("B210")

        For cl = 1 To colN
            
            If .Offset(6, cl).Value = 0 Then
                .Offset(1, cl).Value = 0
            Else
                .Offset(1, cl).FormulaR1C1 = .Offset(0, cl).Value / .Offset(6, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Offset(0, cl).Font.Bold = False
            .Offset(1, cl).NumberFormat = "0.00%"
        
            With .Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("AE218:AK218").Copy ThisWorkbook.Worksheets("Feuil1").Range("B212")
    
    With ThisWorkbook.Worksheets("Feuil1").Range("B212")

        For cl = 1 To colN
            
            If .Offset(4, cl).Value = 0 Then
                .Offset(1, cl).Value = 0
            Else
                .Offset(1, cl).FormulaR1C1 = .Offset(0, cl).Value / .Offset(4, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Offset(0, cl).Font.Bold = False
            .Offset(1, cl).NumberFormat = "0.00%"
        
            With .Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With
    
    wbkOpen.Worksheets("TCD").Range("AE227:AK227").Copy ThisWorkbook.Worksheets("Feuil1").Range("B214")
    
    With ThisWorkbook.Worksheets("Feuil1").Range("B214")

        For cl = 1 To colN
            
            If .Offset(2, cl).Value = 0 Then
                .Offset(1, cl).Value = 0
            Else
                .Offset(1, cl).FormulaR1C1 = .Offset(0, cl).Value / .Offset(2, cl).Value
            End If
            
        Next cl

        For cl = 1 To colN
            .Offset(0, cl).Font.Bold = False
            .Offset(1, cl).NumberFormat = "0.00%"
            .Offset(2, 1).Delete Shift:=xlToLeft
        
            With .Offset(0, cl).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next cl
        
    End With

    wbkOpen.Close False
    wbkOpen2.Close False
    
    With ThisWorkbook.Worksheets("Feuil1").Range("B207")
        .FormulaR1C1 = "MEJ (en M€) SGBS"
        .Offset(1, 0).FormulaR1C1 = "montant d'engagement garanti"
        .Offset(2, 0).FormulaR1C1 = "Taux de sinistralité 1"
        .Offset(3, 0).FormulaR1C1 = "perte provisoire calculée par la banque"
        .Offset(4, 0).FormulaR1C1 = "Taux de sinistralité 2"
        .Offset(5, 0).FormulaR1C1 = "montant d'indemnisation max"
        .Offset(6, 0).FormulaR1C1 = "Taux de sinistralité 3"
        .Offset(7, 0).FormulaR1C1 = "montant d'indemnisation réel"
        .Offset(8, 0).FormulaR1C1 = "Taux de sinistralité 4"
        
        For n = 1 To 4
            .Offset(2 * n - 1, 0).Font.Bold = False
            
            With .Offset(2 * n - 1, 0).Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next n
        
    End With
    
End Sub

Sub Conformité()

    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\Conformité-TCD.xlsx")

    wbkOpen.Worksheets("TdB___Conformité").Range("A1:D6").Copy ThisWorkbook.Worksheets("Feuil1").Range("B220")
    
    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B220").FormulaR1C1 = "Conformité"
        
        With .Range("B220:E220").Interior
             .Pattern = xlSolid
             .PatternColorIndex = xlAutomatic
             .ThemeColor = xlThemeColorAccent1
             .TintAndShade = 0.799981688894314
             .PatternTintAndShade = 0
        End With

        With .Range("B220:E220")
             .Borders(xlEdgeTop).LineStyle = xlNone
             .Borders(xlEdgeLeft).LineStyle = xlNone
             .Borders(xlEdgeRight).LineStyle = xlNone
             .Borders(xlDiagonalUp).LineStyle = xlNone
             .Borders(xlDiagonalDown).LineStyle = xlNone
             .Borders(xlInsideVertical).LineStyle = xlNone
             .Borders(xlInsideHorizontal).LineStyle = xlNone
             .Font.Bold = True
        End With

        With .Range("B220:E220").Borders(xlEdgeBottom)
             .LineStyle = xlContinuous
             .ThemeColor = 5
             .TintAndShade = 0.399945066682943
             .Weight = xlThin
        End With
        
    End With

    wbkOpen.Close False

End Sub

Sub Eligibilité_financière()

    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\Eligibilité financière-TCD.xlsx")

    wbkOpen.Worksheets("TdB___Eligibilité_financière").Range("A1:D6").Copy ThisWorkbook.Worksheets("Feuil1").Range("B235")
    
    With ThisWorkbook.Worksheets("Feuil1")
        .Range("B235").FormulaR1C1 = "Eligibilité financière"
        
        With .Range("B235:E235").Interior
             .Pattern = xlSolid
             .PatternColorIndex = xlAutomatic
             .ThemeColor = xlThemeColorAccent1
             .TintAndShade = 0.799981688894314
             .PatternTintAndShade = 0
        End With

        With .Range("B235:E235")
             .Borders(xlEdgeTop).LineStyle = xlNone
             .Borders(xlEdgeLeft).LineStyle = xlNone
             .Borders(xlEdgeRight).LineStyle = xlNone
             .Borders(xlDiagonalUp).LineStyle = xlNone
             .Borders(xlDiagonalDown).LineStyle = xlNone
             .Borders(xlInsideVertical).LineStyle = xlNone
             .Borders(xlInsideHorizontal).LineStyle = xlNone
             .Font.Bold = True
        End With

        With .Range("B235:E235").Borders(xlEdgeBottom)
             .LineStyle = xlContinuous
             .ThemeColor = 5
             .TintAndShade = 0.399945066682943
             .Weight = xlThin
        End With
        
    End With

    wbkOpen.Close False

End Sub

Sub MEJ_montant_max_grpBancaire()
    
    Dim wbkOpen2 As Workbook
    
    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\MEJ-TCD.xlsm")
    Set wbkOpen2 = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")
    
    wbkOpen.Worksheets("TCD").Range("BH12:BO17").Copy ThisWorkbook.Worksheets("Feuil1").Range("B250")
    
    rowN = wbkOpen.Worksheets("TCD").Range("BO17").Row - wbkOpen.Worksheets("TCD").Range("BH12").Row
    colN = wbkOpen.Worksheets("TCD").Range("BO17").Column - wbkOpen.Worksheets("TCD").Range("BH12").Column

    With ThisWorkbook.Worksheets("Feuil1").Range("B250")
    
        For rw = 1 To rowN
        
            For cl = 1 To colN
                .Offset(rw, cl).FormulaR1C1 = .Offset(rw, cl).Value / 1000000
                .Offset(rw, cl).NumberFormat = "0.00"
            Next cl
            
        Next rw
         
    End With

    wbkOpen2.Worksheets("TCD").Range("A191:H191").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B250").Offset(2, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B250")
        .Offset(2, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(2, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(3, cl).Value = 0 Then
                .Offset(2, cl).Value = 0
            Else
                .Offset(2, cl).FormulaR1C1 = .Offset(1, cl).Value / .Offset(3, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(2, cl).NumberFormat = "0.00%"
            .Offset(3, cl).Delete Shift:=xlUp
        Next cl

        .Offset(3, 0).Delete Shift:=xlUp
    End With
         
    wbkOpen2.Worksheets("TCD").Range("A192:H192").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B250").Offset(4, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B250")
        .Offset(4, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(4, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(5, cl).Value = 0 Then
                .Offset(4, cl).Value = 0
            Else
                .Offset(4, cl).FormulaR1C1 = .Offset(3, cl).Value / .Offset(5, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(4, cl).NumberFormat = "0.00%"
            .Offset(5, cl).Delete Shift:=xlUp
        Next cl

        .Offset(5, 0).Delete Shift:=xlUp
    End With

    wbkOpen2.Worksheets("TCD").Range("A193:H193").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B250").Offset(6, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B250")
        .Offset(6, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(6, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(5, cl).Value = 0 Then
                .Offset(6, cl).Value = 0
            Else
                .Offset(6, cl).FormulaR1C1 = .Offset(5, cl).Value / .Offset(7, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(6, cl).NumberFormat = "0.00%"
            .Offset(7, cl).Delete Shift:=xlUp
        Next cl

        .Offset(7, 0).Delete Shift:=xlUp
    End With

    wbkOpen2.Worksheets("TCD").Range("A194:H194").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B250").Offset(8, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B250")
        .Offset(8, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(8, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(7, cl).Value = 0 Then
                .Offset(8, cl).Value = 0
            Else
                .Offset(8, cl).FormulaR1C1 = .Offset(7, cl).Value / .Offset(9, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(8, cl).NumberFormat = "0.00%"
            .Offset(9, cl).Delete Shift:=xlUp
        Next cl

        .Offset(9, 0).Delete Shift:=xlUp
    End With

    wbkOpen2.Worksheets("TCD").Range("A195:H195").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B250").Offset(10, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B250")
        .Offset(10, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(10, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(9, cl).Value = 0 Then
                .Offset(10, cl).Value = 0
            Else
                .Offset(10, cl).FormulaR1C1 = .Offset(9, cl).Value / .Offset(11, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(10, cl).NumberFormat = "0.00%"
            .Offset(11, cl).Delete Shift:=xlUp
        Next cl

        .Offset(11, 0).Delete Shift:=xlUp
    End With

    wbkOpen.Close False
    wbkOpen2.Close False

    For rw = 1 To rowN - 1
        
        With ThisWorkbook.Worksheets("Feuil1").Range("B250:I250").Offset(2 * rw, 0)
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 5
                .TintAndShade = 0.399914548173467
                .Weight = xlThin
            End With
         
        End With

    Next rw

    With ThisWorkbook.Worksheets("Feuil1").Range("B250")
        .FormulaR1C1 = "MEJ (en M€) montant max/groupe bancaire"
        
        For rw = 1 To rowN
            .Offset(2 * rw, 0).FormulaR1C1 = "Taux de sinistralité"
        Next rw
    End With

End Sub

Sub MEJ_montant_max_nature()
    
    Dim wbkOpen2 As Workbook
    
    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\MEJ-TCD.xlsm")
    Set wbkOpen2 = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")
    
    wbkOpen.Worksheets("TCD").Range("BH34:BO36").Copy ThisWorkbook.Worksheets("Feuil1").Range("B270")
    
    rowN = wbkOpen.Worksheets("TCD").Range("BO36").Row - wbkOpen.Worksheets("TCD").Range("BH34").Row
    colN = wbkOpen.Worksheets("TCD").Range("BO36").Column - wbkOpen.Worksheets("TCD").Range("BH34").Column

    With ThisWorkbook.Worksheets("Feuil1").Range("B270")
    
        For rw = 1 To rowN
        
            For cl = 1 To colN
                .Offset(rw, cl).FormulaR1C1 = .Offset(rw, cl).Value / 1000000
                .Offset(rw, cl).NumberFormat = "0.00"
            Next cl
            
        Next rw
         
    End With

    wbkOpen2.Worksheets("TCD").Range("A214:H214").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B270").Offset(2, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B270")
        .Offset(2, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(2, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(3, cl).Value = 0 Then
                .Offset(2, cl).Value = 0
            Else
                .Offset(2, cl).FormulaR1C1 = .Offset(1, cl).Value / .Offset(3, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(2, cl).NumberFormat = "0.00%"
            .Offset(3, cl).Delete Shift:=xlUp
        Next cl

        .Offset(3, 0).Delete Shift:=xlUp
    End With
         
    wbkOpen2.Worksheets("TCD").Range("A215:H215").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B270").Offset(4, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B270")
        .Offset(4, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(4, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(5, cl).Value = 0 Then
                .Offset(4, cl).Value = 0
            Else
                .Offset(4, cl).FormulaR1C1 = .Offset(3, cl).Value / .Offset(5, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(4, cl).NumberFormat = "0.00%"
            .Offset(5, cl).Delete Shift:=xlUp
        Next cl

        .Offset(5, 0).Delete Shift:=xlUp
    End With
    
    wbkOpen.Close False
    wbkOpen2.Close False

    For rw = 1 To rowN - 1
        
        With ThisWorkbook.Worksheets("Feuil1").Range("B270:I270").Offset(2 * rw, 0)
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 5
                .TintAndShade = 0.399914548173467
                .Weight = xlThin
            End With
         
        End With

    Next rw

    With ThisWorkbook.Worksheets("Feuil1").Range("B270")
        .FormulaR1C1 = "MEJ (en M€) montant max/nature"
        
        For rw = 1 To rowN
            .Offset(2 * rw, 0).FormulaR1C1 = "Taux de sinistralité"
        Next rw
    End With

End Sub

Sub MEJ_montant_max_secteur()
    
    Dim wbkOpen2 As Workbook
    
    Set wbkOpen = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\MEJ-TCD.xlsm")
    Set wbkOpen2 = Workbooks.Open(ThisWorkbook.Path & "\Tableaux Croisés Dynamiques\BDD Principale-TCD.xlsm")
    
    wbkOpen.Worksheets("TCD").Range("BH53:BO69").Copy ThisWorkbook.Worksheets("Feuil1").Range("B288")
    
    rowN = wbkOpen.Worksheets("TCD").Range("BO69").Row - wbkOpen.Worksheets("TCD").Range("BH53").Row
    colN = wbkOpen.Worksheets("TCD").Range("BO69").Column - wbkOpen.Worksheets("TCD").Range("BH53").Column

    With ThisWorkbook.Worksheets("Feuil1").Range("B288")
    
        For rw = 1 To rowN
        
            For cl = 1 To colN
                .Offset(rw, cl).FormulaR1C1 = .Offset(rw, cl).Value / 1000000
                .Offset(rw, cl).NumberFormat = "0.00"
            Next cl
            
        Next rw
         
    End With

    wbkOpen2.Worksheets("TCD").Range("A234:H234").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B288").Offset(2, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B288")
        .Offset(2, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(2, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(3, cl).Value = 0 Then
                .Offset(2, cl).Value = 0
            Else
                .Offset(2, cl).FormulaR1C1 = .Offset(1, cl).Value / .Offset(3, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(2, cl).NumberFormat = "0.00%"
            .Offset(3, cl).Delete Shift:=xlUp
        Next cl

        .Offset(3, 0).Delete Shift:=xlUp
    End With
         
    wbkOpen2.Worksheets("TCD").Range("A235:H235").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B288").Offset(4, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B288")
        .Offset(4, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(4, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(5, cl).Value = 0 Then
                .Offset(4, cl).Value = 0
            Else
                .Offset(4, cl).FormulaR1C1 = .Offset(3, cl).Value / .Offset(5, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(4, cl).NumberFormat = "0.00%"
            .Offset(5, cl).Delete Shift:=xlUp
        Next cl

        .Offset(5, 0).Delete Shift:=xlUp
    End With

    wbkOpen2.Worksheets("TCD").Range("A241:H241").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B288").Offset(6, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B288")
        .Offset(6, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(6, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(5, cl).Value = 0 Then
                .Offset(6, cl).Value = 0
            Else
                .Offset(6, cl).FormulaR1C1 = .Offset(5, cl).Value / .Offset(7, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(6, cl).NumberFormat = "0.00%"
            .Offset(7, cl).Delete Shift:=xlUp
        Next cl

        .Offset(7, 0).Delete Shift:=xlUp
    End With

    wbkOpen2.Worksheets("TCD").Range("A247:H247").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B288").Offset(8, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B288")
        .Offset(8, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(8, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(7, cl).Value = 0 Then
                .Offset(8, cl).Value = 0
            Else
                .Offset(8, cl).FormulaR1C1 = .Offset(7, cl).Value / .Offset(9, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(8, cl).NumberFormat = "0.00%"
            .Offset(9, cl).Delete Shift:=xlUp
        Next cl

        .Offset(9, 0).Delete Shift:=xlUp
    End With

    wbkOpen2.Worksheets("TCD").Range("A248:H248").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B288").Offset(10, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B288")
        .Offset(10, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(10, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(9, cl).Value = 0 Then
                .Offset(10, cl).Value = 0
            Else
                .Offset(10, cl).FormulaR1C1 = .Offset(9, cl).Value / .Offset(11, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(10, cl).NumberFormat = "0.00%"
            .Offset(11, cl).Delete Shift:=xlUp
        Next cl

        .Offset(11, 0).Delete Shift:=xlUp
    End With

    wbkOpen2.Worksheets("TCD").Range("A249:H249").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B288").Offset(12, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B288")
        .Offset(12, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(12, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(11, cl).Value = 0 Then
                .Offset(12, cl).Value = 0
            Else
                .Offset(12, cl).FormulaR1C1 = .Offset(11, cl).Value / .Offset(13, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(12, cl).NumberFormat = "0.00%"
            .Offset(13, cl).Delete Shift:=xlUp
        Next cl

        .Offset(13, 0).Delete Shift:=xlUp
    End With

    wbkOpen2.Worksheets("TCD").Range("A254:H254").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B288").Offset(14, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B288")
        .Offset(14, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(14, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(13, cl).Value = 0 Then
                .Offset(14, cl).Value = 0
            Else
                .Offset(14, cl).FormulaR1C1 = .Offset(13, cl).Value / .Offset(15, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(14, cl).NumberFormat = "0.00%"
            .Offset(15, cl).Delete Shift:=xlUp
        Next cl

        .Offset(15, 0).Delete Shift:=xlUp
    End With

    wbkOpen2.Worksheets("TCD").Range("A256:H256").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B288").Offset(16, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B288")
        .Offset(16, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(16, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(15, cl).Value = 0 Then
                .Offset(16, cl).Value = 0
            Else
                .Offset(16, cl).FormulaR1C1 = .Offset(15, cl).Value / .Offset(17, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(16, cl).NumberFormat = "0.00%"
            .Offset(17, cl).Delete Shift:=xlUp
        Next cl

        .Offset(17, 0).Delete Shift:=xlUp
    End With

    wbkOpen2.Worksheets("TCD").Range("A263:H263").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B288").Offset(18, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B288")
        .Offset(18, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(18, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(17, cl).Value = 0 Then
                .Offset(18, cl).Value = 0
            Else
                .Offset(18, cl).FormulaR1C1 = .Offset(17, cl).Value / .Offset(19, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(18, cl).NumberFormat = "0.00%"
            .Offset(19, cl).Delete Shift:=xlUp
        Next cl

        .Offset(19, 0).Delete Shift:=xlUp
    End With

    wbkOpen2.Worksheets("TCD").Range("A264:H264").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B288").Offset(20, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B288")
        .Offset(20, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(20, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(19, cl).Value = 0 Then
                .Offset(20, cl).Value = 0
            Else
                .Offset(20, cl).FormulaR1C1 = .Offset(19, cl).Value / .Offset(21, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(20, cl).NumberFormat = "0.00%"
            .Offset(21, cl).Delete Shift:=xlUp
        Next cl

        .Offset(21, 0).Delete Shift:=xlUp
    End With

    wbkOpen2.Worksheets("TCD").Range("A266:H266").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B288").Offset(22, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B288")
        .Offset(22, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(22, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(21, cl).Value = 0 Then
                .Offset(22, cl).Value = 0
            Else
                .Offset(22, cl).FormulaR1C1 = .Offset(21, cl).Value / .Offset(23, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(22, cl).NumberFormat = "0.00%"
            .Offset(23, cl).Delete Shift:=xlUp
        Next cl

        .Offset(23, 0).Delete Shift:=xlUp
    End With

    wbkOpen2.Worksheets("TCD").Range("A268:H268").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B288").Offset(24, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B288")
        .Offset(24, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(24, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(23, cl).Value = 0 Then
                .Offset(24, cl).Value = 0
            Else
                .Offset(24, cl).FormulaR1C1 = .Offset(23, cl).Value / .Offset(25, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(24, cl).NumberFormat = "0.00%"
            .Offset(25, cl).Delete Shift:=xlUp
        Next cl

        .Offset(25, 0).Delete Shift:=xlUp
    End With

    wbkOpen2.Worksheets("TCD").Range("A273:H273").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B288").Offset(26, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B288")
        .Offset(26, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(26, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(25, cl).Value = 0 Then
                .Offset(26, cl).Value = 0
            Else
                .Offset(26, cl).FormulaR1C1 = .Offset(25, cl).Value / .Offset(27, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(26, cl).NumberFormat = "0.00%"
            .Offset(27, cl).Delete Shift:=xlUp
        Next cl

        .Offset(27, 0).Delete Shift:=xlUp
    End With

    wbkOpen2.Worksheets("TCD").Range("A276:H276").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B288").Offset(28, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B288")
        .Offset(28, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(28, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(27, cl).Value = 0 Then
                .Offset(28, cl).Value = 0
            Else
                .Offset(28, cl).FormulaR1C1 = .Offset(27, cl).Value / .Offset(29, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(28, cl).NumberFormat = "0.00%"
            .Offset(29, cl).Delete Shift:=xlUp
        Next cl

        .Offset(29, 0).Delete Shift:=xlUp
    End With

    wbkOpen2.Worksheets("TCD").Range("A282:H282").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B288").Offset(30, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B288")
        .Offset(30, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(30, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(29, cl).Value = 0 Then
                .Offset(30, cl).Value = 0
            Else
                .Offset(30, cl).FormulaR1C1 = .Offset(29, cl).Value / .Offset(31, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(30, cl).NumberFormat = "0.00%"
            .Offset(31, cl).Delete Shift:=xlUp
        Next cl

        .Offset(31, 0).Delete Shift:=xlUp
    End With

    wbkOpen2.Worksheets("TCD").Range("A283:H283").Copy
    ThisWorkbook.Worksheets("Feuil1").Range("B288").Offset(32, 0).Insert Shift:=xlDown

    With ThisWorkbook.Worksheets("Feuil1").Range("B288")
        .Offset(32, 0).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
        For cl = 1 To colN
            .Offset(32, cl).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            If .Offset(31, cl).Value = 0 Then
                .Offset(32, cl).Value = 0
            Else
                .Offset(32, cl).FormulaR1C1 = .Offset(31, cl).Value / .Offset(33, cl).Value
            End If
        Next cl

        For cl = 1 To colN
            .Offset(32, cl).NumberFormat = "0.00%"
            .Offset(33, cl).Delete Shift:=xlUp
        Next cl

        .Offset(33, 0).Delete Shift:=xlUp
    End With

    wbkOpen.Close False
    wbkOpen2.Close False

    For rw = 1 To rowN - 1
        
        With ThisWorkbook.Worksheets("Feuil1").Range("B288:I288").Offset(2 * rw, 0)
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ThemeColor = 5
                .TintAndShade = 0.399914548173467
                .Weight = xlThin
            End With
         
        End With

    Next rw

    With ThisWorkbook.Worksheets("Feuil1").Range("B288")
        .FormulaR1C1 = "MEJ (en M€) montant max/secteur"
        
        For rw = 1 To rowN
            .Offset(2 * rw, 0).FormulaR1C1 = "Taux de sinistralité"
        Next rw
    End With
    
End Sub
