Attribute VB_Name = "Convert"
Sub ton_to_m3_Backfilling()

If Sheets("Backfilling total").Range("K10").Value <> "M3" Then
    
    'Button Property Changes
    'm3
    Sheets("BACKFILLING").Shapes.Range(Array("Rounded Rectangle 26")).TextFrame.Characters.Font.ColorIndex = 20
    Sheets("BACKFILLING").Shapes.Range(Array("Rounded Rectangle 26")).Fill.ForeColor.RGB = RGB(143, 69, 199)
    'ton
    Sheets("BACKFILLING").Shapes.Range(Array("Rounded Rectangle 25")).TextFrame.Characters.Font.ColorIndex = 1
    Sheets("BACKFILLING").Shapes.Range(Array("Rounded Rectangle 25")).Fill.ForeColor.RGB = RGB(150, 150, 150)
    
    
    'Pivot Table Changes
    Sheets("Backfilling total").PivotTables("PivotTable1").PivotSelect "'Row Grand Total'", _
        xlDataAndLabel, True
    Sheets("Backfilling total").PivotTables("PivotTable1").PivotFields("Sum of TON").Orientation = _
        xlHidden
    Sheets("Backfilling total").PivotTables("PivotTable1").AddDataField Sheets("Backfilling total").PivotTables( _
        "PivotTable1").PivotFields("m3"), "Sum of m3", xlSum
    With Sheets("Backfilling total").PivotTables("PivotTable1").PivotFields("Sum of m3")
        .NumberFormat = "#,##0.00"
    End With
    
    Sheets("Backfilling per zones").PivotTables("PivotTable2").PivotFields("Sum of TON").Orientation = _
        xlHidden
    Sheets("Backfilling per zones").PivotTables("PivotTable2").AddDataField Sheets("Backfilling per zones").PivotTables( _
        "PivotTable2").PivotFields("m3"), "Sum of m3", xlSum
    
    With Sheets("Backfilling per zones").PivotTables("PivotTable2").PivotFields("Sum of m3")
        .NumberFormat = "#,##0.00"
    End With
    
    Sheets("Backfilling in time").PivotTables("PivotTable1").PivotFields("Sum of TON").Orientation = _
        xlHidden
    Sheets("Backfilling in time").PivotTables("PivotTable1").AddDataField Sheets("Backfilling in time").PivotTables( _
        "PivotTable1").PivotFields("m3"), "Sum of m3", xlSum
    
    With Sheets("Backfilling in time").PivotTables("PivotTable1").PivotFields("Sum of m3")
        .NumberFormat = "#,##0.00"
    End With
    
    Sheets("Backfilling total").Range("K10").Value = "M3"
End If

    
End Sub

Sub m3_to_ton_Backfilling()

If Sheets("Backfilling total").Range("K10").Value <> "TON" Then
    
     'Button Property Changes
    'Button Property Changes
    'ton
    Sheets("BACKFILLING").Shapes.Range(Array("Rounded Rectangle 25")).TextFrame.Characters.Font.ColorIndex = 20
    Sheets("BACKFILLING").Shapes.Range(Array("Rounded Rectangle 25")).Fill.ForeColor.RGB = RGB(143, 69, 199)
    'm3
    Sheets("BACKFILLING").Shapes.Range(Array("Rounded Rectangle 26")).TextFrame.Characters.Font.ColorIndex = 1
    Sheets("BACKFILLING").Shapes.Range(Array("Rounded Rectangle 26")).Fill.ForeColor.RGB = RGB(150, 150, 150)
    
    
    'Pivot Table Changes
    
    Sheets("Backfilling total").PivotTables("PivotTable1").PivotSelect "'Row Grand Total'", _
        xlDataAndLabel, True
    Sheets("Backfilling total").PivotTables("PivotTable1").PivotFields("Sum of m3").Orientation = _
        xlHidden
    Sheets("Backfilling total").PivotTables("PivotTable1").AddDataField Sheets("Backfilling total").PivotTables( _
        "PivotTable1").PivotFields("TON"), "Sum of TON", xlSum
    With Sheets("Backfilling total").PivotTables("PivotTable1").PivotFields("Sum of TON")
        .NumberFormat = "#,##0.00"
    End With
    
    Sheets("Backfilling per zones").PivotTables("PivotTable2").PivotFields("Sum of m3").Orientation = _
        xlHidden
    Sheets("Backfilling per zones").PivotTables("PivotTable2").AddDataField Sheets("Backfilling per zones").PivotTables( _
        "PivotTable2").PivotFields("TON"), "Sum of TON", xlSum
    
    With Sheets("Backfilling per zones").PivotTables("PivotTable2").PivotFields("Sum of TON")
        .NumberFormat = "#,##0.00"
    End With
    
    Sheets("Backfilling in time").PivotTables("PivotTable1").PivotFields("Sum of m3").Orientation = _
        xlHidden
    Sheets("Backfilling in time").PivotTables("PivotTable1").AddDataField Sheets("Backfilling in time").PivotTables( _
        "PivotTable1").PivotFields("TON"), "Sum of TON", xlSum
    
    With Sheets("Backfilling in time").PivotTables("PivotTable1").PivotFields("Sum of TON")
        .NumberFormat = "#,##0.00"
    End With

    
    
    Sheets("Backfilling total").Range("K10").Value = "TON"
    
End If


End Sub






Sub ton_to_m3_incoming()

If Sheets("incoming(total)").Range("L7").Value <> "M3" Then




    'Button Property Changes
    'm3
    Sheets("INCOMING").Shapes.Range(Array("Rounded Rectangle 22")).TextFrame.Characters.Font.ColorIndex = 20
    Sheets("INCOMING").Shapes.Range(Array("Rounded Rectangle 22")).Fill.ForeColor.RGB = RGB(143, 69, 199)
    'ton
    Sheets("INCOMING").Shapes.Range(Array("Rounded Rectangle 21")).TextFrame.Characters.Font.ColorIndex = 1
    Sheets("INCOMING").Shapes.Range(Array("Rounded Rectangle 21")).Fill.ForeColor.RGB = RGB(150, 150, 150)
    

    
    Sheets("incoming(total)").PivotTables("PivotTable1").PivotSelect "", xlDataAndLabel, True
    Sheets("incoming(total)").PivotTables("PivotTable1").PivotFields("Sum of Ton"). _
        Orientation = xlHidden
    Sheets("incoming(total)").PivotTables("PivotTable1").AddDataField Sheets("incoming(total)").PivotTables( _
        "PivotTable1").PivotFields("M3"), "Sum of M3", xlSum
    
    With Sheets("incoming(total)").PivotTables("PivotTable1").PivotFields("Sum of M3")
        .NumberFormat = "#,##0.00"
    End With
    
    Sheets("incoming nesma_sc").PivotTables("PivotTable1").PivotFields("Sum of Ton"). _
        Orientation = xlHidden
    Sheets("incoming nesma_sc").PivotTables("PivotTable1").AddDataField Sheets("incoming nesma_sc").PivotTables( _
        "PivotTable1").PivotFields("M3"), "Sum of M3", xlSum
    
    With Sheets("incoming nesma_sc").PivotTables("PivotTable1").PivotFields("Sum of M3")
        .NumberFormat = "#,##0.00"
    End With
    
    Sheets("incoming by company").PivotTables("PivotTable1").PivotFields("Sum of Ton"). _
        Orientation = xlHidden
    Sheets("incoming by company").PivotTables("PivotTable1").AddDataField Sheets("incoming by company").PivotTables( _
        "PivotTable1").PivotFields("M3"), "Sum of M3", xlSum
    With Sheets("incoming by company").PivotTables("PivotTable1").PivotFields("Sum of M3")
        .NumberFormat = "#,##0.00"
    End With
'
    Sheets("incoming per zones").PivotTables("PivotTable2").PivotFields("Sum of Ton"). _
        Orientation = xlHidden
    Sheets("incoming per zones").PivotTables("PivotTable2").AddDataField Sheets("incoming per zones").PivotTables( _
        "PivotTable2").PivotFields("M3"), "Sum of M3", xlSum
    Range("B11").Select
    With Sheets("incoming per zones").PivotTables("PivotTable2").PivotFields("Sum of M3")
        .NumberFormat = "#,##0.00"
    End With
    
    Sheets("incoming(total)").Range("L7").Value = "M3"
    
End If

    
End Sub




Sub m3_to_ton_incoming()

If Sheets("incoming(total)").Range("L7").Value <> "TON" Then


    'Button Property Changes
    'ton
    Sheets("INCOMING").Shapes.Range(Array("Rounded Rectangle 21")).TextFrame.Characters.Font.ColorIndex = 20
    Sheets("INCOMING").Shapes.Range(Array("Rounded Rectangle 21")).Fill.ForeColor.RGB = RGB(143, 69, 199)
    'm3
    Sheets("INCOMING").Shapes.Range(Array("Rounded Rectangle 22")).TextFrame.Characters.Font.ColorIndex = 1
    Sheets("INCOMING").Shapes.Range(Array("Rounded Rectangle 22")).Fill.ForeColor.RGB = RGB(150, 150, 150)
    
    

    
    Sheets("incoming(total)").PivotTables("PivotTable1").PivotSelect "", xlDataAndLabel, True
    Sheets("incoming(total)").PivotTables("PivotTable1").PivotFields("Sum of M3"). _
        Orientation = xlHidden
    Sheets("incoming(total)").PivotTables("PivotTable1").AddDataField Sheets("incoming(total)").PivotTables( _
        "PivotTable1").PivotFields("Ton"), "Sum of Ton", xlSum
    
    With Sheets("incoming(total)").PivotTables("PivotTable1").PivotFields("Sum of Ton")
        .NumberFormat = "#,##0.00"
    End With
    
    Sheets("incoming nesma_sc").PivotTables("PivotTable1").PivotFields("Sum of M3"). _
        Orientation = xlHidden
    Sheets("incoming nesma_sc").PivotTables("PivotTable1").AddDataField Sheets("incoming nesma_sc").PivotTables( _
        "PivotTable1").PivotFields("Ton"), "Sum of Ton", xlSum
    
    With Sheets("incoming nesma_sc").PivotTables("PivotTable1").PivotFields("Sum of Ton")
        .NumberFormat = "#,##0.00"
    End With
    
    Sheets("incoming by company").PivotTables("PivotTable1").PivotFields("Sum of M3"). _
        Orientation = xlHidden
    Sheets("incoming by company").PivotTables("PivotTable1").AddDataField Sheets("incoming by company").PivotTables( _
        "PivotTable1").PivotFields("Ton"), "Sum of Ton", xlSum
    With Sheets("incoming by company").PivotTables("PivotTable1").PivotFields("Sum of Ton")
        .NumberFormat = "#,##0.00"
    End With
    
    Sheets("incoming per zones").PivotTables("PivotTable2").PivotFields("Sum of M3"). _
        Orientation = xlHidden
    Sheets("incoming per zones").PivotTables("PivotTable2").AddDataField Sheets("incoming per zones").PivotTables( _
        "PivotTable2").PivotFields("Ton"), "Sum of Ton", xlSum
    Range("B11").Select
    With Sheets("incoming per zones").PivotTables("PivotTable2").PivotFields("Sum of Ton")
        .NumberFormat = "#,##0.00"
    End With
    
    Sheets("incoming(total)").Range("L7").Value = "TON"
    
End If
 
    
    
    
End Sub


