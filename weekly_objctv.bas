Attribute VB_Name = "weekly_objctv"
Sub weekly_obj_bckfill()
Attribute weekly_obj_bckfill.VB_ProcData.VB_Invoke_Func = " \n14"
Dim ws As Worksheet
Dim lr, lc As Integer
Dim dt1, dt2 As Date

Set ws = ThisWorkbook.Worksheets("Remaining Backfill")


lr = 25
lc = 14

For lc = 14 To 73
    
    If ws.Cells(lr, lc) >= Date And ws.Cells(lr + 1, lc) < Date Then
    
        dt1 = ws.Cells(lr, lc)
        dt2 = ws.Cells(lr + 1, lc)
        ws.Range("T50").Value = dt2
        ws.Range("U50").Value = dt1
        ws.Range("V50").Value = lc
        ws.PivotTables("RMN BCKFLL LST WEEK M3").PivotFields("ZONE"). _
        ShowDetail = False
        
        ws.PivotTables("RMN BCKFLL LST WEEK M3").PivotFields("DATE").ClearAllFilters
        ws.PivotTables("RMN BCKFLL LST WEEK M3").PivotFields("DATE").PivotFilters.Add Type:=xlDateBetween, Value1:=dt2 - 2, Value2:=dt1 - 2
            
            
        ws.Range("M51") = ws.Cells(27, lc)
        ws.Range("M52") = ws.Cells(28, lc)
        ws.Range("M53") = ws.Cells(29, lc)
        ws.Range("M54") = ws.Cells(30, lc)
        ws.Range("M55") = ws.Cells(31, lc)
        ws.Range("M56") = ws.Cells(32, lc)
        ws.Range("M57") = ws.Cells(33, lc)
        ws.Range("M58") = ws.Cells(34, lc)
        
        
        ws.Range("Q51") = ws.Cells(39, lc)
        ws.Range("Q52") = ws.Cells(40, lc)
        ws.Range("Q53") = ws.Cells(41, lc)
        ws.Range("Q54") = ws.Cells(42, lc)
        ws.Range("Q55") = ws.Cells(43, lc)
        ws.Range("Q56") = ws.Cells(44, lc)
        ws.Range("Q57") = ws.Cells(45, lc)
        ws.Range("Q58") = ws.Cells(46, lc)
        
        
        ws.PivotTables("RMN BCKFLL LST WEEK M3").PivotFields("ZONE"). _
        ShowDetail = True
        
    
    
        Exit Sub
    End If
    
Next lc

   
End Sub



Sub weekly_obj_bckfill_PREV()
Dim ws As Worksheet
Dim lr, lc As Integer
Dim dt1, dt2 As Date

Set ws = ThisWorkbook.Worksheets("Remaining Backfill")


lr = 25
lc = ws.Range("V50").Value - 1
dt1 = ws.Range("U50").Value - 7
dt2 = ws.Range("T50").Value - 7


        ws.Range("T50").Value = dt2
        ws.Range("U50").Value = dt1
        ws.Range("V50").Value = lc
        
    If lc >= 14 Then
        ws.PivotTables("RMN BCKFLL LST WEEK M3").PivotFields("ZONE"). _
        ShowDetail = False
        
        ws.PivotTables("RMN BCKFLL LST WEEK M3").PivotFields("DATE").ClearAllFilters
        ws.PivotTables("RMN BCKFLL LST WEEK M3").PivotFields("DATE").PivotFilters.Add Type:=xlDateBetween, Value1:=dt2 - 2, Value2:=dt1 - 2
            
            
        ws.Range("M51") = ws.Cells(27, lc)
        ws.Range("M52") = ws.Cells(28, lc)
        ws.Range("M53") = ws.Cells(29, lc)
        ws.Range("M54") = ws.Cells(30, lc)
        ws.Range("M55") = ws.Cells(31, lc)
        ws.Range("M56") = ws.Cells(32, lc)
        ws.Range("M57") = ws.Cells(33, lc)
        ws.Range("M58") = ws.Cells(34, lc)
        
        
        ws.Range("Q51") = ws.Cells(39, lc)
        ws.Range("Q52") = ws.Cells(40, lc)
        ws.Range("Q53") = ws.Cells(41, lc)
        ws.Range("Q54") = ws.Cells(42, lc)
        ws.Range("Q55") = ws.Cells(43, lc)
        ws.Range("Q56") = ws.Cells(44, lc)
        ws.Range("Q57") = ws.Cells(45, lc)
        ws.Range("Q58") = ws.Cells(46, lc)
        
        
        ws.PivotTables("RMN BCKFLL LST WEEK M3").PivotFields("ZONE"). _
        ShowDetail = True
        
    Else
    
        Call subClosingPopUp(1, "There is no planned data before this date!", "Limit of Plan", 1)
        Call weekly_obj_bckfill_NEXT
        
    End If
    

    


   
End Sub


Sub weekly_obj_bckfill_NEXT()
Dim ws As Worksheet
Dim lr, lc As Integer
Dim dt1, dt2 As Date

Set ws = ThisWorkbook.Worksheets("Remaining Backfill")


lr = 25
lc = ws.Range("V50").Value + 1
dt1 = ws.Range("U50").Value + 7
dt2 = ws.Range("T50").Value + 7


        ws.Range("T50").Value = dt2
        ws.Range("U50").Value = dt1
        ws.Range("V50").Value = lc
        
    If lc <= 73 Then
        ws.PivotTables("RMN BCKFLL LST WEEK M3").PivotFields("ZONE"). _
        ShowDetail = False
        
        ws.PivotTables("RMN BCKFLL LST WEEK M3").PivotFields("DATE").ClearAllFilters
        ws.PivotTables("RMN BCKFLL LST WEEK M3").PivotFields("DATE").PivotFilters.Add Type:=xlDateBetween, Value1:=dt2 - 2, Value2:=dt1 - 2
            
            
        ws.Range("M51") = ws.Cells(27, lc)
        ws.Range("M52") = ws.Cells(28, lc)
        ws.Range("M53") = ws.Cells(29, lc)
        ws.Range("M54") = ws.Cells(30, lc)
        ws.Range("M55") = ws.Cells(31, lc)
        ws.Range("M56") = ws.Cells(32, lc)
        ws.Range("M57") = ws.Cells(33, lc)
        ws.Range("M58") = ws.Cells(34, lc)
        
        
        ws.Range("Q51") = ws.Cells(39, lc)
        ws.Range("Q52") = ws.Cells(40, lc)
        ws.Range("Q53") = ws.Cells(41, lc)
        ws.Range("Q54") = ws.Cells(42, lc)
        ws.Range("Q55") = ws.Cells(43, lc)
        ws.Range("Q56") = ws.Cells(44, lc)
        ws.Range("Q57") = ws.Cells(45, lc)
        ws.Range("Q58") = ws.Cells(46, lc)
        
        
        ws.PivotTables("RMN BCKFLL LST WEEK M3").PivotFields("ZONE"). _
        ShowDetail = True
        
    Else
    
        Call subClosingPopUp(1, "Earthworks is finished", "End of Plan", 1)
        Call weekly_obj_bckfill_PREV
        
    End If
    
    


   
End Sub




