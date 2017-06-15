Private Sub Worksheet_Activate()
    
	' This will display a databar in each cell in a column. The maximum 
	' value can be set in the cell that is stored in the variable
	' maxNumber.
	
    Dim rc As Range
    Dim DataB As Databar
    Dim maxNumber As Long
    
    Set rc = Range("L3:L100")
    
    For Each c In rc.Cells
        If Intersect(c, rc) Is Nothing Then Exit Sub
        
        c.FormatConditions.Delete
            
        Set DataB = c.FormatConditions.AddDatabar
        DataB.BarFillType = xlDataBarFillGradient
        DataB.BarBorder.Type = xlDataBarBorderSolid
        DataB.BarBorder.Color.ThemeColor = xlThemeColorAccent1
        
        maxNumber = Cells(c.Row, c.Column - 5)
        
        DataB.MinPoint.Modify xlConditionValueNumber, 0
        DataB.MaxPoint.Modify xlConditionValueNumber, maxNumber
    Next
    
End Sub



