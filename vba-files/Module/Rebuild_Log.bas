Attribute VB_Name = "Rebuild_Log"
Public Sub Make_Pretty()

    Call Uppercase
    Call Formulas
    Call Formating
    Call Buttons
    Call Theme.Apply_Theme
    Call Conditional_Formatting
    
End Sub
Public Sub Formulas()

    Dim Row_Number As Long
    
    Application.ScreenUpdating = False
    
    Current_Max_Entries = Strip_Array(ActiveWorkbook.Names("Option_Current_Max_Entries"))

    For Row_Number = 1 To Current_Max_Entries
   
        Call Main_Log_Formulas(Row_Number)
        
    Next

    [Internal_Log_1[ST-Ref]].Rows(1) = "=InternalRef([@ID])"
    [Internal_Log_2[CF-Ref]].Rows(1) = "=InternalRef([@ID])"

    Application.ScreenUpdating = True

End Sub
Public Sub Formating()
    
    With Sheets("Full Log").Range("A3:S2378").Font
    
            .Name = "Calibri"
            .FontStyle = "Bold"
            .Size = 12
            .Strikethrough = False
            .Subscript = False
            .Superscript = False
            
    End With
        
    With Sheets("Full Log").Range("A3:S2378")
            
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
            .Borders.ColorIndex = 1
        
    End With
        
    Sheets("Full Log").Range("R3:R2378").Borders(xlEdgeRight).LineStyle = xlLineStyleNone

End Sub
Public Sub Uppercase()
    
    Application.ScreenUpdating = False
    
    For Each cell In Sheets("Full Log").Range("Main_Log")
        
        If cell.HasFormula() = False Or Trim(cell.value) <> vbNullString Or cell.value = cell.value Or _
        IsError(cell) = False Then cell.value = cell.value
        
    Next cell
    
    Application.ScreenUpdating = True

End Sub
Public Sub Conditional_Formatting()

    Cells.FormatConditions.Delete
    
    Range("Main_Log[Status]").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""DONE"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Font
    
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    
    End With
    
    With Selection.FormatConditions(1).Interior
    
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
    
    End With
    
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("Main_Log[FS]").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="="""""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Font
    
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    
    End With
    
    With Selection.FormatConditions(1).Interior
    
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
    
    End With
    
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("Main_Log[ID]").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=($N3=""DONE"")"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Font
    
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    
    End With
    
    With Selection.FormatConditions(1).Interior
    
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        
    End With
    
    With Selection.FormatConditions(1).Borders(xlLeft)
    
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
        
    End With
    
    With Selection.FormatConditions(1).Borders(xlRight)
    
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
        
    End With
    
    With Selection.FormatConditions(1).Borders(xlTop)
    
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
        
    End With
    
    With Selection.FormatConditions(1).Borders(xlBottom)
    
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
        
    End With
    
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("Main_Log[Int Out]").Select
    
    Selection.FormatConditions.Add Type:=xlTextString, String:="New", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Font
    
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        
    End With
    
    With Selection.FormatConditions(1).Interior
    
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        
    End With
    
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("Main_Log[Int Out]").Select
    
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="REJECTED", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Font
    
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        
    End With
    
    With Selection.FormatConditions(1).Interior
    
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        
    End With
    
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("Main_Log[Int Out]").Select
    
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="Returned", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Font
    
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        
    End With
    
    With Selection.FormatConditions(1).Interior
    
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        
    End With
    
    Selection.FormatConditions(1).StopIfTrue = False
    
End Sub
Public Sub Buttons()

    With Full_Log.Add_Tank_Entry_Button
    
        .Font.Size = 12
        .Font.Bold = True
        .Top = 0
        .Height = 30
        .Width = 125
        .Left = 0
                
    End With
                
    With Full_Log.Weigh_Out_Tank_Entry_Button
    
        .Font.Size = 12
        .Font.Bold = True
        .Top = 0
        .Height = 30
        .Width = 125
        .Left = Full_Log.Add_Tank_Entry_Button.Width + 5
                
    End With
                
    With Full_Log.Edit_Tank_Entry_Button
    
        .Font.Size = 12
        .Font.Bold = True
        .Top = 0
        .Height = 30
        .Width = 95
        .Left = Full_Log.Weigh_Out_Tank_Entry_Button.Left + Full_Log.Weigh_Out_Tank_Entry_Button.Width + 5
                
    End With
                
    With Full_Log.Dashboard_Button
                
        .Font.Size = 12
        .Font.Bold = True
        .Top = 0
        .Height = 30
        .Width = 100
        .Left = Full_Log.Edit_Tank_Entry_Button.Left + Full_Log.Edit_Tank_Entry_Button.Width + 5
                
    End With
    
    With Full_Log.Next_Line_Button
                
        .Font.Size = 12
        .Font.Bold = True
        .Top = 0
        .Height = 30
        .Width = 100
        .Left = Full_Log.Dashboard_Button.Left + Full_Log.Dashboard_Button.Width + 5
                
    End With
    
    With Full_Log.Tools_Button
                
        .Font.Size = 12
        .Font.Bold = True
        .Top = 0
        .Height = 30
        .Width = 100
        .Left = Full_Log.Next_Line_Button.Left + Full_Log.Next_Line_Button.Width + 5
                
    End With
    
End Sub
