Attribute VB_Name = "Theme"
Dim Accent_Color As Long 'border colors,First row of Truck Log Sheet
Dim Font_Color_Secondary As Long 'For Button Fonts
Dim Button_Color As Long
Public Sub Quick_Change_Theme()
Attribute Quick_Change_Theme.VB_ProcData.VB_Invoke_Func = "t\n14"

    Dim New_Theme As String
    
    Select Case Strip_Array(ActiveWorkbook.Names("Option_Current_Theme"))
    
        Case Black: New_Theme = "Blackout"
        
        Case Blackout: New_Theme = "Black"
    
    End Select
    
    Set_Option_Small_List "Option_Current_Theme", New_Theme
    
    Call Apply_Theme
    
End Sub
Public Sub Apply_Theme()

    Dim Table_Theme As String
   
    Select Case Strip_Array(ActiveWorkbook.Names("Option_Current_Theme"))

        Case "Black":
       
            Button_Color = vbBlack
            Font_Color_Secondary = vbWhite 'For Button Fonts
            Table_Theme = "LG-Black"
           
        Case "Blackout":
    
            Button_Color = vbBlack
            Font_Color_Secondary = vbWhite 'For Button Fonts
            Table_Theme = "LG-Blackout"
       
    End Select
   
    ThisWorkbook.Sheets("Full Log").Activate
   
    Application.ScreenUpdating = False
   
    ThisWorkbook.Sheets("Full Log").Cells.ClearFormats 'clears any changes made

    With Full_Log.Add_Tank_Entry_Button
               
        .BackColor = Button_Color
        .ForeColor = Font_Color_Secondary
   
    End With
               
    With Full_Log.Weigh_Out_Tank_Entry_Button
               
        .BackColor = Button_Color
        .ForeColor = Font_Color_Secondary
   
    End With
               
    With Full_Log.Edit_Tank_Entry_Button
               
        .BackColor = Button_Color
        .ForeColor = Font_Color_Secondary
   
    End With
               
    With Full_Log.Dashboard_Button
               
        .BackColor = Button_Color
        .ForeColor = Font_Color_Secondary
   
    End With
   
    With Full_Log.Next_Line_Button
               
        .BackColor = Button_Color
        .ForeColor = Font_Color_Secondary
   
    End With
    
    With Full_Log.Tools_Button
               
        .BackColor = Button_Color
        .ForeColor = Font_Color_Secondary
   
    End With
    
    Sheets("Full Log").ListObjects("Main_Log").TableStyle = Table_Theme
    Sheets("Storage Log").ListObjects("Internal_Log_1").TableStyle = Table_Theme
    Sheets("CFS Log").ListObjects("Internal_Log_2").TableStyle = Table_Theme
    
    
    Select Case Current_Theme

        Case Black:
       
            Accent_Color = 1
            Font_Color = 2
           
        Case Blackout:
   
            Accent_Color = 1
            Font_Color = 2
       
    End Select
   
    With ThisWorkbook.Sheets("Full Log").Rows("1:1")
           
        .Interior.ColorIndex = Accent_Color
        .Font.ColorIndex = Font_Color
               
    End With
   
    Call Theme.Format_Cells
    Call Rebuild_Log.Conditional_Formatting
   
    Application.ScreenUpdating = True

End Sub
Public Sub Format_Cells()
    
    With Range("Main_Log[#All]")
    
        .Font.Size = 11
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        
    End With
    
    With Range("Table_Next_ID[#All]")
    
        .Font.Size = 11
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        
    End With
    
    With [Main_Log[ID]]
    
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
                
    End With
    
    With Sheets("Full Log").ListObjects("Main_Log")
    
        .HeaderRowRange.VerticalAlignment = xlCenter
        .HeaderRowRange.HorizontalAlignment = xlCenter
    
    End With
    
    With Sheets("Storage Log").ListObjects("Internal_Log_1")
    
        .HeaderRowRange.VerticalAlignment = xlCenter
        .HeaderRowRange.HorizontalAlignment = xlCenter
    
    End With
    
    With Range("Internal_Log_1[#All]")
    
        .Font.Size = 11
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        
    End With
    
End Sub
Public Sub Theme_Form(ByVal Current_Form As Object)

    Dim Current_Controller As Control
    Dim Font_Color As Long 'Labels,Check Boxes
    Dim Main_BackGround_Color As Long 'Main color of form
    Dim Secondary_BackGround_Color As Long 'textbox and combobox
    
    Select Case Strip_Array(ActiveWorkbook.Names("Option_Current_Theme"))
    
        Case "Black": Current_Theme = Black
        
        Case "Blackout": Current_Theme = Blackout
    
    End Select
    
    Select Case Current_Theme
   
        Case Black:
       
            Font_Color = vbBlack
            Button_Color = vbBlack
            Font_Color_Secondary = vbWhite
            Main_BackGround_Color = vbWhite
            Secondary_BackGround_Color = vbWhite
            Accent_Color = vbBlack
            
        Case Blackout:
   
            Font_Color = vbWhite
            Button_Color = vbBlack
            Font_Color_Secondary = vbWhite
            Main_BackGround_Color = vbBlack
            Secondary_BackGround_Color = vbBlack
            Accent_Color = vbWhite
           
    End Select

    For Each Current_Controller In Current_Form.Controls
       
            Select Case TypeName(Current_Controller)
           
                Case "Label":

                    Current_Controller.ForeColor = Font_Color
               
                Case "CommandButton":
               
                    Current_Controller.BackStyle = 1
                   
                    Current_Controller.ForeColor = Font_Color_Secondary
                   
                    If Current_Controller.Name <> "Reset_Log_Button" Then
                   
                        Current_Controller.BackColor = Button_Color
                   
                    End If
                   
                Case "Frame":
               
                    Current_Controller.ForeColor = Font_Color
                    Current_Controller.BackColor = Main_BackGround_Color
                   
                Case "TextBox", "ComboBox":
               
                    Current_Controller.ForeColor = Font_Color
                    Current_Controller.BackColor = Secondary_BackGround_Color
               
                Case "CheckBox":
               
                    Current_Controller.ForeColor = Font_Color
                   
                Case "ListBox":
               
                    Current_Controller.ForeColor = Font_Color
                    Current_Controller.BackColor = Secondary_BackGround_Color
                    Current_Controller.BorderColor = Accent_Color
                   
                Case "OptionButton":
               
                    Current_Controller.ForeColor = Font_Color
               
            End Select
       
        Next Current_Controller
       
        Current_Form.BorderColor = Accent_Color
        Current_Form.BorderStyle = 1
        Current_Form.BackColor = Main_BackGround_Color

End Sub
