VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Admin_Box 
   Caption         =   "Dashboard"
   ClientHeight    =   10409
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   7049
   OleObjectBlob   =   "Admin_Box.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Admin_Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    
    Call Theme_Form(Me)
    Call Center_Form(Me)
    
    Call Set_Control_From_Large_Data_List(Products, Me.Sel_Product)
    Call Set_Control_From_Large_Data_List(Carriers, Me.Sel_Carrier)
    Call Set_Control_From_Small_Data_List("List_Plants", Me.Sel_Plant)
    Call Set_Control_From_Small_Data_List("List_Switchers", Me.Switcher_Sel)
    Call Set_Control_From_Small_Data_List("List_Themes", Me.Theme_Select)
    Call Set_Control_From_Small_Data_List("List_Weigh_In_Pounds_Products", Me.Sel_Pounds_Products)
    
    Me.Internal_Reweigh_Day_Limit_Entry = Strip_Array(ActiveWorkbook.Names("Option_InternalDayReweighLimit"))
    Me.Max_Entries_Entry = Strip_Array(ActiveWorkbook.Names("Option_Current_Max_Entries"))
    Me.Theme_Select = Strip_Array(ActiveWorkbook.Names("Option_Current_Theme"))
    
End Sub
Private Sub Max_Entries_Entry_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Set_Option_Small_List "Option_Current_Max_Entries", Me.Max_Entries_Entry

End Sub
Private Sub Internal_Reweigh_Day_Limit_Entry_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Set_Option_Small_List "Option_InternalDayReweighLimit", Me.Internal_Reweigh_Day_Limit_Entry

End Sub
Private Sub Product_Batch_Gen_From_Log_Buttom_Click()
    
    Call Generate_Data_List_From_Log(Products)
    
End Sub
Private Sub Employee_Button_Click()
    
    Plant_List_Type = "Employees"
    Plant_List_Box.Show
    
End Sub
Private Sub Product_Button_Click()
    
    Plant_List_Type = "Products"
    Plant_List_Box.Show
    
End Sub
Private Sub Reset_Log_Button_Click()
    
    Reset_Log_Box.Show
    
End Sub
Private Sub Add_Car_Button_Click()

    Call Add_Item_To_Large_List(Me.Sel_Carrier, Carriers, True)
    Call Sort_Item_Data_List(Carriers)
    Me.Sel_Carrier = vbNullString
    Me.Sel_Carrier.Clear
    Call Set_Control_From_Large_Data_List(Carriers, Me.Sel_Carrier)
    
End Sub
Private Sub Remove_Car_Button_Click()

    Call Remove_Item_From_Large_List(Me.Sel_Carrier, Carriers, True)
    Call Sort_Item_Data_List(Carriers)
    Me.Sel_Carrier = vbNullString
    Me.Sel_Carrier.Clear
    Call Set_Control_From_Large_Data_List(Carriers, Me.Sel_Carrier)
    
End Sub
Private Sub Carrier_Batch_Gen_From_Log_Buttom_Click()

    Call Generate_Data_List_From_Log(Carriers)

End Sub
Private Sub Add_Prod_Button_Click()

    Call Add_Item_To_Large_List(Me.Sel_Product, Products, True)
    Call Sort_Item_Data_List(Products)
    Me.Sel_Product = vbNullString
    Me.Sel_Product.Clear
    Call Set_Control_From_Large_Data_List(Products, Me.Sel_Product)

End Sub
Private Sub Remove_Prod_Button_Click()

    Call Remove_Item_From_Large_List(Me.Sel_Product, Products, True)
    Call Sort_Item_Data_List(Products)
    Me.Sel_Product = vbNullString
    Me.Sel_Product.Clear
    Call Set_Control_From_Large_Data_List(Products, Me.Sel_Product)

End Sub
Private Sub Add_Switcher_Button_Click()

    Call Add_Item_To_Small_List("List_Switchers", Me.Switcher_Sel, "Switcher", True)

End Sub
Private Sub Remove_Switcher_Button_Click()

    Call Remove_Item_From_Small_List("List_Switchers", Me.Switcher_Sel, "Switcher", True)

End Sub
Private Sub Add_Pounds_Products_Button_Click()

    Call Add_Item_To_Small_List("List_Weigh_In_Pounds_Products", Me.Sel_Pounds_Products, "Pound Product", True)

End Sub
Private Sub Remove_Pounds_Products_Button_Click()

    Call Remove_Item_From_Small_List("List_Weigh_In_Pounds_Products", Me.Sel_Pounds_Products, "Pound Product", True)

End Sub
Private Sub Add_Plant_Button_Click()

    Dim PlantMatch As Integer
    Dim msgbResponse As VbMsgBoxResult
    Dim Plant_Employee_List As String
    Dim RemovePlant As String
    
    RemovePlant = Me.Sel_Plant
    
    Set Current_List = ActiveWorkbook.Names("List_Plants")
    
    PlantMatch = InStr(1, Strip_Array(Current_List), RemovePlant, vbTextCompare)
    
    If PlantMatch = 0 Then
    
        Plant_Product_List = "List_Plant_" & Replace(RemovePlant, """", "") & "_Products"
        
        Plant_Employee_List = "List_Plant_" & Replace(RemovePlant, """", "") & "_Employees"
    
        Me.Sel_Plant.AddItem Me.Sel_Plant.value
        
        msgbResponse = MsgBox("Would you like to add a dedicated product list for this plant?", vbYesNo + vbQuestion, "Add product list?")
            
        If msgbResponse = vbYes Then
                
            Plant_Product_List = "List_Plant_" & RemovePlant & "_Products"
            
            ActiveWorkbook.Names.Add Name:=Plant_Product_List, RefersToR1C1:="="""""
                
        End If
            
        msgbResponse = MsgBox("Would you like to add a dedicated employee list for this plant?", vbYesNo + vbQuestion, "Add product list?")
            
        If msgbResponse = vbYes Then
                
            Plant_Employee_List = "List_Plant_" & RemovePlant & "_Employees"
            
            ActiveWorkbook.Names.Add Name:=Plant_Employee_List, RefersToR1C1:="="""""
                
        End If
        
        Make_List "List_Plants", Me.Sel_Plant, Me.Sel_Plant.ListCount
                    
        MsgBox "Plant Added", vbInformation, "Added"
        
        Me.Sel_Plant.value = vbNullString
        
    Else
    
        MsgBox "Plant already exists!", vbCritical, "Duplicated Plant Number"
    
    End If

End Sub
Private Sub Remove_Plant_Button_Click()

    Dim Plant_Employee_List As String
    Dim Current_Box_Count As Long
    Dim i As Integer
    
    If Me.Sel_Plant.value = vbNullString Then Exit Sub

    Current_Box_Count = Me.Sel_Plant.ListCount
    
    For i = 0 To Current_Box_Count - 1
    
        If Me.Sel_Plant.List(i) = Me.Sel_Plant.value Then
            
            Me.Sel_Plant.RemoveItem i
            
            Exit For
            
        End If
        
    Next i
    
    Current_Box_Count = Me.Sel_Plant.ListCount
    
    Make_List "List_Plants", Me.Sel_Plant, Me.Sel_Plant.ListCount
    
    Plant_Product_List = "List_Plant_" & Replace(Me.Sel_Plant.value, """", "") & "_Products"
    Plant_Employee_List = "List_Plant_" & Replace(Me.Sel_Plant.value, """", "") & "_Employees"
    
    On Error Resume Next
    
    For Each Named_Range In ActiveWorkbook.Names
    
        If Named_Range.Name = Plant_Product_List Then ActiveWorkbook.Names(Plant_Product_List).Delete
        
        If Named_Range.Name = Plant_Employee_List Then ActiveWorkbook.Names(Plant_Employee_List).Delete
        
    Next Named_Range
    
    On Error GoTo 0
    
    MsgBox "Plant Removed", vbInformation, "Removed"
    
    Me.Sel_Plant.value = vbNullString

End Sub
Private Sub Theme_Select_Change()

    Set_Option_Small_List "Option_Current_Theme", Me.Theme_Select.value

    Call Theme.Apply_Theme
    
    Call Theme_Form(Me)

End Sub

