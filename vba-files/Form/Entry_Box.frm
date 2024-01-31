VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Entry_Box 
   Caption         =   "Tank Entry"
   ClientHeight    =   7980
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   9478.001
   OleObjectBlob   =   "Entry_Box.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Entry_Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Current_Entry As Main_Log_Entry
Public Current_Internal_Entry As Internal_Log_Entry
Private Sub UserForm_Initialize()

    Call Theme_Form(Me)
    Call Center_Form(Me)
    
    Set Current_Entry = New Main_Log_Entry
    
    Call Set_Data_From_Entry
    
    If Not Process_Mode = Out Then Call Set_Available_Prefixes
    
    Call Set_Entry_ID_Prefix_Control
    
    Call Set_Control_From_Large_Data_List(Products, Sel_Product)
    
    Call Set_Control_From_Small_Data_List("List_Plants", Sel_Plant)
    
    If Not Process_Mode = Add_New Then
    
        If Current_Entry.Primary_Type = Internal Then
            
            Set Current_Internal_Entry = New Internal_Log_Entry
            Current_Internal_Entry.Initialize_Main_Data Current_Entry
    
            Call Set_Control_From_Small_Data_List("List_InternalStatus", Internal_Status_Sel)
            
        End If
    
    End If
    
    Call Change_Form
    
End Sub
Private Sub Set_Available_Prefixes()

    Select Case Left(Next_Storage_ID, 1)
    
        Case PREFIX_STORAGE_ID: Entry_ID_Prefix.AddItem "H - Internal/House Tank"
       
        Case PREFIX_STORAGE_ID_2: Entry_ID_Prefix.AddItem "I - Internal/House Tank"
    
    End Select
    
    Select Case Left(Next_Drop_ID, 1)
    
        Case PREFIX_DROP_ID: Entry_ID_Prefix.AddItem "D - Dropped External Tank"
       
        Case PREFIX_DROP_ID_2: Entry_ID_Prefix.AddItem "T - Dropped External Tank"
    
    End Select
    
    Select Case Left(Next_Central_ID, 1)
    
        Case PREFIX_CENTRAL_ID: Entry_ID_Prefix.AddItem "C - Central Fill Station"
       
        Case PREFIX_CENTRAL_ID_2: Entry_ID_Prefix.AddItem "F - Central Fill Station"
    
    End Select
    
    Entry_ID_Prefix.AddItem "Live Unload/Load"

End Sub
Private Sub Set_Entry_ID_Prefix_Control()

    Select Case Left([Main_Log[ID]].Rows(Current_Entry.Row_Number), 1)
                
        Case PREFIX_STORAGE_ID: Entry_ID_Prefix = "H - Internal/House Tank"
                                            
        Case PREFIX_STORAGE_ID_2: Entry_ID_Prefix = "I - Internal/House Tank"
                                            
        Case PREFIX_CENTRAL_ID: Entry_ID_Prefix = "C - Central Fill Station"
                                        
        Case PREFIX_CENTRAL_ID_2: Entry_ID_Prefix = "F - Central Fill Station"
                                            
        Case PREFIX_DROP_ID: Entry_ID_Prefix = "D - Dropped External Tank"
                                            
        Case PREFIX_DROP_ID_2: Entry_ID_Prefix = "T - Dropped External Tank"
                            
        Case Else: Entry_ID_Prefix = "Live Unload/Load"
                    
    End Select
            
End Sub
Private Sub Set_Data_From_Entry()

    Entry_Number = [Main_Log[ID]].Rows(Current_Entry.Row_Number)
    
    If Not Process_Mode = Add_New Then
    
        Selector_Carrier = [Main_Log[Carrier]].Rows(Current_Entry.Row_Number)
        Entry_Tank_Number = [Main_Log[Tank '#]].Rows(Current_Entry.Row_Number)
        Entry_Truck = [Main_Log[Truck '#]].Rows(Current_Entry.Row_Number)
        Entry_In_Weight = Replace([Main_Log[Weight]].Rows(Current_Entry.Row_Number), "LBS", "")
        Sel_Product = [Main_Log[Product Name]].Rows(Current_Entry.Row_Number)
        Sel_Plant = [Main_Log[PLT '#]].Rows(Current_Entry.Row_Number)
        Entry_Date_In = Format([Main_Log[Date In]].Rows(Current_Entry.Row_Number), DEFAULT_DATE_FORMAT)
        Entry_Time_In = Format([Main_Log[Time In]].Rows(Current_Entry.Row_Number), DEFAULT_TIME_FORMAT)
        Selector_Notified = [Main_Log[Notified]].Rows(Current_Entry.Row_Number)
        Entry_initials_In = [Main_Log[Int In]].Rows(Current_Entry.Row_Number)
        Entry_Net_Weight = [Main_Log[Net Weight]].Rows(Current_Entry.Row_Number)
        Entry_initials_Out = [Main_Log[Int Out]].Rows(Current_Entry.Row_Number)
        Ref_ID_Code = [Main_Log[RefID]].Rows(Current_Entry.Row_Number)
        
        If Process_Mode = Edit Then
            
            Entry_Date_Out = Format([Main_Log[Date Out]].Rows(Current_Entry.Row_Number), DEFAULT_DATE_FORMAT)
            Entry_Time_Out = Format([Main_Log[Time Out]].Rows(Current_Entry.Row_Number), DEFAULT_TIME_FORMAT)
            
        End If
        
        If Right([Main_Log[Weight]].Rows(Current_Entry.Row_Number), 3) = "LBS" Then Check_Is_Pounds = True
        
    End If
    
End Sub
Sub Change_Form()

    Const FORM_SHORT = 310
    Const FORM_SHORT_WIDTH = 310
    Const FORM_FULL_WIDTH = 483
    Const DROPPED_WEIGHED As String = "DW"
    
    Select Case Process_Mode
    
        Case Add_New:
            
            Caption = "Add Tank Entry"
            
            Select Case Current_Entry.Primary_Type
            
                Case External:
                
                    Entry_Tank_Number.RowSource = Empty
                    Entry_Tank_Number.ShowDropButtonWhen = fmShowDropButtonWhenNever
                    
                    Notified_Label = "Notified"
                    Selector_Notified.Clear
                    Selector_Notified.ShowDropButtonWhen = fmShowDropButtonWhenNever
                    
                    Call Remove_Data_List_From_Control(Selector_Carrier)
                    Call Set_Control_From_Large_Data_List(Carriers, Selector_Carrier)
                    
                    Entry_Prev_ID_Date = Empty
                    Batch_Num_Entry = Empty
                    Internal_Status_Sel = Empty
                    Internal_Ref_ID = Empty
                                
                    Entry_Number.Enabled = False
                    Entry_Date_Out.Enabled = False
                    Entry_Time_Out.Enabled = False
                    Entry_Net_Weight.Enabled = False
                    Entry_initials_Out.Enabled = False
                    Out_Info_Frame.Visible = False
                    Confirm_Entry_Button.Top = 246
                    Me.Height = FORM_SHORT
                    Me.Width = FORM_SHORT_WIDTH
                    
                    Entry_Date_In = Format(Date, DEFAULT_DATE_FORMAT)
                    Entry_Time_In = Format(Time, DEFAULT_TIME_FORMAT)
                    
                    Select Case Current_Entry.Secondary_Type
                    
                        Case Drop:
                        
                            Entry_Truck = DROPPED_WEIGHED
                            Entry_Truck.Enabled = False
                            
                        Case Live:

                            If Entry_Truck = DROPPED_WEIGHED Then Entry_Truck = Empty
                            Entry_Truck.Enabled = True
                            
                    End Select
                    
                Case Internal:
                
                    Entry_Truck = DROPPED_WEIGHED
                    Entry_Truck.Enabled = False

                    Internal_Status_lbl.Visible = False
                    Internal_Status_Sel.Visible = False
                    
                    Call Remove_Data_List_From_Control(Selector_Carrier)
                    Call Set_Control_From_Large_Data_List(Internal_Carriers, Selector_Carrier)
                    Call Set_Control_From_Small_Data_List("List_Switchers", Selector_Notified)
                    
                    Selector_Notified.ShowDropButtonWhen = fmShowDropButtonWhenAlways
                    Entry_Tank_Number.ShowDropButtonWhen = fmShowDropButtonWhenAlways
                    
                    Select Case Current_Entry.Secondary_Type
                    
                        Case Storage:
                        
                            Entry_Tank_Number.RowSource = "Internal_Log_1[Tank '#]"
                            
                        Case Central: Entry_Tank_Number.RowSource = "Internal_Log_2[Tank '#]"
                            
                    End Select
                    
                    Select Case Current_Internal_Entry.Internal_Sub_Type
                            
                        Case New_Tank: Notified_Label = "Notified"
                                    
                        Case Returning_Tank:
                        
                            Notified_Label = "Notified"
                        
                            Width = FORM_FULL_WIDTH
                                    
                        Case Current_Tank:
                        
                            Notified_Label = "Switcher"
                        
                            Width = FORM_FULL_WIDTH
                                    
                    End Select
                    
            End Select

        Case Edit:
        
            Caption = "Edit Tank Entry"
            
            Entry_ID_Prefix.Style = fmStyleDropDownCombo
            
            Select Case Current_Entry.Primary_Type
            
                Case External: Notified_Label = "Notified"
                
            End Select
        
        Case Out:
        
            Caption = "Weigh Out Tank Entry"
            
            Entry_ID_Prefix.Enabled = False
            Entry_Number.Enabled = False
            Selector_Carrier.Enabled = False
            Entry_Tank_Number.Enabled = False
            Entry_Truck.Enabled = False
            Entry_In_Weight.Enabled = False
            Sel_Product.Enabled = False
            Sel_Plant.Enabled = False
            Entry_Date_In.Enabled = False
            Entry_Time_In.Enabled = False
            Selector_Notified.Enabled = False
            Entry_initials_In.Enabled = False
            Check_Is_Pounds.Enabled = False

            Entry_Date_Out = Format(Date, DEFAULT_DATE_FORMAT)
            Entry_Time_Out = Format(Time, DEFAULT_TIME_FORMAT)
            
            Select Case Current_Entry.Primary_Type
            
                Case External:
                
                    Width = FORM_SHORT_WIDTH
                    
                    Check_Reject_Entry.Visible = True
                    
                    If Current_Entry.Row_Number <> Count_Entries Then Check_Reset_Entry.Visible = False
                    
                Case Internal:
                
                    Width = FORM_FULL_WIDTH
                    
                    Internal_Status_lbl.Visible = True
                    Internal_Status_Sel.Visible = True
                
            End Select
        
    End Select
    
    If Current_Entry.Primary_Type = Internal Then
    
        If (Current_Internal_Entry.Internal_Type = Storage And Current_Internal_Entry.Internal_Sub_Type = Current_Tank) Then
                        
            Batch_Num_Entry.Visible = False
            Batch_Num_lbl.Visible = False
                            
        ElseIf (Current_Internal_Entry.Internal_Type = Central And Current_Internal_Entry.Internal_Sub_Type = Current_Tank) Then
                        
            Batch_Num_Entry.Visible = True
            Batch_Num_lbl.Visible = True
                            
        Else
                        
            Batch_Num_Entry.Visible = False
            Batch_Num_lbl.Visible = False
                            
        End If
    
    End If
    
End Sub
Private Sub Entry_ID_Prefix_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    If Current_Entry.Primary_Type = Internal Then
        
        Set Current_Internal_Entry = New Internal_Log_Entry
        
        Current_Internal_Entry.Initialize_Main_Data Current_Entry
        
    Else
    
        Entry_Tank_Number = Empty
        Selector_Carrier = Empty
        Entry_Prev_ID_Date = Empty
        Batch_Num_Entry = Empty
        Internal_Status_Sel = Empty
        Entry_Internal_Ref_ID = Empty
        Set Current_Internal_Entry = Nothing
    
    End If
    
    Call Change_Form
    
End Sub
Private Sub Sel_Product_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    If Current_Entry.Primary_Type = Internal Or Trim(Sel_Product) = Empty Then Exit Sub
    
    Sel_Plant = Find_Plant_For_Product
        
End Sub
Private Function Find_Plant_For_Product() As String

    Dim List_Data() As String
    Dim Plant_List() As String
    Dim List_Item As Integer
    Dim Item_Index As Integer
    
    Set Current_List = ActiveWorkbook.Names("List_Plants")
        
    Plant_List = Split(Strip_Array(Current_List), ",")
    
    For List_Item = LBound(Plant_List) To UBound(Plant_List)
        
        If Check_Plant_For_Product_List(Plant_List(List_Item)) = True Then
            
            Plant_Product_List = "List_Plant_" & Plant_List(List_Item) & "_Products"
            
            List_Data = Split(Strip_Array(ActiveWorkbook.Names(Plant_Product_List)), ",")
                        
            For Item_Index = 0 To UBound(List_Data)
                        
                If Sel_Product = Replace(List_Data(Item_Index), """", "") Then
                        
                    Find_Plant_For_Product = Plant_List(List_Item)
                            
                    Exit Function
                            
                End If
                                        
            Next
                    
        End If
        
    Next List_Item

End Function
Private Sub Sel_Plant_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Current_Entry.Primary_Type = Internal Or Sel_Plant = Empty Then Exit Sub
    
    Dim Plant_Employee_List As String
    
    On Error Resume Next

    Selector_Notified = Empty
    
    Plant_Employee_List = "List_Plant_" & Sel_Plant & "_Employees"
    
    Call Set_Control_From_Small_Data_List(Plant_Employee_List, Selector_Notified)
    
    On Error GoTo 0

End Sub
Private Sub Entry_Date_In_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Entry_Date_In = Format(Entry_Date_In, DEFAULT_DATE_FORMAT)

End Sub
Private Sub Entry_Time_In_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Entry_Time_In = Format(Entry_Time_In, DEFAULT_TIME_FORMAT)

End Sub
Private Sub Entry_Date_Out_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Entry_Date_Out = Format(Entry_Date_Out, DEFAULT_DATE_FORMAT)

End Sub
Private Sub Entry_Time_Out_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Entry_Time_Out = Format(Entry_Time_Out, DEFAULT_TIME_FORMAT)

End Sub
Private Sub Entry_Tank_Number_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    If Current_Entry.Primary_Type = Internal Then

        Current_Internal_Entry.Initialize_Secondary_Data
        
        If Current_Internal_Entry.Internal_Sub_Type = Current_Tank Then Current_Internal_Entry.Set_Previous_Entry_Data
        
        Call Change_Form
        
    End If

End Sub
Private Sub Check_Reject_Entry_Click()
    
    If Check_Reject_Entry = True Then
    
        Entry_Net_Weight = 0
        Entry_initials_Out = "Rejected"
        Entry_Date_Out.Enabled = False
        Entry_Time_Out.Enabled = False
        Entry_Net_Weight.Enabled = False
        Entry_initials_Out.Enabled = False
        
    Else
    
        Entry_Date_Out = Empty
        Entry_Time_Out = Empty
        Entry_Net_Weight = Empty
        Entry_initials_Out = Empty
        Entry_Date_Out.Enabled = True
        Entry_Time_Out.Enabled = True
        Entry_Net_Weight.Enabled = True
        Entry_initials_Out.Enabled = True
        
    End If

End Sub
Private Sub Confirm_Entry_Button_Click()
    
    If Check_Reset_Entry = True Then
    
        Call Reset_Row(Current_Entry.Row_Number)
        
        Call Update_Log
        
        Exit Sub
        
    End If
    
    If Check_If_Controls_Are_Blank <> False And Process_Mode <> Edit Then Exit Sub
    
    If IsNumeric(Entry_In_Weight) <> True And Trim(Entry_In_Weight) <> Empty Then
    
        MsgBox "This weight is not a number.", vbCritical, "Invalid Weight"
        
        Exit Sub
        
    End If
    
    If Check_Reject_Entry = True Then
        
        Current_Entry.Out_Data_To_Main_Log
        
        [Main_Log[Status]].Rows(Current_Entry.Row_Number) = ENTRY_INACTIVE
        
        Call Update_Log
        
        Exit Sub
        
    End If
    
    If Current_Entry.Primary_Type = Internal Then
    
        If Process_Mode = Out And Internal_Status_Sel = Empty Then
                            
            MsgBox "Unless the tank is leaving site or needs repair Internal Tanks can only be weighed out using the built in process and not through the weight out process"
            Internal_Status_Sel.BackColor = vbYellow
            
            Exit Sub
                        
        End If
    
        If Current_Entry.Secondary_Type = Central And Current_Internal_Entry.Internal_Sub_Type = Current_Tank And Batch_Num_Entry = Empty Then
            
            Batch_Num_Entry.BackColor = vbYellow
            
            MsgBox "This is a CFS tank and no batch number has been entered. Please provide one to continue." & _
            " If this is a new or returning tank please select the corresponding option.", vbCritical, "No Batch Number Found!"
                    
            Exit Sub
                    
        End If
    
    End If
    
    Call Transfer_Data
    Call Update_Log
    
End Sub
Private Sub Transfer_Data()
    
    If Process_Mode = Add_New Then Current_Entry.In_Data_To_Main_Log Else Current_Entry.Out_Data_To_Main_Log
    
    If Current_Entry.Primary_Type = Internal Then
        
        Current_Internal_Entry.Data_To_Main_Log
        
        Current_Internal_Entry.Data_To_Internal_Log
        
        If Current_Internal_Entry.Internal_Sub_Type = New_Tank Then Call Make_Internal_Carrier_List
    
    End If
    
    Select Case Process_Mode
                
        Case Add_New: MsgBox "Tank" & INSERT_SPACE & Entry_Tank_Number & INSERT_SPACE & "Has Been Added", vbInformation, "Entry Added."
            
        Case Edit: MsgBox "Tank" & INSERT_SPACE & Entry_Tank_Number & INSERT_SPACE & "Has Been Changed", vbInformation, "Entry Edited."

        Case Out: MsgBox "Tank" & INSERT_SPACE & Entry_Tank_Number & INSERT_SPACE & "Has Been Weighed Out", vbInformation, "Entry Weighed Out."
        
    End Select

End Sub
Private Sub Update_Log()

    If Not Process_Mode = Out Then
    
        Call Add_Item_To_Large_List(Selector_Carrier, Carriers, False)
        Call Add_Item_To_Large_List(Sel_Product, Products, False)
    
    End If
    
    Set Current_Entry = Nothing
    
    Unload Me
    
    Call Set_Next_ID
    
End Sub
Private Function Check_If_Controls_Are_Blank() As Boolean

    Dim Validate_Check As Variant
    Dim Validate_In() As String
    Dim Validate_Out() As String
    
    Validate_In = Split("Selector_Carrier" & ";" & "Entry_Tank_Number" & ";" & "Entry_Truck" & ";" & "Entry_In_Weight" & ";" & "Sel_Plant" & ";" & "Entry_Date_In" & ";" & "Entry_Time_In" & ";" & "Selector_Notified" & ";" & "Entry_initials_In" & ";" & "Entry_ID_Prefix", ";")
    Validate_Out = Split("Entry_Date_Out" & ";" & "Entry_Time_Out" & ";" & "Entry_Net_Weight" & ";" & "Entry_initials_Out", ";")
    
    Select Case Process_Mode
            
        Case Add_New:

            Entry_ID_Prefix.BackColor = vbWhite
            
            Entry_Number.BackColor = vbWhite
            
            Selector_Carrier.BackColor = vbWhite
            
            Entry_Tank_Number.BackColor = vbWhite
            
            Entry_Truck.BackColor = vbWhite
            
            Entry_In_Weight.BackColor = vbWhite
            
            Sel_Plant.BackColor = vbWhite
            
            Entry_Date_In.BackColor = vbWhite
            
            Entry_Time_In.BackColor = vbWhite
            
            Selector_Notified.BackColor = vbWhite
            
            Entry_initials_In.BackColor = vbWhite
            
            For Each Validate_Check In Validate_In
            
                If Controls(Validate_Check).value = Empty Then
            
                    MsgBox "Please fill out all required Info", vbInformation, "Removed"
            
                    If Trim(Entry_ID_Prefix.value) = Empty Then Entry_ID_Prefix.BackColor = vbYellow
                        
                    If Trim(Selector_Carrier.value) = Empty Then Selector_Carrier.BackColor = vbYellow
                        
                    If Trim(Entry_Tank_Number.value) = Empty Then Entry_Tank_Number.BackColor = vbYellow
                        
                    If Trim(Entry_Truck.value) = Empty Then Entry_Truck.BackColor = vbYellow
            
                    If Trim(Entry_In_Weight.value) = Empty Then Entry_In_Weight.BackColor = vbYellow
            
                    If Trim(Sel_Plant.value) = Empty Then Sel_Plant.BackColor = vbYellow
            
                    If Trim(Entry_Date_In.value) = Empty Then Entry_Date_In.BackColor = vbYellow
            
                    If Trim(Entry_Time_In.value) = Empty Then Entry_Time_In.BackColor = vbYellow
            
                    If Trim(Selector_Notified.value) = Empty Then Selector_Notified.BackColor = vbYellow
            
                    If Trim(Entry_initials_In.value) = Empty Then Entry_initials_In.BackColor = vbYellow
                    
                    Check_If_Controls_Are_Blank = True
                    
                    Exit Function
                
                End If
            
            Next Validate_Check
                
        Case Out:
            
            For Each Validate_Check In Validate_Out

            Entry_Date_Out.BackColor = vbWhite
        
            Entry_Time_Out.BackColor = vbWhite
        
            Entry_Net_Weight.BackColor = vbWhite
        
            Entry_initials_Out.BackColor = vbWhite
            
            If Controls(Validate_Check).value = Empty Then
            
                MsgBox "Please fill out all required Info", vbInformation, "Removed"
                     
                If Trim(Entry_Date_Out.value) = Empty Then Entry_Date_Out.BackColor = vbYellow
                    
                If Trim(Entry_Time_Out.value) = Empty Then Entry_Time_Out.BackColor = vbYellow
                    
                If Trim(Entry_Net_Weight.value) = Empty Then Entry_Net_Weight.BackColor = vbYellow
                    
                If Trim(Entry_initials_Out.value) = Empty Then Entry_initials_Out.BackColor = vbYellow
                
                Check_If_Controls_Are_Blank = True
                
                Exit Function
            
            End If
            
            Next Validate_Check
            
            Check_If_Controls_Are_Blank = False
                                                                        
    End Select

End Function






