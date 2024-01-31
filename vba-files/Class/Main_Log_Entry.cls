VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Main_Log_Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Truck_Number As String
Public Time_In As Date
Public Date_Out As String
Public Time_Out As String
Public Notified As String
Public Initials_In As String
Public Initials_Out As String
Private Main_Log_Target_Row As Long
Private Sub Class_Initialize()

    Select Case Process_Mode
        
        Case Add_New: Main_Log_Target_Row = Count_Entries + 1
        
        Case Else: Main_Log_Target_Row = Application.WorksheetFunction.Match(Choose_Entry_Box.Choose_Entry_Sel, Sheets("Full Log").Range("Main_Log[RefID]"), 0)
        
    End Select
    
End Sub
Public Property Get Row_Number() As Long

    Row_Number = Main_Log_Target_Row

End Property
Public Property Get ID() As String

    Select Case Process_Mode
        
        Case Add_New:
            
            Select Case Left(Entry_Box.Entry_ID_Prefix, 1)
    
                Case PREFIX_CENTRAL_ID, PREFIX_CENTRAL_ID_2, PREFIX_DROP_ID, PREFIX_DROP_ID_2, PREFIX_STORAGE_ID, PREFIX_STORAGE_ID_2: ID = Left(Entry_Box.Entry_ID_Prefix, 1) & Entry_Box.Entry_Number
                            
                Case Else: ID = Entry_Box.Entry_Number
                
            End Select
            
        Case Else: ID = [Main_Log[ID]].Rows(Row_Number)
        
    End Select
    
End Property
Public Property Get ID_Prefix()

    Select Case Left(ID, 1)
      
        Case PREFIX_CENTRAL_ID, PREFIX_CENTRAL_ID_2, PREFIX_DROP_ID, PREFIX_DROP_ID_2, PREFIX_STORAGE_ID, PREFIX_STORAGE_ID_2: ID_Prefix = Left(ID, 1)
            
        Case Else: ID_Prefix = Empty
    
    End Select

End Property
Private Property Get ID_Number()

    Select Case IsNumeric(ID)
      
        Case True: ID_Number = ID
            
        Case False: ID_Number = Replace(ID, ID_Prefix, Empty)
    
    End Select

End Property
Public Property Get Primary_Type() As Entry_Typing

    Select Case ID_Prefix
        
        Case PREFIX_STORAGE_ID, PREFIX_STORAGE_ID_2, PREFIX_CENTRAL_ID, PREFIX_CENTRAL_ID_2: Primary_Type = Internal

        Case Else: Primary_Type = External
                
    End Select
    
End Property
Public Property Get Secondary_Type() As External_Typing

    Select Case ID_Prefix
                    
        Case PREFIX_DROP_ID, PREFIX_DROP_ID_2: Secondary_Type = Drop

        Case Else: Secondary_Type = Live
                
    End Select
    
End Property
Sub Set_In_Data()

    Truck_Number = Entry_Box.Entry_Truck.value
    Time_In = Format(Entry_Box.Entry_Time_In.value, DEFAULT_TIME_FORMAT)
    Notified = Entry_Box.Selector_Notified.value
    Initials_In = Entry_Box.Entry_initials_In.value

End Sub
Public Property Get Plant_Number() As String

    Plant_Number = Entry_Box.Sel_Plant.value

End Property
Public Property Get Date_In() As String

    Date_In = Format(Entry_Box.Entry_Date_In.value, DEFAULT_DATE_FORMAT)

End Property
Sub Set_Out_Data()

    Date_Out = Format(Entry_Box.Entry_Date_Out.value, DEFAULT_DATE_FORMAT)
    Time_Out = Format(Entry_Box.Entry_Time_Out.value, DEFAULT_TIME_FORMAT)
    Initials_Out = Entry_Box.Entry_initials_Out.value

End Sub
Public Property Get Tank_Number() As String

    Tank_Number = Entry_Box.Entry_Tank_Number.value

End Property
Public Property Get Carrier_Name() As String

    Carrier_Name = Entry_Box.Selector_Carrier.value

End Property
Private Property Get User_Input(ByVal Input_Body_Text As String, ByVal Input_Title As String) As String

    Do

        User_Input = Request_Input(Input_Body_Text, Input_Title)

    Loop Until Trim(User_Input) <> Empty

End Property
Public Property Get Product_Name() As String

    Product_Name = Entry_Box.Sel_Product.value

End Property
Public Property Get In_Weight() As Long
 
    In_Weight = Entry_Box.Entry_In_Weight

End Property
Public Property Get Out_Weight() As Long

    Out_Weight = Entry_Box.Entry_Out_Weight

End Property
Public Property Get Entry_Status() As String

    Select Case Process_Mode
    
        Case Add_New:
        
            Entry_Status = ENTRY_ACTIVE
            
        Case Edit:
        
            If Trim(Entry_Net_Weight) <> Empty Then
            
               Entry_Status = ENTRY_ACTIVE
               
            Else
            
                Entry_Status = ENTRY_INACTIVE
            
            End If
        
        Case Out: Entry_Status = ENTRY_INACTIVE
        
    End Select

End Property
Public Property Get Net_Weight() As Long

    If Process_Mode = Out Then
                
        Net_Weight = Entry_Box.Entry_Net_Weight
        
        Exit Property
        
    Else
    
        Net_Weight = Empty
    
    End If

End Property
Public Property Get Main_Log_Reference_ID() As String

    If Process_Mode = Add_New Then
    
        Main_Log_Reference_ID = ID & DEFAULT_DELIMITER & Carrier_Name & DEFAULT_DELIMITER & Tank_Number & DEFAULT_DELIMITER & Row_Number
    
    Else
    
        Main_Log_Reference_ID = [Main_Log[RefID]].Rows(Row_Number)
    
    End If
    
End Property
Public Property Get In_Data_Array() As Variant

    Call Set_In_Data
    
    In_Data_Array = Array(ID, Carrier_Name, Tank_Number, Truck_Number, In_Weight, Product_Name, Plant_Number, Date_In, Time_In, Notified, Initials_In)

End Property
Public Property Get Out_Data_Array() As Variant

    Call Set_Out_Data
    
    Out_Data_Array = Array(Date_Out, Time_Out, Net_Weight, Initials_Out)
    
End Property
Public Sub In_Data_To_Main_Log()

    Dim Array_Count As Integer
    Dim Current_Column As Range
    Dim User_Response As VbMsgBoxResult
    
    For Array_Count = LBound(In_Data_Array) To 8
    
        Set Current_Column = Range("Main_Log[[#Headers],[ID]]").Offset(Row_Number, Array_Count)
        
        Current_Column.value = In_Data_Array(Array_Count)
        
        If In_Data_Array(Array_Count) = In_Weight And Entry_Box.Check_Is_Pounds = True Then
        
            Current_Column.value = Current_Column.value & INSERT_SPACE & "LBS"
        
        End If
    
    Next Array_Count
    
    For Array_Count = 9 To 10
    
        Range("Main_Log[[#Headers],[ID]]").Offset(Row_Number, Array_Count + 2) = In_Data_Array(Array_Count)
    
    Next Array_Count

    If Process_Mode <> Add_New Or Primary_Type = Internal Then
    
        Exit Sub
        
    End If
    
    User_Response = MsgBox("Would you like to add a BOL Number or PO Number?", vbYesNo + vbQuestion, "Add BOL or PO Number?")
            
    If User_Response = vbYes Then
            
        With [Main_Log[Product Name]].Rows(Row_Number)
        
            .AddCommentThreaded (User_Input("Please Enter Any Reference Numbers For This Entry.", "Enter Information"))
            .CommentThreaded.Resolved = True
        
        End With
            
    End If

End Sub
Public Sub Out_Data_To_Main_Log()
    
    Dim Array_Count As Integer
    Dim User_Response As VbMsgBoxResult
    
    For Array_Count = 0 To 1
    
        Range("Main_Log[[#Headers],[ID]]").Offset(Row_Number, Array_Count + 9) = Out_Data_Array(Array_Count)
    
    Next Array_Count
    
    If Entry_Box.Check_Is_Pounds = True Then
    
        [Main_Log[Net Weight]].Rows(Row_Number) = Out_Data_Array(2) & INSERT_SPACE & "LBS"
        
    Else
    
        [Main_Log[Net Weight]].Rows(Row_Number) = Out_Data_Array(2)
    
    End If
    
    [Main_Log[Int Out]].Rows(Row_Number) = Out_Data_Array(3)
    
    If Check_Reject_Entry <> True Then
    
        Exit Sub
    
    End If
    
    User_Response = MsgBox("Would you like to add a reason for the tank being rejected?", vbYesNo + vbQuestion, "Add Rejection Reason?")
                
    If User_Response = vbYes Then
            
        With [Main_Log[Int Out]].Rows(Row_Number)
        
            .AddCommentThreaded (User_Input("Please enter reason this entry is being rejected.", "Enter Information"))
            .CommentThreaded.Resolved = True
        
        End With
            
    End If
        
End Sub
