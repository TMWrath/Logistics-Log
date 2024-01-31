VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Internal_Log_Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Main_Data As Main_Log_Entry
Private Current_Main_Log_Row As Long
Private Previous_Main_Log_Row As Long
Private Internal_Log_Row As Long
Private ID_Prefix As String
Private ID As String
Private Has_Previous_Entry As Boolean
Public Internal_Type As Internal_Primary_Typing
Public Internal_Sub_Type As Internal_Secondary_Typing
Public Sub Initialize_Main_Data(ByRef Main_Log_Entry_Data As Main_Log_Entry)

    Set Main_Data = Main_Log_Entry_Data
    
    ID = Main_Data.ID
    ID_Prefix = Main_Data.ID_Prefix
    Current_Main_Log_Row = Main_Data.Row_Number
    Internal_Type = Get_Internal_Type
    
End Sub
Public Sub Initialize_Secondary_Data()
    
    Has_Previous_Entry = Get_Has_Previous_Entry
    Internal_Log_Row = Get_Internal_Log_Row
    Previous_Main_Log_Row = Get_Previous_Main_Log_Row
    Internal_Sub_Type = Get_Internal_Sub_Type
    
End Sub
Private Property Get Location() As String

    Location = Main_Data.Plant_Number

End Property
Private Property Get Date_In() As String

    Date_In = Format(Main_Data.Date_In, DEFAULT_DATE_FORMAT)

End Property
Private Property Get Product_Name() As String

    Product_Name = Main_Data.Product_Name

End Property
Private Property Get Carrier_Name() As String

    Carrier_Name = Main_Data.Carrier_Name

End Property
Private Property Get Tank_Number() As String

    Tank_Number = Main_Data.Tank_Number

End Property
Private Property Get Weight() As Long

    Weight = Main_Data.In_Weight

End Property
Private Function Get_Previous_Main_Log_Row() As Long

    Dim Previous_Entry_Reference_ID As String
    Dim Internal_Table_RefID As Object
    
    If Has_Previous_Entry = True Then
    
        Select Case Internal_Type
        
            Case Storage: Set Internal_Table_RefID = [Internal_Log_1[RefID]]
    
            Case Central: Set Internal_Table_RefID = [Internal_Log_2[RefID]]
            
        End Select

        Previous_Entry_Reference_ID = Internal_Table_RefID.Rows(Internal_Log_Row)
            
        Get_Previous_Main_Log_Row = Application.WorksheetFunction.Match(Previous_Entry_Reference_ID, [Main_Log[RefID]], 0)
    
    End If
    
End Function
Private Function Get_Internal_Log_Row() As Long
    
    Dim Internal_Table_Tank As Object

    Select Case Internal_Type
    
        Case Storage: Set Internal_Table_Tank = [Internal_Log_1[Tank '#]]
            
        Case Central: Set Internal_Table_Tank = [Internal_Log_2[Tank '#]]
            
    End Select
    
    If Has_Previous_Entry = True Then
    
        Get_Internal_Log_Row = Application.WorksheetFunction.Match(CLng(Tank_Number), Internal_Table_Tank, 0)
                
    Else
    
        Get_Internal_Log_Row = Application.WorksheetFunction.CountIf(Internal_Table_Tank, "<>") + 1

    End If
    
End Function
Public Function Get_Has_Previous_Entry() As Boolean
    
    Dim Internal_Table_Tank As Object
    Dim Tank_Match As Integer
    
    Select Case Internal_Type
    
        Case Storage: Set Internal_Table_Tank = [Internal_Log_1[Tank '#]]
            
        Case Central: Set Internal_Table_Tank = [Internal_Log_2[Tank '#]]
            
    End Select
    
    On Error Resume Next
    
    Tank_Match = Application.WorksheetFunction.Match(CLng(Tank_Number), Internal_Table_Tank, 0)

    On Error GoTo 0
    
    If Tank_Match > 0 Then Get_Has_Previous_Entry = True Else Get_Has_Previous_Entry = False

End Function
Public Function Get_Internal_Type() As Internal_Primary_Typing

        Select Case ID_Prefix
        
        Case PREFIX_STORAGE_ID, PREFIX_STORAGE_ID_2: Get_Internal_Type = Storage
                    
        Case PREFIX_CENTRAL_ID, PREFIX_CENTRAL_ID_2: Get_Internal_Type = Central
                
    End Select

End Function
Private Function Get_Internal_Sub_Type() As Internal_Secondary_Typing

    Dim Internal_Table_Status As Object

    Select Case Internal_Type
    
        Case Storage: Set Internal_Table_Status = [Internal_Log_1[Status]]

        Case Central: Set Internal_Table_Status = [Internal_Log_2[Status]]
        
    End Select

    Select Case Process_Mode
    
        Case Add_New:
        
            If Has_Previous_Entry = True Then
                
                Select Case Internal_Table_Status.Rows(Internal_Log_Row)
                
                    Case IS_EMPTY, IS_PARTIAL, IS_FULL: Get_Internal_Sub_Type = Current_Tank
                    
                    Case Else: Get_Internal_Sub_Type = Returning_Tank
                    
                End Select
                
            Else
                    
                Get_Internal_Sub_Type = New_Tank
                    
            End If
            
        Case Out: Get_Internal_Sub_Type = Current_Tank
    
    End Select

End Function
Public Sub Set_Previous_Entry_Data()
    
    Dim Internal_Table_ID As Object
    Dim Internal_Table_Carrier As Object
    Dim Internal_Table_Product As Object
    Dim Internal_Table_Date As Object
    Dim Internal_Table_Batch As Object
    Dim Internal_Table_Reference As Object
    
    Select Case Internal_Type
    
        Case Storage:
        
            Set Internal_Table_ID = [Internal_Log_1[ID]]
            Set Internal_Table_Carrier = [Internal_Log_1[Carrier]]
            Set Internal_Table_Product = [Internal_Log_1[Product Name]]
            Set Internal_Table_Date = [Internal_Log_1[Date In]]
            Set Internal_Table_Reference = [Internal_Log_1[Internal-Ref]]
        
        Case Central:
        
            Set Internal_Table_ID = [Internal_Log_2[ID]]
            Set Internal_Table_Carrier = [Internal_Log_2[Carrier]]
            Set Internal_Table_Product = [Internal_Log_2[Product Name]]
            Set Internal_Table_Date = [Internal_Log_2[Date In]]
            Set Internal_Table_Batch = [Internal_Log_2[Batch '#]]
            Set Internal_Table_Reference = [Internal_Log_2[Internal-Ref]]
        
    End Select

    On Error Resume Next
            
    Entry_Box.Sel_Product = Internal_Table_Product.Rows(Internal_Log_Row)
    Entry_Box.Selector_Carrier = Application.Proper(Internal_Table_Carrier.Rows(Internal_Log_Row))
    Entry_Box.Entry_Prev_ID_Date = Internal_Table_ID.Rows(Internal_Log_Row) & INSERT_SPACE & Format(Internal_Table_Date.Rows(Internal_Log_Row), DEFAULT_DATE_FORMAT)
    Entry_Box.Entry_Internal_Ref_ID = Internal_Table_Reference.Rows(Internal_Log_Row)
    Entry_Box.Batch_Num_Entry = Internal_Table_Batch.Rows(Internal_Log_Row)
            
    On Error GoTo 0
    
End Sub
Private Property Get Batch_Number() As String
    
    Do

        Batch_Number = Request_Input("Please Enter A Batch Number For This Entry.", "Enter Information")

    Loop Until Trim(Batch_Number) <> Empty

End Property
Private Property Get Current_Product_Amount() As Long

    If Internal_Sub_Type = Current_Tank Then
    
        Current_Product_Amount = Weight - Empty_Weight
        
    Else
    
        Current_Product_Amount = 0
    
    End If

End Property
Private Property Get Empty_Weight() As Long

    Dim Internal_Table_Empty As Object

    Select Case Internal_Type
    
        Case Storage: Set Internal_Table_Empty = [Internal_Log_1[Empty WT]]
            
        Case Central: Set Internal_Table_Empty = [Internal_Log_2[Empty WT]]
            
    End Select

    If Internal_Sub_Type = Current_Tank Then
    
        Empty_Weight = Internal_Table_Empty.Rows(Internal_Log_Row)
        
    Else
    
    Empty_Weight = Weight
    
    End If

End Property
Private Property Get Last_Net() As Long

    If Internal_Sub_Type = Current_Tank Then
    
        Select Case CLng([Main_Log[Weight]].Rows(Current_Main_Log_Row))
        
            Case Is > CLng([Main_Log[Weight]].Rows(Previous_Main_Log_Row)): Last_Net = CLng([Main_Log[Weight]].Rows(Current_Main_Log_Row)) - CLng([Main_Log[Weight]].Rows(Previous_Main_Log_Row))
        
            Case Is < CLng([Main_Log[Weight]].Rows(Previous_Main_Log_Row)): Last_Net = CLng([Main_Log[Weight]].Rows(Previous_Main_Log_Row)) - CLng([Main_Log[Weight]].Rows(Current_Main_Log_Row))
        
        End Select
        
    Else
    
        Last_Net = Empty
    
    End If

End Property
Private Property Get Weight_Status() As String

    Weight_Status = Get_Weight_Status(Weight)

End Property
Private Property Get Entry_Status() As String

    If Process_Mode = Out Then
    
        If Internal_Status_Sel <> Empty Then Entry_Status = Left(Internal_Status_Sel, 3) Else Entry_Status = Main_Data.Entry_Status
        
    Else
    
        Entry_Status = Main_Data.Entry_Status
    
    End If
    
End Property
Private Property Get Internal_Log_Reference_ID() As String
    
    Select Case Internal_Type
        
        Case Storage: Ref_Prefix = "ST-"
            
        Case Central: Ref_Prefix = "CF-"
        
    End Select
        
    Internal_Log_Reference_ID = Ref_Prefix & DEFAULT_DELIMITER & Carrier_Name & DEFAULT_DELIMITER & Tank_Number
    
End Property
Public Property Get Storage_Data_Array() As Variant
    
    Storage_Data_Array = Array(ID, Carrier_Name, Tank_Number, Product_Name, Date_In, Weight, Last_Net, Weight_Status, Current_Product_Amount, Empty_Weight)
        
End Property
Public Property Get Central_Data_Array() As Variant

    Central_Data_Array = Array(ID, Carrier_Name, Tank_Number, Product_Name, Date_In, Weight, Batch_Number, Location, Weight_Status, Current_Amount, Empty_Weight)

End Property
Public Sub Data_To_Main_Log()

    Select Case Internal_Sub_Type
    
        Case New_Tank: [Main_Log[Int Out]].Rows(Current_Main_Log_Row) = "New"
            
        Case Returning_Tank: [Main_Log[Int Out]].Rows(Current_Main_Log_Row) = "Returned"
        
        Case Current_Tank:
        
            If Process_Mode <> Out Then [Main_Log[Int Out]].Rows(Current_Main_Log_Row) = [Main_Log[ID]].Rows(Previous_Main_Log_Row) & INSERT_SPACE & Format([Main_Log[Date In]].Rows(Previous_Main_Log_Row), DEFAULT_DATE_FORMAT)
            
            If Process_Mode = Out Then
                    
                [Main_Log[Status]].Rows(Current_Main_Log_Row) = Left(Entry_Box.Internal_Status_Sel, 3)
                
                Exit Sub
                
            End If
                        
            If Internal_Sub_Type = Current_Tank Then [Main_Log[Net Weight]].Rows(Previous_Main_Log_Row) = Last_Net
        
            [Main_Log[Date Out]].Rows(Previous_Main_Log_Row) = Format([Main_Log[Date In]].Rows(Current_Main_Log_Row), DEFAULT_DATE_FORMAT)
            
            [Main_Log[Time Out]].Rows(Previous_Main_Log_Row) = Format([Main_Log[Time In]].Rows(Current_Main_Log_Row), DEFAULT_TIME_FORMAT)
            
            If [Main_Log[Int Out]].Rows(Previous_Main_Log_Row) = "NEW" Then [Main_Log[Product Name]].Rows(Previous_Main_Log_Row) = Sel_Product
            
    End Select
    
    If Internal_Sub_Type <> New_Tank Then
    
        [Main_Log[Status]].Rows(Previous_Main_Log_Row) = ENTRY_INACTIVE
        
    End If
    
End Sub
Public Sub Data_To_Internal_Log()

    Dim Current_Table As String
    Dim Column_Count As Integer
    Dim Internal_Table_RefID As Object
    Dim Internal_Data_Array As Variant
    
    Select Case Internal_Type
    
        Case Storage:
        
            Set Internal_Table_RefID = [Internal_Log_1[RefID]]
            Current_Table = "Internal_Log_1[[#Headers],[ID]]"
            Internal_Data_Array = Storage_Data_Array
            Column_Count = 9
        
        Case Central:
        
            Set Internal_Table_RefID = [Internal_Log_2[RefID]]
            Current_Table = "Internal_Log_2[[#Headers],[ID]]"
            Internal_Data_Array = Central_Data_Array
            Column_Count = 10
        
    End Select
    
    For Array_Count = LBound(Internal_Data_Array) To Column_Count

            Range(Current_Table).Offset(Internal_Log_Row, Array_Count) = Internal_Data_Array(Array_Count)

    Next Array_Count
    
    Internal_Table_RefID.Rows(Internal_Log_Row) = Main_Data.Main_Log_Reference_ID
    
End Sub
