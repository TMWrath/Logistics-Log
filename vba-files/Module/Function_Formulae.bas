Attribute VB_Name = "Function_Formulae"
Public Function Main_Log_Status(ByVal Current_Carrier As String, ByVal Current_Net As String) As String

    If Current_Carrier = Empty Then
    
        Main_Log_Status = Empty
        
        Exit Function
        
    End If

    If Current_Net <> Empty Then
    
        Main_Log_Status = "DONE"
        
        Exit Function
        
    End If

    Main_Log_Status = "IN HOUSE"

    '=IF(OR([@Carrier] = "", [@[Tank '#]] = "",[@[Truck '#]] = "",[@Weight] = "",[@[PLT '#]] = "",[@[Date In]] = "",[@[Time In]] ="",[@Notified] = "",[@[Int In]] = ""), "",IF(OR([@[Date Out]] = "",[@[Time Out]] = "",[@[Net Weight]] = "",[@[Int Out]] = ""), "IN HOUSE","DONE"))

End Function
Public Function Main_Log_Imternal_Fill_Status(ByVal Current_ID As String, ByVal Current_Weight) As String
    
    Select Case Left(Current_ID, 1)
            
        Case PREFIX_STORAGE_ID, PREFIX_STORAGE_ID_2, PREFIX_CENTRAL_ID, PREFIX_CENTRAL_ID_2: Main_Log_Imternal_Fill_Status = Left(Get_Weight_Status(Current_Weight), 1)
        
        Case Else: Main_Log_Imternal_Fill_Status = ""
            
    End Select
        
'=IF(OR(ISNUMBER(SEARCH("C",[@ID])), ISNUMBER(SEARCH("F",[@ID])), ISNUMBER(SEARCH("I",[@ID])), ISNUMBER(SEARCH("H",[@ID]))), IF([@Weight]>16000,"F",IF([@Weight] > 7000,"P","E")), "")

End Function
Public Function Main_Log_Ref(ByVal Current_Carrier As String) As String
    
    Dim Current_Tank_Number As String
    Dim Current_Row As Long

    Current_Row = Application.Caller.Row - 2
    Current_ID = [Main_Log[ID]].Rows(Current_Row)
    Current_Tank_Number = [Main_Log[Tank '#]].Rows(Current_Row)

    If Current_Carrier = Empty Then
    
        Main_Log_Ref = Empty
        
        Exit Function
        
    End If

    Main_Log_Ref = Current_ID & DEFAULT_DELIMITER & Current_Carrier & DEFAULT_DELIMITER & Current_Tank_Number & DEFAULT_DELIMITER & Current_Row

End Function
Public Function InternalRef(ByVal Current_ID As String) As String
    
    Dim Current_Table As String
    Dim Current_Row As Long
    Dim Current_Carrier As String
    Dim Current_Prefix As String
    Dim Current_Tank_Number As String
    Dim Current_Range As Range
    
    Current_Table = Application.Caller.ListObject.Name
    Current_Row = Application.Caller.Row - 2
    
    Select Case Current_Table
            
        Case "Internal_Log_1":
                
            Current_Prefix = "ST-"
            Set Current_Range = Range("Internal_Log_1[[#Headers],[ID]]")
                
        Case "Internal_Log_2":
            
            Current_Prefix = "CF-"
            Set Current_Range = Range("Internal_Log_2[[#Headers],[ID]]")
            
    End Select
        
    Current_Carrier = Current_Range.Offset(Current_Row, 1)
    Current_Tank_Number = Current_Range.Offset(Current_Row, 2)
        
    If Current_ID = Empty Then
        
        InternalRef = Empty
            
        Exit Function
            
    End If
        
    InternalRef = Current_Prefix & DEFAULT_DELIMITER & Current_Carrier & DEFAULT_DELIMITER & Current_Tank_Number & DEFAULT_DELIMITER & Current_Row

End Function
Public Function Internal_Main_Log_RefID(ByVal Current_ID As String) As String
    
    Dim Current_Tank As Long
    Dim Current_Tank_Number As String
    Dim Current_Table As String
    Dim Current_Entry_Status As String
    Dim Row_Number As Long
    
    Current_Table = Application.Caller.ListObject.Name
    Current_Row = Application.Caller.Row - 2
    
    Select Case Current_Table
        
        Case "Internal_Log_1": Set Current_Range = Range("Internal_Log_1[[#Headers],[ID]]")
            
        Case "Internal_Log_2": Set Current_Range = Range("Internal_Log_2[[#Headers],[ID]]")
        
    End Select
    
    Current_Tank_Number = Current_Range.Offset(Current_Row, 2)
    
    If Current_ID = Empty Then

        Internal_Main_Log_RefID = Empty
        
        Exit Function
        
    End If

    For Row_Number = Count_Entries To 1 Step -1
        
        Current_Main_Log_ID = Left([Main_Log[ID]].Rows(Row_Number), 1)
        
        If Current_Entry_Status <> ENTRY_ACTIVE Then
        
            GoTo ContinueLoop
            
        End If
        
        Select Case Current_Main_Log_ID
            
            Case PREFIX_STORAGE_ID, PREFIX_STORAGE_ID_2, PREFIX_CENTRAL_ID, PREFIX_CENTRAL_ID_2:
                
                Current_Entry_Status = [Main_Log[Status]].Rows(Row_Number)
            
                If [Main_Log[Tank '#]].Rows(Row_Number) = Current_Tank_Number Then
                            
                    Internal_Main_Log_RefID = [Main_Log[RefID]].Rows(Row_Number)
                            
                End If
            
        End Select
        
ContinueLoop:

    Next
        
End Function
