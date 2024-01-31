Attribute VB_Name = "Globals_Functions"
Function Strip_Array(ByVal Raw_Array As Variant) As String

    Dim Array_String As String
    
    Array_String = Replace(Raw_Array, "=", "")
    Array_String = Replace(Array_String, "{", "")
    Array_String = Replace(Array_String, "}", "")
    Array_String = Replace(Array_String, """", "")
    Strip_Array = Array_String
    
End Function
Function Check_Plant_For_Product_List(ByVal Plant_Number As String) As Boolean

    Dim Plant_Product_List_Name As String
    Dim Current_Name As Name
    
    Plant_Product_List_Name = "List_Plant_" & Plant_Number & "_Products"
        
    For Each Current_Name In ActiveWorkbook.Names
        
        If Current_Name.Name = Plant_Product_List_Name Then
            
                Check_Plant_For_Product_List = True
                
                Exit Function
                
        End If
    
    Next Current_Name

    Check_Plant_For_Product_List = False
    
End Function
Function Request_Input(ByVal Body_Text As String, ByVal Title As String) As String

    Custom_Input_Box.Body_Text = Body_Text
    Custom_Input_Box.Caption = Title
    Custom_Input_Box.Show
    Request_Input = Custom_Input_Output
    Custom_Input_Output = Empty

End Function
Function Count_Entries() As Long

    Count_Entries = Application.WorksheetFunction.CountIf([Main_Log[Carrier]], "<>")

End Function
Public Function Get_Weight_Status(ByVal Tank_Weight As Long) As String

    Select Case Tank_Weight
                                
        Case Is < EMPTY_TANK_WEIGHT: Get_Weight_Status = IS_EMPTY
                                    
        Case EMPTY_TANK_WEIGHT + 1 To PARTIAL_TANK_WEIGHT: Get_Weight_Status = IS_PARTIAL
                                    
        Case Is > PARTIAL_TANK_WEIGHT: Get_Weight_Status = IS_FULL
        
        Case Else: Get_Weight_Status = ""
        
    End Select

End Function


