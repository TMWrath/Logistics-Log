Attribute VB_Name = "Globals_Vars"

    Public Enum Form_Mode
    
        Add_New
        Out
        Edit
        
    End Enum
    
    Public Enum Entry_Typing
    
        External
        Internal
    
    End Enum
    
    Public Enum External_Typing
    
        Live
        Drop
    
    End Enum
    
    Public Enum Internal_Primary_Typing
    
        Storage
        Central
    
    End Enum
    
    Public Enum Internal_Secondary_Typing
    
        New_Tank
        Returning_Tank
        Current_Tank
    
    End Enum
    
    Public Enum Item_Typing
    
        Carriers
        Products
        Internal_Carriers
    
    End Enum
    
    Public Enum Alternate_Table
    
        Internal_Reweighs
        
        Active_Entries
    
    End Enum
    
    Enum Theme_Mode
    
        Black
        Blackout
    
    End Enum
    
    Public Process_Mode As Form_Mode
    Public Table_Mode As Alternate_Table
    Public Current_Theme As Theme_Mode
    Public Next_ID As Integer
    Public Next_Storage_ID As String
    Public Next_Central_ID As String
    Public Next_Drop_ID As String
    Public Plant_List_Type As String
    Public Plant_Product_List As String
    Public Custom_Input_Output As String
    Public Current_List As Name
    Public Current_Max_Entries As Long
    
    Public Const TABLE_STORAGE_TANK As String = ""
    Public Const PARTIAL_TANK_WEIGHT = 19000
    Public Const EMPTY_TANK_WEIGHT = 7700
    Public Const PREFIX_STORAGE_ID As String = "H"
    Public Const PREFIX_STORAGE_ID_2 As String = "I"
    Public Const PREFIX_CENTRAL_ID As String = "C"
    Public Const PREFIX_CENTRAL_ID_2 As String = "F"
    Public Const PREFIX_DROP_ID As String = "D"
    Public Const PREFIX_DROP_ID_2 As String = "T"
    Public Const DEFAULT_DATE_FORMAT As String = "dd-MMM"
    Public Const DEFAULT_TIME_FORMAT As String = "HH:mm"
    Public Const INSERT_SPACE As String = " "
    Public Const ENTRY_INACTIVE As String = "DONE"
    Public Const ENTRY_ACTIVE As String = "IN HOUSE"
    Public Const DEFAULT_DELIMITER As String = "-"
    Public Const IS_EMPTY As String = "Empty"
    Public Const IS_PARTIAL As String = "Partial"
    Public Const IS_FULL As String = "Full"
