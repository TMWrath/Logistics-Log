Attribute VB_Name = "Global_Subs"
Public Sub Set_Next_ID()

    Dim Last_Row As Long
    Dim Row_Number As Long
    Dim Drop_ID As String
    Dim Drop_ID_2 As String
    Dim Storage_ID As String
    Dim Storage_ID_2 As String
    Dim Central_ID As String
    Dim Central_ID2 As String
    Dim Current_Row_Count As Long
    Dim Row_Status As String
    Dim Live_ID_Column As Range
    Dim Drop_ID_Column As Range
    Dim Storage_ID_Column As Range
    Dim Central_ID_Column As Range
    Const TABLE_ROW_START As Byte = 1
    
    Last_Row = Count_Entries
    Next_ID = [Main_Log[ID]].Rows(Last_Row + 1)

    Storage_ID = PREFIX_STORAGE_ID & Next_ID
    Storage_ID_2 = PREFIX_STORAGE_ID_2 & Next_ID
    Central_ID = PREFIX_CENTRAL_ID & Next_ID
    Central_ID_2 = PREFIX_CENTRAL_ID_2 & Next_ID
    Drop_ID = PREFIX_DROP_ID & Next_ID
    Drop_ID_2 = PREFIX_DROP_ID_2 & Next_ID
    
    Set Live_ID_Column = [Table_Next_ID].Columns(2)
    Set Storage_ID_Column = [Table_Next_ID].Columns(4)
    Set Central_ID_Column = [Table_Next_ID].Columns(6)
    Set Drop_ID_Column = [Table_Next_ID].Columns(8)
    
    Live_ID_Column = Next_ID
    Storage_ID_Column = Storage_ID
    Central_ID_Column = Central_ID
    Drop_ID_Column = Drop_ID
    
    If Last_Row <> 0 Then
        
        For Row_Number = TABLE_ROW_START To Last_Row
        
            Row_Status = [Main_Log[Status]].Rows(Row_Number)
            
            If Row_Status <> ENTRY_INACTIVE Then
            
                Select Case Left([Main_Log[ID]].Rows(Row_Number), 1)
                
                    Case PREFIX_STORAGE_ID, PREFIX_STORAGE_ID_2:
                    
                        If [Main_Log[ID]].Rows(Row_Number) = Storage_ID Then
                        
                            Storage_ID_Column = Storage_ID_2
                            
                        Else
                        
                            Storage_ID_Column = Storage_ID
                            
                        End If
                    
                    Case PREFIX_CENTRAL_ID, PREFIX_CENTRAL_ID_2:
                    
                        If [Main_Log[ID]].Rows(Row_Number) = Central_ID Then
                        
                            Central_ID_Column = Central_ID_2
                            
                        Else
                            
                            Central_ID_Column = Central_ID
                            
                        End If
                        
                    Case PREFIX_DROP_ID, PREFIX_DROP_ID_2:
                    
                        If [Main_Log[ID]].Rows(Row_Number) = Drop_ID Then
                        
                            Drop_ID_Column = Drop_ID_2
                            
                        Else
                            
                            Drop_ID_Column = Drop_ID
                            
                        End If
                
                End Select
                
            End If
            
        Next Row_Number
    
    End If

    Next_Storage_ID = Storage_ID_Column
    Next_Central_ID = Central_ID_Column
    Next_Drop_ID = Drop_ID_Column

End Sub
Public Sub Center_Form(ByVal Current_Form As Object)

    With Current_Form
    
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
                    
    End With

End Sub
Public Sub Set_Control_From_Small_Data_List(ByVal ListName As String, ByVal List_Controller As Object)

    Dim Array_String As String
    Dim Temporay_List() As String
    Dim List_Item As Integer

    Set Current_List = ActiveWorkbook.Names(ListName)
    
    Array_String = Strip_Array(Current_List)
    
    Temporay_List = Split(Array_String, ",")
    
    For List_Item = LBound(Temporay_List) To UBound(Temporay_List)
        
            List_Controller.AddItem Temporay_List(List_Item)
        
    Next List_Item

End Sub
Public Sub Make_List(ByVal Current_List_Name As String, ByVal Current_Control As Object, ByVal Current_Control_Count As Long)

    Dim List_Item As String
    Dim Control_Index As Integer
    
    If Current_Control_Count <> 0 Then
    
        For Control_Index = 0 To Current_Control_Count - 1
            
            If List_Item = vbNullString Then
            
                List_Item = """" & Current_Controller.List(Control_Index) & """"
                
            Else
    
                List_Item = List_Item & "," & """" & Current_Control.List(Control_Index) & """"
    
            End If
       
        Next
        
    Else
    
        List_Item = vbNullString
    
    End If
    
    If List_Item = vbNullString Then
    
        ActiveWorkbook.Names.Add Name:=Current_List_Name, RefersToR1C1:="="""""
    
    Else
    
        ActiveWorkbook.Names.Add Name:=Current_List_Name, RefersToR1C1:=List_Item
    
    End If

End Sub
Public Sub Set_Option_Small_List(ByVal Option_Name As String, ByVal Option_Value As String)

    ActiveWorkbook.Names.Add Name:=Option_Name, RefersToR1C1:=Option_Value

End Sub
Public Sub Set_Control_From_Large_Data_List(ByVal Item_Type As Item_Typing, ByVal List_Controller As Object)

    Dim cell As Range
    Dim Item_Table As Object
    
    Select Case Item_Type
    
        Case Carriers:
        
            Set Item_Table = [Database_Carriers[List]]
            
        Case Products:
        
            Set Item_Table = [Database_Products[List]]
            
        Case Internal_Carriers:
        
            Set Item_Table = [Database_Internal_Carriers[List]]
            
    End Select
    
    For Each cell In Item_Table
            
            List_Controller.AddItem Application.Proper(cell.value)
                
    Next cell

End Sub
Public Sub Remove_Data_List_From_Control(ByVal List_Controller As Object)
    
    If List_Controller.ListCount <> 0 Then
    
        For Item_Count = List_Controller.ListCount - 1 To 0 Step -1
                
                List_Controller.RemoveItem Item_Count
                    
        Next Item_Count
    
    End If

End Sub
Public Sub Add_Item_To_Large_List(ByVal NewItem As String, ByVal Item_Type As Item_Typing, ByVal Show_Message As Boolean)

    Dim To_List As String
    Dim To_Table As Object
    Dim Item_Match As Integer
    Dim TargetListRow As Integer
    
    Select Case Item_Type
    
        Case Carriers:
        
            To_List = "Database_Carriers"
            
            Set To_Table = [Database_Carriers[List]]
            
        Case Products:
        
            To_List = "Database_Products"
            
            Set To_Table = [Database_Products[List]]
            
    End Select
    
    On Error Resume Next
    
    Item_Match = Application.WorksheetFunction.Match(Application.Proper(NewItem), To_Table, 0)
    
    On Error GoTo 0

    If Item_Match = 0 Then
            
        TargetListRow = Application.WorksheetFunction.CountIf(To_Table, "<>") + 1

        To_Table.Rows(TargetListRow) = Application.Proper(NewItem)
        
        If Show_Message = True Then MsgBox Application.Proper(NewItem) & " Added", vbInformation, "Added"
        
    Else
    
        If Show_Message = True Then MsgBox Application.Proper(NewItem) & " Is already In The List", vbInformation, "Already Exists"
            
    End If
    
End Sub
Public Sub Add_Item_To_Small_List(ByVal List_Name As String, ByVal List_Controller As Object, ByVal Item_Type As String, ByVal Show_Message As Boolean)
    
    Dim Item_Match As Integer
    Dim Item_String As String
    
    Set Current_List = ActiveWorkbook.Names(List_Name)
    
    Item_String = Strip_Array(Current_List)
    
    Item_Match = InStr(1, Item_String, List_Controller, vbTextCompare)
    
    If Item_Match = 0 Then
    
        List_Controller.AddItem List_Controller.value
        
        Call Make_List(List_Name, List_Controller, List_Controller.ListCount)
        
        
        If Show_Message = True Then MsgBox Item_Type & " Added", vbInformation, "Added"
        
        List_Controller.value = vbNullString
    
    End If

End Sub
Public Sub Remove_Item_From_Small_List(ByVal List_Name As String, ByVal List_Controller As Object, ByVal Item_Type As String, ByVal Show_Message As Boolean)
    
    Dim Current_Box_Count As Integer
    Dim List_Index As Integer
    
    If List_Controller.value = vbNullString Then
    
        Exit Sub
        
    End If

    Current_Box_Count = List_Controller.ListCount
    
    For List_Index = 0 To Current_Box_Count - 1
    
        If List_Controller.List(List_Index) = List_Controller.value Then
            
            List_Controller.RemoveItem List_Index
            
            Exit For
            
        End If
        
    Next List_Index
    
    Current_Box_Count = List_Controller.ListCount
    
    Call Make_List(List_Name, List_Controller, List_Controller.ListCount)
    
    If Show_Message = True Then MsgBox Item_Type & " Removed", vbInformation, "Removed"
    
    List_Controller.value = vbNullString

End Sub
Public Sub Remove_Item_From_Large_List(ByVal NewItem As String, ByVal Item_Type As Item_Typing, ByVal Show_Message As Boolean)

    Dim To_List As String
    Dim To_Table As Object
    Dim Item_Match As Integer
    Dim Target_List_Row As Integer
    
    Select Case Item_Type
    
        Case Carriers:
        
            To_List = "Database_Carriers"
            
            Set To_Table = [Database_Carriers[List]]
            
        Case Products:
        
            To_List = "Database_Products"
            
            Set To_Table = [Database_Products[List]]
            
    End Select
    
    On Error Resume Next
    
    Item_Match = Application.WorksheetFunction.Match(Application.Proper(NewItem), To_Table, 0)
    
    On Error GoTo 0
    
    If Item_Match > 0 Then
        
        Target_List_Row = Item_Match
        
        To_Table.Rows(Target_List_Row) = vbNullString
        
        If Show_Message = True Then MsgBox Application.Proper(NewItem) & " Removed", vbInformation, "Removed"

    Else
    
        If Show_Message = True Then MsgBox Application.Proper(NewItem) & " Is Not In The List", vbInformation, "Doesn't Exists"
    
    End If
    
End Sub
Public Sub Sort_Item_Data_List(ByVal Item_Type As Item_Typing)

    Dim To_List As String
    Dim To_Table As String
    
    Select Case Item_Type
    
        Case Carriers:
        
            To_List = "Database_Carriers"
            
            To_Table = "Database_Carriers[[#All],[List]]"
            
        Case Products:
        
            To_List = "Database_Products"
            
            To_Table = "Database_Products[[#All],[List]]"
            
        Case Internal_Carriers:
        
            To_List = "Database_Internal_Carriers"
            
            To_Table = "Database_Internal_Carriers[[#All],[List]]"
            
    End Select
    
    ActiveWorkbook.Worksheets("Database").ListObjects(To_List).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Database").ListObjects(To_List).Sort. _
    SortFields.Add2 Key:=Range(To_Table), SortOn:= _
    xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    With ActiveWorkbook.Worksheets("Database").ListObjects(To_List).Sort
    
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        
    End With

End Sub
Public Sub Generate_Data_List_From_Log(ByVal Item_Type As Item_Typing)

    Dim Item_Array() As Variant
    Dim iCount As Long
    Dim cell As Range
    Dim List_Item As String
    Dim From_Table As Object
    Dim To_Table As Object
    
    Select Case Item_Type
    
        Case Carriers:
        
            Set From_Table = [Main_Log[Carrier]]
            Set To_Table = [Database_Carriers[List]]
            Sheets("Database").ListObjects("Database_Carriers").ListColumns("List").DataBodyRange.ClearContents
            
        Case Products:
        
            Set From_Table = [Main_Log[Product Name]]
            Set To_Table = [Database_Products[List]]
            Sheets("Database").ListObjects("Database_Products").ListColumns("List").DataBodyRange.ClearContents
            
    End Select
    
    For Each cell In From_Table
            
        ReDim Preserve Item_Array(0 To iCount)
                
        List_Item = Join(Item_Array, ",")
                
        If InStr(1, List_Item, cell.value, vbTextCompare) = 0 Then
            
            Item_Array(iCount) = cell.value
            
            To_Table.Rows(iCount + 1) = Application.Proper(cell.value)
            
            iCount = iCount + 1
                
        End If
                
    Next cell

    MsgBox "List Has been Created", vbInformation, "Removed"

End Sub
Public Sub Reset_Row(ByVal Row_Number As Long)

    Dim Column_Number As Integer
    
    Sheets("Full Log").ListObjects("Main_Log").ListRows(Row_Number).Range.ClearContents

    Call Reset_ID(Row_Number)
    
    Call Main_Log_Formulas(Row_Number)
   
End Sub
Public Sub Reset_ID(ByVal Target_Row As Long)
    
    Dim Previous_Entry_ID As String
    Dim Previous_Entry As Range
    
    Set Previous_Entry = [Main_Log[ID]].Rows(Target_Row - 1)
    
    Const ID_START = 1
    Const ID_END = 99
    
    If Previous_Entry.value = "ID" Then
    
        [Main_Log[ID]].Rows(Target_Row) = ID_START
        
        Exit Sub
    
    End If
    
    Select Case IsNumeric(Previous_Entry_ID)
    
        Case False: Previous_Entry_ID = Replace(Previous_Entry_ID, Left(Previous_Entry_ID, 1), "")
        
        Case Else: Previous_Entry_ID = Previous_Entry.value
    
    End Select
    
    Select Case Previous_Entry_ID
            
            Case ID_END: [Main_Log[ID]].Rows(Target_Row) = ID_START
                
            Case Else: [Main_Log[ID]].Rows(Target_Row) = CInt(Previous_Entry_ID) + 1
            
    End Select

End Sub
Public Sub Main_Log_Formulas(ByVal Target_Row As Long)
   
    Select Case [Main_Log[Status]].Rows(Target_Row)
       
        Case "OFR", "OTC", "NRP", "OFR": [Main_Log[Status]].Rows(Target_Row) = [Main_Log[Status]].Rows(Target_Row)
       
        Case Else: [Main_Log[Status]].Rows(Target_Row) = "=IF(OR([@Carrier] = """", [@[Tank '#]] = """",[@[Truck '#]] = """",[@Weight] = """",[@[PLT '#]] = """",[@[Date In]] = """",[@[Time In]] ="""",[@Notified] = """",[@[Int In]] = """"), """",IF(OR([@[Date Out]] = """",[@[Time Out]] = """",[@[Net Weight]] = """",[@[Int Out]] = """"), ""IN HOUSE"",""DONE""))"
       
    End Select
   
    [Main_Log[RefID]].Rows(Target_Row) = "=IF(ISBLANK([@Carrier]),"""",CONCAT([@ID],""-"",[@Carrier],""-"",[@[Tank '#]],""-"",ROW()-2))"
    [Main_Log[FS]].Rows(Target_Row) = "=IF(OR(ISNUMBER(SEARCH(""C"",[@ID])), ISNUMBER(SEARCH(""F"",[@ID])), ISNUMBER(SEARCH(""I"",[@ID])), ISNUMBER(SEARCH(""H"",[@ID]))), IF([@Weight]>16000,""F"",IF([@Weight] > 7000,""P"",""E"")), """")"

End Sub
Public Sub Make_Internal_Carrier_List()

    Dim Storage_Array() As Variant
    Dim Central_Array() As Variant
    Dim Combined_Arrays As String
    Dim Internal_Carriers As Variant
    Dim iCount As Long
    Dim cell As Range
    Dim Search_String As String
    Dim List_Item As Long
    Dim Table_Row As Long
    
    [Database_Internal_Carriers[List]].Clear
    
    For Each cell In [Internal_Log_1[Carrier]]
        
        ReDim Preserve Storage_Array(0 To iCount)
                
        Search_String = Join(Storage_Array, ",")
                
        If InStr(1, Search_String, cell.value, vbTextCompare) = 0 Then
                
            Storage_Array(iCount) = Application.Proper(cell.value)
                    
            iCount = iCount + 1
                
        End If
                
    Next cell
    
    iCount = 0
        
    For Each cell In [Internal_Log_2[Carrier]]
            
        ReDim Preserve Central_Array(0 To iCount)
                
        Search_String = Join(Central_Array, ",")
                
        If InStr(1, Search_String, cell.value, vbTextCompare) = 0 Then
                
            Central_Array(iCount) = Application.Proper(cell.value)
                    
            iCount = iCount + 1
                
        End If
                
    Next cell
    

    Combined_Arrays = Join(Storage_Array, ",") & Join(Central_Array, ",")
    Internal_Carriers = Split(Combined_Arrays, ",")
    
    For List_Item = LBound(Internal_Carriers) To UBound(Internal_Carriers)
    
        On Error Resume Next
        
        If Application.WorksheetFunction.Match(Internal_Carriers(List_Item), Sheets("Database").Range("Database_Internal_Carriers[List]"), 0) = 0 Or Internal_Carriers(List_Item) = vbNullString Then
        
            Table_Row = Table_Row + 1
            
            [Database_Internal_Carriers[List]].Rows(Table_Row) = Internal_Carriers(List_Item)
        
        End If
        
        On Error GoTo 0
        
    Next
    
End Sub
