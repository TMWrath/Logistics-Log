VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Plant_List_Box 
   Caption         =   "Product Locations"
   ClientHeight    =   5838
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   4060
   OleObjectBlob   =   "Plant_List_Box.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Plant_List_Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewEmployeeList As String
Dim Plant_Employee_List As String
Private Sub UserForm_Initialize()
    
    Me.Height = 200
    
    Call Theme_Form(Me)
    Call Center_Form(Me)
    
    Set Current_List = ActiveWorkbook.Names("List_Plants")
                
    Plant_List = Split(Strip_Array(Current_List), ",")
    
    Select Case Plant_List_Type
    
        Case "Products":
        
            Me.Plant_Box.ColumnCount = 2
            
            For List_Index = LBound(Plant_List) To UBound(Plant_List)
            
                If Plant_List(List_Index) <> vbNullString Then
                
                    If Check_Plant_For_Product_List(Plant_List(List_Index)) = True Then
                    
                        Me.Sel_Plant.AddItem Plant_List(List_Index)
                    
                    End If
                
                End If
                
            Next List_Index
            
            Me.List_Item_lbl.Caption = "Products"
            Me.Item_Status_lbl.Caption = "Product Status"
            
            Call Set_Control_From_Large_Data_List(Products, Me.Plant_Item_Sel)
            Me.Plant_Item_Status_Sel.AddItem "In"
            Me.Plant_Item_Status_Sel.AddItem "Out"
            
        Case "Employees":
    
            Me.Plant_Box.ColumnCount = 1
            
            For List_Index = LBound(Plant_List) To UBound(Plant_List)
            
                If Plant_List(List_Index) <> vbNullString Then
                
                    If ChkPlantEmp(Plant_List(List_Index)) = True Then
                    
                        Me.Sel_Plant.AddItem Plant_List(List_Index)
                    
                    End If
                
                End If
                
            Next List_Index
            
            Me.List_Item_lbl.Caption = "Employees"
            Me.Item_Status_lbl.Caption = vbNullString
            Me.Plant_Item_Status_Sel.Visible = False
            Me.Plant_Item_Sel.Style = fmStyleDropDownCombo
            
    End Select
    
End Sub
Private Sub Add_Remove_Allow_Button_Click()

    Plant_List_Box.Height = 316.75

End Sub
Private Sub Sel_Plant_Change()

    Dim List_Data() As String
    Dim Item_Index As Integer
    
    Me.Plant_Box.Clear
    Me.Plant_Item_Sel = vbNullString
    
    If Sel_Plant <> vbNullString Then
    
        Select Case Plant_List_Type
        
            Case "Products":
                
                Plant_Product_List = "List_Plant_" & Sel_Plant & "_Products"
                
                List_Data = Split(Strip_Array(ActiveWorkbook.Names(Plant_Product_List)), ",")
                    
                For Item_Index = 0 To UBound(List_Data)
                            
                    Me.Plant_Box.AddItem Replace(List_Data(Item_Index), """", "")
                            
                    List_Item = Item_Index + 1
                            
                    Me.Plant_Box.List(Plant_Box.ListCount - 1, 1) = Replace(List_Data(Item_Index), """", "")
                                
                Next Item_Index
            
            Case "Employees":
            
                PlantEmployeeList = "List_Plant_" & Sel_Plant & "_Employees"
                
                List_Data = Split(Strip_Array(ActiveWorkbook.Names(PlantEmployeeList)), ",")
                    
                For Item_Index = 0 To UBound(List_Data)
                            
                    Me.Plant_Box.AddItem Replace(List_Data(Item_Index), """", "")
                    Me.Plant_Item_Sel.AddItem Replace(List_Data(Item_Index), """", "")
                                
                Next Item_Index
                
        End Select
        
    End If
    
End Sub
Private Sub Add_Plant_Item_Button_Click()

    Dim New_Row As Long
    
    If Trim(Me.Plant_Item_Sel) = vbNullString Then
    
        If Trim(Me.Plant_Item_Sel) = vbNullString Then Me.Plant_Item_Sel.BackColor = vbYellow
    
        Exit Sub
    
    End If
    
    If Plant_List_Type = "Products" Then
    
        If Plant_Item_Status_Sel = vbNullString Then
        
            MsgBox "Please Fill In All Required Information", vbInformation, "Removed"
            
            If Trim(Me.Plant_Item_Status_Sel) = vbNullString Then Me.Plant_Item_Status_Sel.BackColor = vbYellow
            
            Exit Sub
        
        End If
        
    End If
    
    New_Row = Plant_Box.ListCount
    Me.Plant_Box.AddItem Me.Plant_Item_Sel.value
    
    Select Case Plant_List_Type
    
        Case "Products":
        
            Me.Plant_Box.List(New_Row, 1) = Me.Plant_Item_Status_Sel.value
            
            Call Create_Plant_Product_List
            
            Call Add_Item_To_Large_List(Me.Plant_Item_Sel, Products, True)
            
            MsgBox "Product Added", vbInformation, "Added"
                
            Call Reset_Form
            
        Case "Employees":
        
            Plant_Employee_List = "List_Plant_" & Me.Sel_Plant.value & "_Employees"
                    
            Make_List Plant_Employee_List, Me.Plant_Box, Me.Plant_Box.ListCount
                
            MsgBox "Employee Added", vbInformation, "Added"
                
            Call Reset_Form
                
        End Select

End Sub
Private Sub Remove_PLant_Item_Button_Click()

    If Me.Plant_Box.ListIndex > -1 Then
    
        Me.Plant_Box.RemoveItem Me.Plant_Box.ListIndex
    
        Select Case Plant_List_Type
    
            Case "Products":
            
                Call Create_Plant_Product_List
                
                MsgBox "Product Removed", vbInformation, "Removed"
                
                Call Reset_Form
            
            Case "Employees":

                Plant_Employee_List = "List_Plant_" & Me.Sel_Plant.value & "_Employees"
            
                Make_List Plant_Employee_List, Me.Plant_Box, Me.Plant_Box.ListCount
                
                MsgBox "Employee Removed", vbInformation, "Removed"
                
                Call Reset_Form
            
        End Select
        
    Else
    
        MsgBox "Please Select Product To Remove", vbInformation, "Removed"
    
    End If

End Sub
Private Sub Create_Plant_Product_List()

    Dim New_Product_List As String
    

    For i = 0 To Plant_Box.ListCount - 1
                    
        If New_Product_List = vbNullString Then
                        
            New_Product_List = """" & Plant_Box.List(i) & """"
                            
        Else
                        
            New_Product_List = New_Product_List & "," & """" & Plant_Box.List(i) & """"
                            
        End If
                    
        New_Product_List = New_Product_List & "," & """" & Plant_Box.List(i, 1) & """"
                        
    Next
    
    Plant_Product_List = "List_Plant_" & Me.Sel_Plant.value & "_Products"
    
    If New_Product_List = vbNullString Then
                
        ActiveWorkbook.Names.Add Name:=Plant_Product_List, RefersToR1C1:="="""""
                
    Else
                
        ActiveWorkbook.Names.Add Name:=Plant_Product_List, RefersToR1C1:=New_Product_List
                
    End If

End Sub
Private Sub Reset_Form()

    Me.Sel_Plant = vbNullString
    Me.Plant_Item_Sel = vbNullString
    Me.Plant_Item_Status_Sel = vbNullString

End Sub
Private Function ChkPlantEmp(ByVal PlantName As String) As Boolean

    Dim PlantEmployeeListName As String

    PlantEmployeeListName = "List_Plant_" & PlantName & "_Employees"
        
    For Each Nm In ActiveWorkbook.Names
        
        If Nm.Name = PlantEmployeeListName Then
            
                ChkPlantEmp = True
                
                Exit Function
                
        End If
    
    Next Nm

    ChkPlantEmp = False
    
End Function
