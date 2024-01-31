VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Alternate_Tables_Box 
   Caption         =   "Reweighs"
   ClientHeight    =   2478
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   10843
   OleObjectBlob   =   "Alternate_Tables_Box.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Alternate_Tables_Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    Call Center_Form(Me)
    
    Select Case Table_Mode
    
        Case Internal_Reweighs:
        
            Call Populate_Reweigh_Internal_Entries_Table
            
            Me.Caption = "Internal Reweighs"
        
        Case Active_Entries:
        
            Call Populate_Actives_Entries_Table
            
            Me.Caption = "Active Entries"
        
    End Select
    
End Sub
Private Sub Populate_Reweigh_Internal_Entries_Table()

    Dim Current_Rows As Long
    Dim Day_Limit As Integer
    
    Day_Limit = Strip_Array(ActiveWorkbook.Names("Option_InternalDayReweighLimit"))
    
    Current_Rows = Count_Entries
    
    For Row_Number = 1 To Current_Rows
    
        If DateDiff("d", CDate([Main_Log[Date In]].Rows(Row_Number)), Date) > Day_Limit And [Main_Log[Status]].Rows(Row_Number) = ENTRY_ACTIVE Then
        
            Select Case Left([Main_Log[ID]].Rows(Row_Number), 1)
            
                Case PREFIX_STORAGE_ID, PREFIX_STORAGE_ID_2, PREFIX_CENTRAL_ID, PREFIX_CENTRAL_ID_2:
                
                    With Me
                    
                        .Entry_Table.AddItem [Main_Log[ID]].Rows(Row_Number)
                        .Entry_Table.List(Entry_Table.ListCount - 1, 1) = [Main_Log[Tank '#]].Rows(Row_Number)
                        .Entry_Table.List(Entry_Table.ListCount - 1, 2) = [Main_Log[RefID]].Rows(Row_Number)
                        
                    End With
                    
            End Select
            
        End If
        
    Next
    
End Sub
Private Sub Populate_Actives_Entries_Table()
    
    For Row_Number = 1 To Count_Entries
    
        If [Main_Log[Status]].Rows(Row_Number) <> ENTRY_INACTIVE Then

            With Me
                    
                .Entry_Table.AddItem [Main_Log[ID]].Rows(Row_Number)
                .Entry_Table.List(Entry_Table.ListCount - 1, 1) = [Main_Log[Tank '#]].Rows(Row_Number)
                .Entry_Table.List(Entry_Table.ListCount - 1, 2) = [Main_Log[RefID]].Rows(Row_Number)
                        
            End With
                    
            
        End If
        
    Next

End Sub
