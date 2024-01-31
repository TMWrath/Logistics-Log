VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Choose_Entry_Box 
   Caption         =   "Choose Entry"
   ClientHeight    =   2282
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   5537
   OleObjectBlob   =   "Choose_Entry_Box.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Choose_Entry_Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    
    Call Theme_Form(Me)
    Call Center_Form(Me)
        
    If Process_Mode = Out Then
    
        Me.Entry_Box_Label.Caption = "Select Tank To Weigh Out:"
        
        Me.Edit_Entry_Sel_Button.Caption = "Weigh Out Tank"
        
    Else
    
        Me.Entry_Box_Label.Caption = "Select Tank Entry To Edit:"
        
        Me.Edit_Entry_Sel_Button.Caption = "Edit Entry"
    
    End If
    
    For Row_Number = 1 To Count_Entries
            
        Select Case Process_Mode
            
            Case Out: If [Main_Log[Status]].Rows(Row_Number) <> ENTRY_INACTIVE Then Choose_Entry_Sel.AddItem [Main_Log[RefID]].Cells(Row_Number)
            
            Case Edit:
            
                Entry_Prefix = Left([Main_Log[ID]].Cells(Row_Number), 1)
            
                Select Case Entry_Prefix ' wtf is this
                
                    Case PREFIX_STORAGE_ID, PREFIX_STORAGE_ID_2, PREFIX_CENTRAL_ID, PREFIX_CENTRAL_ID_2:
                    
                    Case Else: Choose_Entry_Sel.AddItem [Main_Log[RefID]].Rows(Row_Number)
                
                End Select
            
        End Select
        
    Next

End Sub
Private Sub Edit_Entry_Sel_Button_Click()
    
    If Trim(Choose_Entry_Sel) = vbNullString Then
    
        MsgBox "Please Select An Entry", vbInformation, "Removed"
        
        Choose_Entry_Sel.BackColor = vbYellow
    
        Exit Sub
    
    End If
    
    If Process_Mode = Edit Then
    
        Select Case Left(Choose_Entry_Sel, 1)
        
            Case PREFIX_STORAGE_ID, PREFIX_STORAGE_ID_2, PREFIX_CENTRAL_ID, PREFIX_CENTRAL_ID_2:
            
                MsgBox "Internal Entries Cannot be edited", vbInformation, "Entry Editing Not Allowed"
                
                Exit Sub
                
        End Select
    
    End If
    
    Entry_Box.Show vbModeless
    
    Unload Me
    
End Sub
