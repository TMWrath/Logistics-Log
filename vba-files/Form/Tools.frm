VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Tools 
   Caption         =   "Tools"
   ClientHeight    =   3703
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   4046
   OleObjectBlob   =   "Tools.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Tools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    
    Call Theme_Form(Me)
    Call Center_Form(Me)
    
End Sub
Private Sub Rebuild_Log_Button_Click()
    
    Call Rebuild_Log.Make_Pretty
    
End Sub
Private Sub Reform_Formulas_Button_Click()
    
    Call Rebuild_Log.Formulas
    
End Sub
Private Sub Convert_Weight_Button_Click()
    
    Weight_Convert_Box.Show
    
End Sub
Private Sub Decoder_Button_Click()
    
    Entry_Decoder.Show
    
End Sub
Private Sub Employee_Button_Click()
    
    Plant_List_Type = "Employees"
    Plant_List_Box.Show
    
End Sub
Private Sub Internal_Reweighs_Button_Click()
    
    Table_Mode = Internal_Reweighs
    
    Alternate_Tables_Box.Show
    
End Sub
Private Sub Active_Entries_Button_Click()

    Table_Mode = Active_Entries
    
    Alternate_Tables_Box.Show

End Sub
Private Sub Product_Button_Click()
    
    Plant_List_Type = "Products"
    Plant_List_Box.Show
    
End Sub
Private Sub Updater_Button_Click()
    
    Updater_Box.Show
    
End Sub
Private Sub Weigh_In_Pounds_Button_Click()
    
    Weigh_In_Pounds.Show
    
End Sub
