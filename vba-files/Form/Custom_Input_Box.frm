VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Custom_Input_Box 
   ClientHeight    =   1680
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   4564
   OleObjectBlob   =   "Custom_Input_Box.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Custom_Input_Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Confirm_Button_Click()

    Custom_Input_Output = Entry
    Unload Me
    
End Sub
