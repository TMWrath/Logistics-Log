VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Weight_Convert_Box 
   Caption         =   "Weight Conversion"
   ClientHeight    =   1953
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   4564
   OleObjectBlob   =   "Weight_Convert_Box.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Weight_Convert_Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    Call Theme_Form(Me)
    Call Center_Form(Me)

    Me.Conversion_Sel.AddItem "Lbs to Kgs"
    Me.Conversion_Sel.AddItem "Kgs to Lbs"

End Sub
Private Sub Convert_Button_Click()

    Dim Weight_To_LBS As Long
    Dim Weight_To_Kgs As Long
    Dim Weight_Value As Long
    
    Weight_Value = Me.Weight_Entry.value
    
    Weight_To_LBS = Weight_Value * 2.20462262
    Weight_To_Kgs = Weight_Value / 2.20462262

    Me.Conversion_Sel.BackColor = vbWhite
    
    If Conversion_Sel <> vbNullString Then
    
        Select Case Me.Conversion_Sel
        
            Case "Lbs to Kgs": MsgBox WorksheetFunction.Round(Weight_To_Kgs, -1), vbInformation, "Converted!"
            
            Case "Kgs to Lbs": MsgBox WorksheetFunction.Round(Weight_To_LBS, -1), vbInformation, "Converted!"
        
        End Select
        
    Else
    
        MsgBox "Please Select Conversion Process", vbInformation, "Converted!"
        Me.Conversion_Sel.BackColor = vbYellow
        
    End If

End Sub

