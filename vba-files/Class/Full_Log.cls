VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Full_Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Add_Tank_Entry_Button_Click()

    Call Set_Next_ID

    Process_Mode = Add_New

    Entry_Box.Show vbModeless

End Sub
Private Sub Next_Line_Button_Click()

    Dim NextLine As Long

    NextLine = Count_Entries + 1
    
    [Main_Log[ID]].Rows(NextLine).Select

End Sub
Private Sub Dashboard_Button_Click()
    
    Admin_Box.Show vbModeless

End Sub
Private Sub Tools_Button_Click()
    
    Tools.Show vbModeless

End Sub
Private Sub Edit_Tank_Entry_Button_Click()
    
    Process_Mode = Edit
    
    Choose_Entry_Box.Show vbModeless

End Sub
Private Sub Help_Button_Click()

    Tools.Show vbModeless

End Sub
Private Sub Weigh_Out_Tank_Entry_Button_Click()
    
    Process_Mode = Out

    Choose_Entry_Box.Show vbModeless
    
End Sub
