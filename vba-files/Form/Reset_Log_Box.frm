VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Reset_Log_Box 
   Caption         =   "Reset Log"
   ClientHeight    =   2460
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   4564
   OleObjectBlob   =   "Reset_Log_Box.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Reset_Log_Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    
    Call Theme_Form(Me)
    Call Center_Form(Me)

End Sub
Private Sub Confirm_Pass_Button_Click()

    Dim File_Name As String
    Dim File_Location As String
    
    ThisWorkbook.Save
    
    Application.Goto Reference:=Sheets("Full Log").Range("A1"), Scroll:=True
    
    File_Location = Environ("USERPROFILE") & "\Documents\"
    
    
    If Keep_Data_Check = True Then
        
        Call Reset_Keep_Data
        
        File_Name = File_Location & "Main_Log" & "_" & Format(Date, "yyyy") & ".xlsm"
        
    Else
        
        Call Reset_No_Data
        
        File_Name = File_Location & "Main_Log" & "_" & "Blank" & ".xlsm"
        
    End If
    
    Call Rebuild_Log.Formulas
    Call Theme.Apply_Theme
    
    Call Set_Next_ID
    
    MsgBox "The Log Has Been Erased"
    
    ActiveWorkbook.SaveAs FileName:=File_Name

    MsgBox "The file has been saved to" & INSERT_SPACE & File_Name & "." & INSERT_SPACE & "Please save this to your preferred area."

    Unload Reset_Log_Box
    Unload Admin_Box

End Sub
Private Sub Reset_No_Data()

    Dim Row_Number As Long
    
    Current_Max_Entries = Strip_Array(ActiveWorkbook.Names("Option_Current_Max_Entries"))
        
    For Each Worksheet In ThisWorkbook.Worksheets
        
        For Each listOBJ In Worksheet.ListObjects
            
            If listOBJ.Name <> "Table_Next_ID" Then listOBJ.DataBodyRange.ClearContents
            
        Next listOBJ
        
    Next Worksheet
        
    For Row_Number = 1 To Current_Max_Entries
    
        Call Reset_ID(Row_Number)
        
    Next
    
    Sheets("Full Log").Range("Main_Log").ClearComments
    
End Sub
Sub Reset_Keep_Data()

    Dim Current_Row_Count As Long
    Dim Row_Number As Long

    Current_Row_Count = Count_Entries

    For Row_Number = 1 To Current_Row_Count
    
        If [Main_Log[Status]].Rows(Row_Number) = ENTRY_INACTIVE Then
       
            [Main_Log[ID]].Rows(Row_Number).EntireRow.Delete
            
            Current_Row_Count = Count_Entries
            
            Row_Number = Row_Number - 1
            
        End If
        
    Next
    
    Current_Row_Count = Count_Entries + 1

    [Main_Log[ID]].Rows(Current_Row_Count) = 1
    
    Current_Row_Count = Current_Row_Count + 1
    
    For Row_Number = Current_Row_Count To Current_Max_Entries
    
        Call Reset_ID(Row_Number)
        
    Next
    
End Sub


