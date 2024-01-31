VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Updater_Box 
   Caption         =   "Updater"
   ClientHeight    =   2905
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   4298
   OleObjectBlob   =   "Updater_Box.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Updater_Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Update_Log_Button_Click()
    
    Dim Out_Workbook As Workbook
    Dim In_Workbook As Workbook
    Dim Out_Row As ListRow
    Dim In_Row As ListRow
    Dim File_Name As String
    Dim File_Location As String
    Dim Row_Number As Long
    
    On Error GoTo ErrorHandler
    
    If Sel_Workbooks_Out.value = vbNullString Then
    
        MsgBox "Please select files", vbInformation, "Removed"
        
        Exit Sub
        
    End If
    
    MsgBox "Please wait. This may take some time. Click ok to begin", vbInformation, "Removed"
    
    Set Out_Workbook = Workbooks(Sel_Workbooks_Out.value)
    Set In_Workbook = Workbooks(ThisWorkbook.Name)

    Application.ScreenUpdating = False

    Row_Number = 1
    
    For Each Out_Row In Out_Workbook.Sheets("Full Log").ListObjects("Main_Log").ListRows
        
        Set In_Row = In_Workbook.Sheets("Full Log").ListObjects("Main_Log").ListRows(Row_Number)
            
        In_Row.Range.value = CStr(Out_Row.Range.value)

        Row_Number = Row_Number + 1
        
    Next Out_Row
    
    Row_Number = 1
    
    For Each Out_Row In Out_Workbook.Sheets("Storage Log").ListObjects("Internal_Log_1").ListRows
        
        Set In_Row = In_Workbook.Sheets("Storage Log").ListObjects("Internal_Log_1").ListRows(Row_Number)
            
        In_Row.Range.value = CStr(Out_Row.Range.value)

        Row_Number = Row_Number + 1
        
    Next Out_Row
    
    Row_Number = 1
    
    For Each Out_Row In Out_Workbook.Sheets("CFS Log").ListObjects("Internal_Log_2").ListRows
        
        Set In_Row = In_Workbook.Sheets("CFS Log").ListObjects("Internal_Log_2").ListRows(Row_Number)
            
        In_Row.Range.value = CStr(Out_Row.Range.value)

        Row_Number = Row_Number + 1
        
    Next Out_Row
    
    Call Generate_Data_List_From_Log(Carriers)
    Call Generate_Data_List_From_Log(Products)
    
    Call Rebuild_Log.Formulas
    
    Application.ScreenUpdating = True
    
    Out_Workbook.Close
    
    File_Name = Sel_Workbooks_Out.value
    
    File_Location = ActiveWorkbook.Path & "/" & File_Name
    
    ActiveWorkbook.SaveAs FileName:=File_Name
    
    MsgBox "Update Complete", vbInformation, "Done"
    
    Exit Sub
    
ErrorHandler:
    
    MsgBox "Warning! File Error.", vbCritical, "Error"
    
End Sub
Private Sub UserForm_Initialize()

    Call Theme_Form(Me)
    Call Center_Form(Me)

    Dim OpenWorkbook As Workbook

    For Each OpenWorkbook In Workbooks
    
        If OpenWorkbook.Name <> ThisWorkbook.Name Then
        
            Sel_Workbooks_Out.AddItem OpenWorkbook.Name
        
        End If
        
    Next OpenWorkbook

End Sub
