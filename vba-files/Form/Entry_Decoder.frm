VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Entry_Decoder 
   Caption         =   "Entry Decoder"
   ClientHeight    =   5537
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   9961.001
   OleObjectBlob   =   "Entry_Decoder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Entry_Decoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    Call Center_Form(Me)
    Call Theme_Form(Me)
    
    For Row_Number = 1 To Count_Entries

        Me.RefID_Selector.AddItem [Main_Log[RefID]].Cells(Row_Number)

    Next
    
End Sub
Private Sub RefID_Selector_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim Entry_ID As String
    Dim Entry_Status As String
    Dim Scale_Process As String
    Dim Log_Process As String
    Dim Additional_Info As String
    
    With Me
    
        .Decoder_ID = vbNullString
        .Decoder_Tank_Number = vbNullString
        .Decoder_Date_In = vbNullString
        .Decoder_Date_Out = vbNullString
        .Decoder_Time_In = vbNullString
        .Decoder_Time_Out = vbNullString
        .Decoder_Initials_In = vbNullString
        .Decoder_Initials_Out = vbNullString
        .Weigh_Process = vbNullString
        .Log_Process = vbNullString
        .Decoder_Product = vbNullString
        
    End With
    
    Scale_Process = vbNullString
    Log_Process = vbNullString
    Additional_Info = vbNullString
    
    Target_Row = Application.WorksheetFunction.Match(EntryRef, Sheets("Full Log").Range("Main_Log[RefID]"), 0)
    Entry_ID = [Main_Log[ID]].Rows(Target_Row)
    Entry_Status = [Main_Log[Status]].Rows(Target_Row)
    
    Select Case Left([Main_Log[ID]].Rows(Target_Row), 1)
    
        Case PREFIX_STORAGE_ID, PREFIX_STORAGE_ID_2:
        
            Me.Decoder_Tank_Type = "Storage"
            Me.Decoder_Entry_Type = "Internal"
            
        Case PREFIX_CENTRAL_ID, PREFIX_CENTRAL_ID_2:
        
            Me.Decoder_Tank_Type = "Central"
            Me.Decoder_Entry_Type = "Internal"
            
        Case PREFIX_DROP_ID, PREFIX_DROP_ID_2:
        
            Me.Decoder_Tank_Type = "Drop"
            Me.Decoder_Entry_Type = "External"
            
        Case Else:
        
            Me.Decoder_Tank_Type = "Live"
            Me.Decoder_Entry_Type = "External"
            
    End Select
    
    With Me
    
        .Decoder_ID = Entry_ID
        .Decoder_Tank_Number = [Main_Log[Tank '#]].Rows(Target_Row)
        .Decoder_Date_In = [Main_Log[Date In]].Rows(Target_Row)
        .Decoder_Date_Out = [Main_Log[Date Out]].Rows(Target_Row)
        .Decoder_Time_In = [Main_Log[Time In]].Rows(Target_Row)
        .Decoder_Time_Out = [Main_Log[Time Out]].Rows(Target_Row)
        .Decoder_Initials_In = [Main_Log[Int In]].Rows(Target_Row)
        .Decoder_Initials_Out = [Main_Log[Int Out]].Rows(Target_Row)
        .Decoder_Product = [Main_Log[Product Name]].Rows(Target_Row)
        
    End With
    
    On Error Resume Next
    
    Me.Decoder_Refrence_Number = [Main_Log[Product Name]].Rows(Target_Row).CommentThreaded.Text
    Me.Decoder_Rejection_Message = [Main_Log[Int Out]].Rows(Target_Row).CommentThreaded.Text
    
    On Error GoTo 0
    
    If Me.Decoder_Entry_Type = "Internal" Then
    
        Scale_Process = "Weigh the tank out on the scale printer under it's current ID:" & INSERT_SPACE _
        & Me.Decoder_ID & "." & INSERT_SPACE & "Once that is done you will make a new entry and get the next available ID for" _
        & INSERT_SPACE & Me.Decoder_Tank_Type & INSERT_SPACE & "Tanks"
        Log_Process = "Use the Add Entry button as it will add the new entry and will go through the internal process and weigh out the past entry for you. If you are entering the entries manually you will need to find the old entry and weigh it out yourself."
    
    End If
    
    If Me.Decoder_Tank_Type = "Drop" Then
    
        Scale_Process = "Weigh in the tank without the truck attached and on the scale. this applys to when its weighed in and weighed out."
        Log_Process = "When adding the entry make sure that you put ""DW"" in the place of the truck number as we will not need this information."
    
    End If
    
    If Me.Decoder_Tank_Type = "Live" Then
    
        Scale_Process = "The tank will be weighed in and out with both truck and trailer attached."
        Log_Process = "Weigh in the entry and weigh out the entry"
                
    End If
    
    If Me.Decoder_Product = "Liquid Nitrogen" Then
    
        Additional_Info = "Since this is a liquid nitrogen tank if the driver is filling up both onsite tanks. Two different entries will need to be created."
    
    End If
    
    Me.Weigh_Process = "When this tank is weighed the process will go as follow:" & INSERT_SPACE & Scale_Process
    Me.Log_Process = "When logging in this tank after it is weighed you will need to:" & INSERT_SPACE & Log_Process & INSERT_SPACE & Additional_Info
    
End Sub

