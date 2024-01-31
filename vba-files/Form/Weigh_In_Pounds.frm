VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Weigh_In_Pounds 
   Caption         =   "Products In Pounds"
   ClientHeight    =   4802
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   4298
   OleObjectBlob   =   "Weigh_In_Pounds.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Weigh_In_Pounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    Call Theme_Form(Me)
    Call Center_Form(Me)
    
    Call Set_Control_From_Small_Data_List("List_Weigh_In_Pounds_Products", Me.Sel_Product)
    Call Set_Control_From_Small_Data_List("List_Weigh_In_Pounds_Products", Me.Product_Box)

End Sub
Private Sub Add_Plant_Item_Button_Click()

    Call Add_Item_To_Small_List("List_Weigh_In_Pounds_Products", Me.Sel_Product, "Pound", True)
    Me.Product_Box.Clear
    Call Set_Control_From_Small_Data_List("List_Weigh_In_Pounds_Products", Me.Product_Box)

End Sub
Private Sub Remove_PLant_Item_Button_Click()

    Call Remove_Item_From_Small_List("List_Weigh_In_Pounds_Products", Me.Sel_Product, "Product", True)
    Me.Product_Box.Clear
    Call Set_Control_From_Small_Data_List("List_Weigh_In_Pounds_Products", Me.Product_Box)

End Sub

