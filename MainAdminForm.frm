VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainAdminForm 
   Caption         =   "AdminMain"
   ClientHeight    =   13440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   23470
   OleObjectBlob   =   "MainAdminForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "MainAdminForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label9_Click()

    NewResin.Show
    
End Sub

Private Sub UserForm_Activate()
    RoletxtLBL.Caption = CurrentUserRole
End Sub


