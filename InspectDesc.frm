VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InspectDesc 
   Caption         =   "Inspection"
   ClientHeight    =   15490
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   11780
   OleObjectBlob   =   "InspectDesc.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "InspectDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Userform_Initialize()
    Dim ws As Worksheet
    Dim i As Integer
    
    
    Set ws = ActiveSheet
    
    
    For i = 1 To 17
        With Me.Controls("Label" & i)
            If ws.Cells(5, i).Value <> "" Then
                .Caption = ws.Cells(5, i).Value
                .Visible = True
                Me.Controls("TextBox" & i).Visible = True
            Else
                .Visible = False
                Me.Controls("TextBox" & i).Visible = False
            End If
        End With
    Next i
End Sub

