VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InspectResin 
   Caption         =   "Inspection"
   ClientHeight    =   7440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7780
   OleObjectBlob   =   "InspectResin.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "InspectResin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub InspectResin_Initialize()
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ListSheets.AddItem ws.Name
    Next ws
    
End Sub


Private Sub SearchBox_Change()
    
    Dim ws As Worksheet
    ListSheets.Clear
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "*" & SearchBox.Text & "*" Then
            ListSheets.AddItem ws.Name
        End If
    Next ws
End Sub



Private Sub SelectBtN_Click()
    If ListSheets.ListIndex <> -1 Then
        Dim SelectedSheet As String
        SelectedSheet = ListSheets.List(ListSheets.ListIndex)
        Set SelectedWorksheet = ThisWorkbook.Worksheets(SelectedSheet)
        
        ThisWorkbook.Worksheets(SelectedSheet).Activate
        InspectDesc.Show
        
        
    Else
        MsgBox "Please select a worksheet from the list."
    End If
End Sub


Private Sub CancelButton_Click()
    Unload Me
End Sub


