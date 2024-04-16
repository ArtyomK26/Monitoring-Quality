VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MinMaxForm 
   Caption         =   "UserForm2"
   ClientHeight    =   12440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   14310
   OleObjectBlob   =   "MinMaxForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MinMaxForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub DoneBtn_Click()
    ' Assuming SelectedWorksheet is the one where the data needs to go
    If SelectedWorksheet Is Nothing Then
        MsgBox "No worksheet selected. Please select a worksheet first."
        Exit Sub
    End If

    Dim txtBoxMin As MSForms.TextBox
    Dim txtBoxMax As MSForms.TextBox
    Dim ctrlNameMin As String
    Dim ctrlNameMax As String

    For i = 1 To 15
        ctrlNameMin = "Textbox" & (i * 2 - 1)
        ctrlNameMax = "Textbox" & (i * 2)

        On Error Resume Next ' Ignore errors in the next statement
        Set txtBoxMin = Me.Controls(ctrlNameMin)
        Set txtBoxMax = Me.Controls(ctrlNameMax)
        If Err.Number <> 0 Then
            MsgBox "Could not find textbox: " & ctrlNameMin & " or " & ctrlNameMax
            Err.Clear
            Exit For
        End If
        On Error GoTo 0 ' Stop ignoring errors

        ' Now write the values to the worksheet
        SelectedWorksheet.Cells(3, i + 4).Value = txtBoxMin.Text ' Min value
        SelectedWorksheet.Cells(4, i + 4).Value = txtBoxMax.Text ' Max value
    Next i

    ' Adjust columns width to fit the content
    SelectedWorksheet.Columns("A:X").AutoFit


     Set SelectedWorksheet = ActiveSheet
    Set SelectedWorksheet = ThisWorkbook.Sheets("SheetName")



    ' Unload this form
    Unload Me
End Sub
Private Sub UserForm_Activate()
    InitializeControls
End Sub

Private Sub InitializeControls()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Or however you reference the relevant sheet

    Dim lastCol As Long
    Dim i As Integer
    Dim ctrlIndex As Integer
    
    ' Find the last used column starting from E (5th column) in row 5
    lastCol = ws.Cells(5, ws.Columns.Count).End(xlToLeft).Column
    
    ' Hide all controls initially
    For i = 1 To 15 ' Since you have 15 labels and 30 textboxes
        Me.Controls("Label" & i).Visible = False
        Me.Controls("Textbox" & (i * 2 - 1)).Visible = False ' Odd for Min
        Me.Controls("Textbox" & (i * 2)).Visible = False ' Even for Max
    Next i

    ' Make visible only the necessary ones and assign values
    ctrlIndex = 1
    For i = 5 To lastCol
        If ctrlIndex <= 15 Then ' Ensure you don't exceed the control count
            With Me
                .Controls("Label" & ctrlIndex).Caption = ws.Cells(5, i).Value
                .Controls("Label" & ctrlIndex).Visible = True

                ' Directly assign the value from the cells to the textboxes
                .Controls("Textbox" & (ctrlIndex * 2 - 1)).Value = ws.Cells(3, i).Value ' Min value
                .Controls("Textbox" & (ctrlIndex * 2 - 1)).Visible = True ' Min box

                .Controls("Textbox" & (ctrlIndex * 2)).Value = ws.Cells(4, i).Value ' Max value
                .Controls("Textbox" & (ctrlIndex * 2)).Visible = True ' Max box
            End With
            ctrlIndex = ctrlIndex + 1
        Else
            Exit For
        End If
    Next i



End Sub



