VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewResin2 
   Caption         =   "New resin worksheet (Admin only)"
   ClientHeight    =   7550
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   10080
   OleObjectBlob   =   "NewResin2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewResin2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub AddBTN_Click()
Addtest.Show
End Sub


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












Private Sub Label3_Click()

End Sub

Private Sub RemoveBTN_Click()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim i As Integer
    Dim selectedTest As String

    If NewResin2.TestBar.ListIndex <> -1 Then ' Check if an item is selected
        selectedTest = NewResin2.TestBar.List(NewResin2.TestBar.ListIndex)
        For i = 1 To 5
            If ws.Cells(4, i).Value = selectedTest Then
                ws.Cells(4, i).Value = "" ' Clear the cell
                NewResin2.TestBar.RemoveItem NewResin2.TestBar.ListIndex ' Remove from listbox
                Exit Sub
            End If
        Next i
    Else
        MsgBox "Please select a test to remove."
    End If
End Sub

Private Sub TestBar_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim ws As Worksheet
    Dim selectedTest As String
    Dim i As Integer
    Dim cell As Range

    Set ws = ActiveSheet ' Assuming the active sheet is where you want to add the value

    ' Check if an item is selected in the list box
    If NewResin2.TestBar.ListIndex > -1 Then
        selectedTest = NewResin2.TestBar.List(NewResin2.TestBar.ListIndex)
        
        ' Find the first empty cell in row 3
        Set cell = ws.Cells(3, 1)
        i = 1
        While cell.Value <> ""
            i = i + 1
            Set cell = ws.Cells(3, i)
        Wend
        
        ' Place the selected test in the found cell
        cell.Value = selectedTest
    Else
        MsgBox "Please select a test to add."
    End If
End Sub


Private Sub UserForm_Click()

End Sub

