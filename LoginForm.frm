VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Login"
   ClientHeight    =   5360
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9470.001
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Private Sub LoginBTN_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Control-Sheet")  ' Ensure this is the correct sheet name
    Dim found As Boolean
    found = False
    
    Dim r As Range
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row  ' Get the last row with data in column B
    
    For Each r In ws.Range("B3:B" & lastRow)
            
            If Trim(Me.cmbUsername.Value) = Trim(r.Value) And Trim(Me.txtPassword.Value) = Trim(r.Offset(0, 2).Value) Then  ' Using Trim to remove any accidental spaces
                found = True
                If Trim(r.Offset(0, 5).Value) = "Admin" Then  ' Make sure "Admin" matches exactly with what's in the sheet
                   
                Else
                    
                End If
                Me.Hide
                Exit For
            End If
        Next r
        
                If Not found Then
                    MsgBox "Invalid username or password!", vbCritical
                End If
                
            If found Then
            CurrentUserRole = Trim(r.Offset(0, 5).Value)  ' Store the role in the global variable
            Unload Me  ' Unload the login form completely
        
            If CurrentUserRole = "Admin" Then
                MainAdminForm.Show vbModal  ' Show as a modal dialog
            Else
                TeamMainForm.Show vbModal  ' Show as a modal dialog
            End If
        Else
            MsgBox "Invalid username or password!", vbCritical
        End If
        


End Sub




Private Sub Userform_Initialize()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Control-Sheet")  ' Change to your actual sheet name
    Dim userRange As Range
    Set userRange = ws.Range("B3:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row)

    For Each cell In userRange
        Me.cmbUsername.AddItem Trim(cell.Value)
    Next cell
End Sub

