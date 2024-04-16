VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Addtest 
   Caption         =   "Add new resin"
   ClientHeight    =   3850
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8350.001
   OleObjectBlob   =   "Addtest.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Addtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelBTN_Click()
 Me.Hide
End Sub

Private Sub DoneBtn_Click()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim i As Integer
    
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "Lot#"
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "Description"
    Range("C5").Select
    ActiveCell.FormulaR1C1 = "Company"
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "Shelf life"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Min"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "Max"
    
        Range("A5:X5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Rows("5:5").Select
    Selection.RowHeight = 35
    
    
    
    
    For i = 5 To 20
        If ws.Cells(5, i).Value = "" Then
            ws.Cells(5, i).Value = Me.TestName.Text
            NewResin2.TestBar.AddItem Me.TestName.Text
            Unload Me
            Exit Sub
        End If
    Next i

    
    
    MsgBox "All test slots are filled."


End Sub


Private Sub TestName_Change()

End Sub
