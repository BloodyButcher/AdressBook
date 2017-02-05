VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Danbook 
   Caption         =   "UserForm1"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9420.001
   OleObjectBlob   =   "Interface.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ContactBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
'DialogBox.Open
'TextBox1.PasswordChar = vbNullChar
enterPASS.Show


End Sub

Private Sub frmEmpDetails_Initialize()
    Me.Label4.Caption = Format(Of Date,'dd/mmmm/yyyy')
    Application.Visible = False
End Sub

Private Sub cmdClose_Click()
    ThisWorkbook.Save
    ThisWorkbook.Close
   
End Sub

Private Sub cmdOpen_Click()
    frmEmpDetails.Show
End Sub

Private Sub cmdRozpiski_Click()
    Rozpiski.Show
End Sub


Private Sub UserForm_Deactivate()
    ThisWorkbook.Save
    ThisWorkbook.Close
End Sub
