VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} enterPASS 
   Caption = "Password"
ClientHeight    =   570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "enterPASS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "enterPASS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub buttonOK_Click()
Dim strPass As String

    strPass = "password1"

    If Pasword.Value = strPass Then
        MsgBox "Welcome, master!"

        Application.Visible = True
        Sheets("Home").Select
        enterPASS.Hide
        ContactBook.Hide
    Else
        If MsgBox("The password was wrong, try again", vbYesNo) = vbNo Then
    enterPASS.Hide
    Else
            Password.Value = ""
        End If
End If
End Sub

Private Sub UserForm_Initialize()
    Password.Value = ""
End Sub
