VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEmpDetails 
   Caption = "CONTACTBOOK v1.2"
ClientHeight    =   11085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17100
   OleObjectBlob   =   "frmEmpDetails.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEmpDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnNew As Boolean
Dim TRows, i As Long

Private Sub cmdClose_Click()
    If cmdClose.Caption = "Close" Then
        Unload Me
    Else
        cmdClose.Caption = "Close"
        cmdNew.Enabled = True
        cmdDelete.Enabled = True
        ComboBox1.Enabled = True

    End If
End Sub

Private Sub cmdDelete_Click() 'delete button
    TRows = Worksheets("Data").Range("A1").CurrentRegion.Rows.Count
    Dim strDel
    strDel = MsgBox("Delete ?", vbYesNo, "Delete")
    If strDel = vbYes Then
        For i = 2 To TRows
            If Trim(Worksheets("Data").Cells(i, 1).Value) = Trim(ComboBox1.Text) Then

                '  Sheet1.Range(i & ":" & i).Delete
                Worksheets("Data").Range(i & ":" & i).Delete

                txtSAP_NR.Text = "" 'Customer SAP number
                txtCUSTOMER.Text = "" 'Cusotmer name
                txtCITY.Text = "" 'Customer city
                txtPOSTAL_CODE.Text = "" ' CUstomer postal code
                txtSTREET.Text = "" 'Customer street name
                txtSALES_REPRESENTATIVE.Text = "" 'Customer sales representative
                txtTEL_REPRESENTATIVE.Text = "" 'Phone number to sales representative
                txtKAM.Text = "" 'Key Account Manager
                txtCC.Text = "" 'Customer care employee responsible for client
                txtCONTACT_PERSON.Text = "" 'Contact person on customer side
                txtPOSITION.Text = "" 'Contatc person position on customer side
                txtTEL_1.Text = "" '1st Telephone number to customer
                txtTEL_2.Text = "" '2nd telephone number to cusotmer
                txtemail.Text = "" 'Customer e-mail adress
                txtFAX.Text = ""
                txtSOLD_TO.Text = "" 'optional when goods are sold to another customer number
                txtACTIVE.Text = "" 'is the customer active
                txtCC2.Text = "" 'Customer care employee responsible for client when main person is off
                txtCC_2.Text = "" 'Customer care employee responsible for client- shortcut
                CheckBoxMon.Value = "" 'check box for Monday
                CheckBoxTue.Value = "" 'check box for Tuesday
                CheckBoxWed.Value = "" 'check box for Wednesday
                CheckBoxThu.Value = "" 'check box for Thuesday
                CheckBoxFri.Value = "" 'check box for Friday
                CheckBoxSat.Value = "" 'check box for Saturday
                txtForm.Text = "" 'How client make and order
                txtUnit.Text = "" 'Unit on order
                txtPLOD.Text = "" 'Route
                txtWHS.Text = "" 'Warehouse
                txtCOT.Text = "" 'Cut of time
                txtCOMMENTS.Text = ""
                txtINFO.Text = "" 'additional information
                Call prComboBoxFill()

                Exit For
            End If
        Next i
        If Trim(ComboBox1.Text) = "" Then
            cmdSave.Enabled = False
            cmdDelete.Enabled = False
        Else
            cmdSave.Enabled = True
            cmdDelete.Enabled = True
        End If
        cmdNew.Enabled = True
        cmdClose.Caption = "CLOSE"


    End If

    If Trim(txtSAP_NR.Text) = "" Then
        cmdSave.Enabled = False
        cmdDelete.Enabled = False
        Frame2.Enabled = True
        cmdNew.Enabled = False
        ComboBox1.Enabled = True
        cmdClose.Caption = "CLOSE"
    Else
        cmdSave.Enabled = True
        cmdDelete.Enabled = True
        Frame2.Enabled = True
        cmdNew.Enabled = True
        cmdClose.Caption = "CLOSE"
    End If
End Sub

Private Sub cmdNew_Click()
    blnNew = True
    txtSAP_NR.Text = ""
    txtCUSTOMER.Text = ""
                txtCITY.Text = ""
                txtPOSTAL_CODE.Text = ""
                txtSTREET.Text = ""
                txtSALES_REPRESENTATIVE.Text = ""
                txtTEL_REPRESENTATIVE.Text = ""
                txtKAM.Text = ""
                txtCC.Text = ""
                txtCONTACT_PERSON.Text = ""
                txtPOSITION.Text = ""
                txtTEL_1.Text = ""
                txtTEL_2.Text = ""
                txtemail.Text = ""
                txtFAX.Text = ""
                txtSOLD_TO.Text = ""
                txtACTIVE.Text = ""
                txtCC2.Text = ""
                txtCC_2.Text = ""
                CheckBoxMon.Value = "0"
                CheckBoxTue.Value = "0"
                CheckBoxWed.Value = "0"
                CheckBoxThu.Value = "0"
                CheckBoxFri.Value = "0"
                CheckBoxSat.Value = "0"
                txtForm.Text = ""
                txtUnit.Text = ""
                txtPLOD.Text = ""
                txtWHS.Text = ""
                txtCOT.Text = ""
                txtCOMMENTS.Text = ""
                txtINFO.Text = ""
    cmdClose.Caption = "CANCEL"
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
    cmdSave.Enabled = True
    Frame2.Enabled = True
    ComboBox1.Enabled = True

End Sub

Private Sub cmdSave_Click()
    If Trim(txtSAP_NR.Text) = "" Then
        MsgBox("Enter customer SAP number", vbCritical, "Save")
        txtSAP_NR.SetFocus
        Exit Sub
    End If
    Call prSave()
    cmdClose.Caption = "Close"
    cmdNew.Enabled = True
    ComboBox1.Enabled = True
    ThisWorkbook.Save
    
End Sub
Private Sub prSave()
    
    If blnNew = True Then
        TRows = Worksheets("Data").Range("A1").CurrentRegion.Rows.Count
        With Worksheets("Data").Range("A1")
            .Offset(TRows, 0).Value = txtSAP_NR.Text
            .Offset(TRows, 1).Value = txtCUSTOMER.Text
            .Offset(TRows, 2).Value = txtCITY.Text
            .Offset(TRows, 3).Value = txtPOSTAL_CODE.Text
            .Offset(TRows, 4).Value = txtSTREET.Text
            .Offset(TRows, 5).Value = txtSALES_REPRESENTATIVE.Text
            .Offset(TRows, 6).Value = txtTEL_REPRESENTATIVE.Text
            .Offset(TRows, 7).Value = txtKAM.Text
            .Offset(TRows, 8).Value = txtCC.Text
            .Offset(TRows, 9).Value = txtCONTACT_PERSON.Text
            .Offset(TRows, 10).Value = txtPOSITION.Text
            .Offset(TRows, 11).Value = txtTEL_1.Text
            .Offset(TRows, 12).Value = txtTEL_2.Text
            .Offset(TRows, 13).Value = txtemail.Text
            .Offset(TRows, 14).Value = txtFAX.Text
            .Offset(TRows, 15).Value = txtSOLD_TO.Text
            .Offset(TRows, 16).Value = txtACTIVE.Text
            .Offset(TRows, 17).Value = txtCC2.Text
            .Offset(TRows, 18).Value = txtCC_2.Text
            .Offset(TRows, 19).Value = CheckBoxMon.Value = .Value
            .Offset(TRows, 20).Value = CheckBoxTue.Value = .Value
            .Offset(TRows, 21).Value = CheckBoxWed.Value = .Value
            .Offset(TRows, 22).Value = CheckBoxThu.Value = .Value
            .Offset(TRows, 23).Value = CheckBoxFri.Value = .Value
            .Offset(TRows, 24).Value = CheckBoxSat.Value = .Value
            .Offset(TRows, 25).Value = txtForm.Text
            .Offset(TRows, 26).Value = txtUnit.Text
            .Offset(TRows, 27).Value = txtPLOD.Text
            .Offset(TRows, 28).Value = txtWHS.Text
            .Offset(TRows, 29).Value = txtCOT.Text
            .Offset(TRows, 30).Value = txtCOMMENTS.Text
            .Offset(TRows, 31).Value = txtINFO.Text
          
         End With
        txtSAP_NR.Text = "" 'Customer SAP number
        txtCUSTOMER.Text = "" 'Cusotmer name
        txtCITY.Text = "" 'Customer city
        txtPOSTAL_CODE.Text = "" ' CUstomer postal code
        txtSTREET.Text = "" 'Customer street name
        txtSALES_REPRESENTATIVE.Text = "" 'Customer sales representative
        txtTEL_REPRESENTATIVE.Text = "" 'Phone number to sales representative
        txtKAM.Text = "" 'Key Account Manager
        txtCC.Text = "" 'Customer care employee responsible for client
        txtCONTACT_PERSON.Text = "" 'Contact person on customer side
        txtPOSITION.Text = "" 'Contatc person position on customer side
        txtTEL_1.Text = "" '1st Telephone number to customer
        txtTEL_2.Text = "" '2nd telephone number to cusotmer
        txtemail.Text = "" 'Customer e-mail adress
        txtFAX.Text = ""
        txtSOLD_TO.Text = "" 'optional when goods are sold to another customer number
        txtACTIVE.Text = "" 'is the customer active
        txtCC2.Text = "" 'Customer care employee responsible for client when main person is off
        txtCC_2.Text = "" 'Customer care employee responsible for client- shortcut
        CheckBoxMon.Value = "" 'check box for Monday
        CheckBoxTue.Value = "" 'check box for Tuesday
        CheckBoxWed.Value = "" 'check box for Wednesday
        CheckBoxThu.Value = "" 'check box for Thuesday
        CheckBoxFri.Value = "" 'check box for Friday
        CheckBoxSat.Value = "" 'check box for Saturday
        txtForm.Text = "" 'How client make and order
        txtUnit.Text = "" 'Unit on order
        txtPLOD.Text = "" 'Route
        txtWHS.Text = "" 'Warehouse
        txtCOT.Text = "" 'Cut of time
        txtCOMMENTS.Text = ""
        txtINFO.Text = "" 'additional information
        Call prComboBoxFill()

    Else
        For i = 2 To TRows
            If Trim(Worksheets("Data").Cells(i, 1).Value) = Trim(ComboBox1.Text) Then
                Worksheets("Data").Cells(i, 1).Value = txtSAP_NR.Text
                Worksheets("Data").Cells(i, 2).Value = txtCUSTOMER.Text
                Worksheets("Data").Cells(i, 3).Value = txtCITY.Text
                Worksheets("Data").Cells(i, 4).Value = txtPOSTAL_CODE.Text
                Worksheets("Data").Cells(i, 5).Value = txtSTREET.Text
                Worksheets("Data").Cells(i, 6).Value = txtSALES_REPRESENTATIVE.Text
                Worksheets("Data").Cells(i, 7).Value = txtTEL_REPRESENTATIVE.Text
                Worksheets("Data").Cells(i, 8).Value = txtKAM.Text
                Worksheets("Data").Cells(i, 9).Value = txtCC.Text
                Worksheets("Data").Cells(i, 10).Value = txtCONTACT_PERSON.Text
                Worksheets("Data").Cells(i, 11).Value = txtPOSITION.Text
                Worksheets("Data").Cells(i, 12).Value = txtTEL_1.Text
                Worksheets("Data").Cells(i, 13).Value = txtTEL_2.Text
                Worksheets("Data").Cells(i, 14).Value = txtemail.Text
                Worksheets("Data").Cells(i, 15).Value = txtFAX.Text
                Worksheets("Data").Cells(i, 16).Value = txtSOLD_TO.Text
                Worksheets("Data").Cells(i, 17).Value = txtACTIVE.Text
                Worksheets("Data").Cells(i, 18).Value = txtCC2.Text
                Worksheets("Data").Cells(i, 19).Value = txtCC_2.Text
                Worksheets("Data").Cells(i, 20).Value = CheckBoxMon.Value
                Worksheets("Data").Cells(i, 21).Value = CheckBoxTue.Value
                Worksheets("Data").Cells(i, 22).Value = CheckBoxWed.Value
                Worksheets("Data").Cells(i, 23).Value = CheckBoxThu.Value
                Worksheets("Data").Cells(i, 24).Value = CheckBoxFri.Value
                Worksheets("Data").Cells(i, 25).Value = CheckBoxSat.Value
                Worksheets("Data").Cells(i, 26).Value = txtForm.Text
                Worksheets("Data").Cells(i, 27).Value = txtUnit.Text
                Worksheets("Data").Cells(i, 28).Value = txtPLOD.Text
                Worksheets("Data").Cells(i, 29).Value = txtWHS.Text
                Worksheets("Data").Cells(i, 30).Value = txtCOT.Text
                Worksheets("Data").Cells(i, 31).Value = txtCOMMENTS.Text
                Worksheets("Data").Cells(i, 32).Value = txtINFO.Text

                txtSAP_NR.Text = "" 'Customer SAP number
                txtCUSTOMER.Text = "" 'Cusotmer name
                txtCITY.Text = "" 'Customer city
                txtPOSTAL_CODE.Text = "" ' CUstomer postal code
                txtSTREET.Text = "" 'Customer street name
                txtSALES_REPRESENTATIVE.Text = "" 'Customer sales representative
                txtTEL_REPRESENTATIVE.Text = "" 'Phone number to sales representative
                txtKAM.Text = "" 'Key Account Manager
                txtCC.Text = "" 'Customer care employee responsible for client
                txtCONTACT_PERSON.Text = "" 'Contact person on customer side
                txtPOSITION.Text = "" 'Contatc person position on customer side
                txtTEL_1.Text = "" '1st Telephone number to customer
                txtTEL_2.Text = "" '2nd telephone number to cusotmer
                txtemail.Text = "" 'Customer e-mail adress
                txtFAX.Text = ""
                txtSOLD_TO.Text = "" 'optional when goods are sold to another customer number
                txtACTIVE.Text = "" 'is the customer active
                txtCC2.Text = "" 'Customer care employee responsible for client when main person is off
                txtCC_2.Text = "" 'Customer care employee responsible for client- shortcut
                CheckBoxMon.Value = "" 'check box for Monday
                CheckBoxTue.Value = "" 'check box for Tuesday
                CheckBoxWed.Value = "" 'check box for Wednesday
                CheckBoxThu.Value = "" 'check box for Thuesday
                CheckBoxFri.Value = "" 'check box for Friday
                CheckBoxSat.Value = "" 'check box for Saturday
                txtForm.Text = "" 'How client make and order
                txtUnit.Text = "" 'Unit on order
                txtPLOD.Text = "" 'Route
                txtWHS.Text = "" 'Warehouse
                txtCOT.Text = "" 'Cut of time
                txtCOMMENTS.Text = ""
                txtINFO.Text = "" 'additional information
                Exit For
            End If
        Next i
      End If
    blnNew = False
    
    If Trim(txtSAP_NR.Text) = "" Then
        cmdSave.Enabled = False
        cmdDelete.Enabled = False
        Frame2.Enabled = True
        cmdClose.Caption = "Close"
    Else
        cmdSave.Enabled = True
        cmdDelete.Enabled = True
        Frame2.Enabled = True
        Frame3.Enabled = True
        cmdClose.Caption = "Close"
    End If
End Sub

Private Sub cmdSearch_Click()
    blnNew = False
    txtSAP_NR.Text = "" 'Customer SAP number
    txtCUSTOMER.Text = "" 'Cusotmer name
    txtCITY.Text = "" 'Customer city
    txtPOSTAL_CODE.Text = "" ' CUstomer postal code
    txtSTREET.Text = "" 'Customer street name
    txtSALES_REPRESENTATIVE.Text = "" 'Customer sales representative
    txtTEL_REPRESENTATIVE.Text = "" 'Phone number to sales representative
    txtKAM.Text = "" 'Key Account Manager
    txtCC.Text = "" 'Customer care employee responsible for client
    txtCONTACT_PERSON.Text = "" 'Contact person on customer side
    txtPOSITION.Text = "" 'Contatc person position on customer side
    txtTEL_1.Text = "" '1st Telephone number to customer
    txtTEL_2.Text = "" '2nd telephone number to cusotmer
    txtemail.Text = "" 'Customer e-mail adress
    txtFAX.Text = ""
    txtSOLD_TO.Text = "" 'optional when goods are sold to another customer number
    txtACTIVE.Text = "" 'is the customer active
    txtCC2.Text = "" 'Customer care employee responsible for client when main person is off
    txtCC_2.Text = "" 'Customer care employee responsible for client- shortcut
    CheckBoxMon.Value = "" 'check box for Monday
    CheckBoxTue.Value = "" 'check box for Tuesday
    CheckBoxWed.Value = "" 'check box for Wednesday
    CheckBoxThu.Value = "" 'check box for Thuesday
    CheckBoxFri.Value = "" 'check box for Friday
    CheckBoxSat.Value = "" 'check box for Saturday
    txtForm.Text = "" 'How client make and order
    txtUnit.Text = "" 'Unit on order
    txtPLOD.Text = "" 'Route
    txtWHS.Text = "" 'Warehouse
    txtCOT.Text = "" 'Cut of time
    txtCOMMENTS.Text = ""
    txtINFO.Text = "" 'additional information

    TRows = Worksheets("Data").Range("A1").CurrentRegion.Rows.Count
    For i = 2 To TRows
        If Val(Trim(Worksheets("Data").Cells(i, 1).Value)) = Val(Trim(ComboBox1.Text)) Then
                
            txtSAP_NR.Text = Worksheets("Data").Cells(i, 1).Value
            txtCUSTOMER.Text = Worksheets("Data").Cells(i, 2).Value
            txtCITY.Text = Worksheets("Data").Cells(i, 3).Value
            txtPOSTAL_CODE.Text = Worksheets("Data").Cells(i, 4).Value
            txtSTREET.Text = Worksheets("Data").Cells(i, 5).Value
            txtSALES_REPRESENTATIVE.Text = Worksheets("Data").Cells(i, 6).Value
            txtTEL_REPRESENTATIVE.Text = Worksheets("Data").Cells(i, 7).Value
            txtKAM.Text = Worksheets("Data").Cells(i, 8).Value
            txtCC.Text = Worksheets("Data").Cells(i, 9).Value
            txtCONTACT_PERSON.Text = Worksheets("Data").Cells(i, 10).Value
            txtPOSITION.Text = Worksheets("Data").Cells(i, 11).Value
            txtTEL_1.Text = Worksheets("Data").Cells(i, 12).Value
            txtTEL_2.Text = Worksheets("Data").Cells(i, 13).Value
            txtemail.Text = Worksheets("Data").Cells(i, 14).Value
            txtFAX.Text = Worksheets("Data").Cells(i, 15).Value
            txtSOLD_TO.Text = Worksheets("Data").Cells(i, 16).Value
            txtACTIVE.Value = Worksheets("Data").Cells(i, 17).Value
            txtCC2.Text = Worksheets("Data").Cells(i, 18).Value
            txtCC_2.Text = Worksheets("Data").Cells(i, 19).Value
            CheckBoxMon.Value = Worksheets("Data").Cells(i, 20).Value
            CheckBoxTue.Value = Worksheets("Data").Cells(i, 21).Value
            CheckBoxWed.Value = Worksheets("Data").Cells(i, 22).Value
            CheckBoxThu.Value = Worksheets("Data").Cells(i, 23).Value
            CheckBoxFri.Value = Worksheets("Data").Cells(i, 24).Value
            CheckBoxSat.Value = Worksheets("Data").Cells(i, 25).Value
            txtForm.Text = Worksheets("Data").Cells(i, 26).Value
            txtUnit.Text = Worksheets("Data").Cells(i, 27).Value
            txtPLOD.Text = Worksheets("Data").Cells(i, 28).Value
            txtWHS.Text = Worksheets("Data").Cells(i, 29).Value
            txtCOT.Text = Worksheets("Data").Cells(i, 30).Value
            txtCOMMENTS.Text = Worksheets("Data").Cells(i, 31).Value
            txtINFO.Text = Worksheets("Data").Cells(i, 32).Value
            Exit For
        End If
    Next i
    If Trim(txtSAP_NR.Text) = "" Then
        cmdSave.Enabled = False
        cmdDelete.Enabled = False
        Frame2.Enabled = False
        Frame3.Enabled = True
        Frame1.Enabled = True
        cmdNew.Enabled = True
    Else
        cmdSave.Enabled = True
        cmdDelete.Enabled = True
        Frame2.Enabled = True
        Frame3.Enabled = True
        Frame1.Enabled = True
        cmdClose.Caption = "Close"
    End If
End Sub
Private Sub prComboBoxFill()
    TRows = Worksheets("Data").Range("A1").CurrentRegion.Rows.Count
    ComboBox1.Clear
    For i = 2 To TRows
        ComboBox1.AddItem Worksheets("Data").Cells(i, 1).Value
    Next i
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)

End Sub

Private Sub txtACTIVE_Change()
    If txtACTIVE.Value = "YES" Then
        'txtACTIVE.ForeColor = RGB(0, 255, 0)
        txtACTIVE.BackColor = RGB(0, 255, 0)
    Else
        If txtACTIVE.Value = "NO" Then
            'txtACTIVE.ForeColor = RGB(255, 0, 0)
            txtACTIVE.BackColor = RGB(255, 0, 0)

        End If
    End If
End Sub


Private Sub txtCOT_Change()
        txtCOT.Text = Format(txtCOT, "h:nn")
End Sub


Private Sub UserForm_Initialize()
    Call prComboBoxFill
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    Frame2.Enabled = True
    cmdNew.Enabled = True
    cmdClose.Enabled = True
    cmdClose.Caption = "Close"
    Application.Visible = False
End Sub
