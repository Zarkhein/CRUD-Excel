VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formUpdate 
   Caption         =   "XX-XXXX | Modifier une machine"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8265
   OleObjectBlob   =   "formUpdate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    range("A" & formSearch.valueCell).Value = UCase(txtNum.Text)
    range("B" & formSearch.valueCell).Value = UCase(txtNamePc.Text)
    range("C" & formSearch.valueCell).Value = UCase(txtIp.Text)
    range("D" & formSearch.valueCell).Value = UCase(txtUser.Text)
    range("E" & formSearch.valueCell).Value = UCase(txtPos.Text)
    range("F" & formSearch.valueCell).Value = UCase(cmbPc.Value)
    range("G" & formSearch.valueCell).Value = UCase(cmbStatut.Value)
    formAdd.CheckAllIp
End Sub


Private Sub UserForm_Initialize()
    txtNum.Text = range("A" & formSearch.valueCell).Value
    txtNamePc.Text = range("B" & formSearch.valueCell).Value
    txtIp.Text = range("C" & formSearch.valueCell).Value
    txtUser.Text = range("D" & formSearch.valueCell).Value
    txtPos.Text = range("E" & formSearch.valueCell).Value
    cmbPc.Value = range("F" & formSearch.valueCell).Value
    cmbStatut.Value = range("G" & formSearch.valueCell).Value
    
    
    With formUpdate.cmbPc
        .AddItem "DELL Latitude 5590"
        .AddItem "Dell Latitude 7490"
        .AddItem "Dell Latitude 7410"
        .AddItem "LENOVO X13"
    End With
    With formUpdate.cmbStatut
        .AddItem "Stock"
        .AddItem "Service"
    End With
    
End Sub
