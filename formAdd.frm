VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formAdd 
   Caption         =   "XX-XXXX | Ajouter une machine"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8265
   OleObjectBlob   =   "formAdd.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nbrPortable As Integer
Private Sub cmbStatut_Change()
    Dim bool As Boolean
    If cmbStatut.Value = "Stock" Then
        bool = False
        txtUser.Enabled = bool
        txtIp.Enabled = bool
        txtPos.Enabled = bool
        
        txtUser.BackColor = RGB(146, 142, 142)
        txtIp.BackColor = RGB(146, 142, 142)
        txtPos.BackColor = RGB(146, 142, 142)
    Else
        bool = True
        txtUser.Enabled = bool
        txtIp.Enabled = bool
        txtPos.Enabled = bool
        txtUser.BackColor = RGB(255, 255, 255)
        txtIp.BackColor = RGB(255, 255, 255)
        txtPos.BackColor = RGB(255, 255, 255)
    End If
End Sub

Private Sub CommandButton1_Click()

    Dim rangeTbl As Integer
    Dim tblOut As Variant
    Dim tmp As String
    
    rangeTbl = range("A4").End(xlDown).Row + 1
    
    'rangeTbl = range("G" & Rows.count).End(xlUp).Row
    tblOut = checkInfo()
    
    If tblOut(7) = 7 Or tblOut(7) = -2 Then
        addSheet tblOut(4), tblOut(0), tblOut(1), tblOut(2), tblOut(3), tblOut(5), tblOut(6), rangeTbl
        CheckAllIp
   
    End If
    
End Sub

Function addSheet(group, num, namePc, ip, user, marque, statut, tblDim)
        
        Dim pos As Integer
        pos = tblDim + 1
        nbrPortable = tblDim - 4
        
        range("A" & tblDim).Value = UCase(num)
        range("B" & tblDim).Value = UCase(namePc)
        range("C" & tblDim).Value = ip
        range("D" & tblDim).Value = UCase(group)
        range("E" & tblDim).Value = UCase(user)
        range("F" & tblDim).Value = UCase(marque)
        range("G" & tblDim).Value = UCase(statut)
        range("J12").Value = tblDim - 4
        
End Function

Function checkInfo()
    
    Dim tblInfo(7) As Variant
    Dim check As Integer
    
    check = 0
    
    tblInfo(0) = txtNum.Text
    tblInfo(1) = txtNamePc.Text
    tblInfo(2) = txtIp.Text
    tblInfo(3) = txtUser.Text
    tblInfo(4) = txtPos.Text
    tblInfo(5) = cmbPc.Text
    tblInfo(6) = cmbStatut.Text
    
    If cmbStatut.Text = "Service" Then
        For i = 0 To 6
            If tblInfo(i) = Empty Then
                MsgBox ("Erreur le champ est vide")
            Else
                check = check + 1
            End If
        Next i
        tblInfo(7) = check
    Else
        For i = 0 To 1
            If tblInfo(i) = Empty Then
                MsgBox ("Erreur le champ est vide")
            Else
                check = check - 1
            End If
        Next i
        tblInfo(7) = check
    End If
    
    checkInfo = tblInfo
    
    txtNum.Text = Empty
    txtNamePc.Text = Empty
    txtIp.Text = Empty
    txtPos.Text = Empty
    txtUser.Text = Empty
    
    
    
End Function
Public Sub CheckAllIp()

    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    
    Dim valueCell As Integer

    n = range("C" & Rows.count).End(xlUp).Row
    ' MsgBox (n)
    For i = 1 To n
        valueCellI = range("C" & i)
        For j = 1 To n
            valueCellJ = range("C" & j)
            ' MsgBox (valueCellI & " == " & valueCellJ)
            If valueCellI = valueCellJ Then
                If valueCellI <> Empty Then
                    cpt = cpt + 1
                    If cpt >= 2 Then
                        range("C" & i).Interior.Color = RGB(255, 128, 128)
                        range("C" & j).Interior.Color = RGB(255, 128, 128)
                    Else
                        range("C" & i).Interior.Color = xlNone
                    End If
                End If
            End If
        Next j
        cpt = 0
    Next i
    
End Sub


Private Sub txtNum_Change()

End Sub

Private Sub UserForm_Initialize()
    With formAdd.cmbPc
        .AddItem "DELL Latitude 5590"
        .AddItem "Dell Latitude 7490"
        .AddItem "Dell Latitude 7410"
        .AddItem "LENOVO X13"
    End With
    With formAdd.cmbStatut
        .AddItem "Stock"
        .AddItem "Service"
    End With
End Sub


