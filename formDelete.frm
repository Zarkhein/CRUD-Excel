VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formDelete 
   Caption         =   "XX-XXXX| Supprimer une machine"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8265
   OleObjectBlob   =   "formDelete.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function searchComputer2(numSeries As String)
    
    Dim i As Integer
    Dim count As Integer
    Dim valueCellule As String
    
    count = range("A" & Rows.count).End(xlUp).Row
    
    If numSeries <> Empty Then
        For i = 0 To count
            valueCellule = range("A" & 5 + i).Value
            If numSeries = range("A" & 5 + i).Value Then
                range("A" & 5 + i).Value = Empty
                range("B" & 5 + i).Value = Empty
                range("C" & 5 + i).Value = Empty
                range("D" & 5 + i).Value = Empty
                range("E" & 5 + i).Value = Empty
                range("F" & 5 + i).Value = Empty
                range("G" & 5 + i).Value = Empty
                
                range("C" & 5 + i).Interior.Color = xlNone
                
                formAdd.nbrPortable = range("J12").Value '5
                range("J12").Value = formAdd.nbrPortable - 1 '4
                Exit For
            ElseIf (i = count And numSeries <> range("A" & 5 + i).Value) Then
                MsgBox ("Erreur")
            End If
        Next i
    End If
    
    
End Function

Private Sub btnSend_Click()
    searchComputer2 (txtNum.Value)
End Sub


