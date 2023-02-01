VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formSearch 
   Caption         =   "XX-XXXX| Rechercher"
   ClientHeight    =   2415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "formSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public num As String
Public valueCell As Integer
Private Sub CommandButton1_Click()
    
    num = UCase(txtNum.Value)
    searchComputer num
    
End Sub


Function searchComputer(numSeries As String)
    
    Dim i As Integer
    Dim count As Integer
    Dim valueCellule As String
    
    
    count = range("A" & Rows.count).End(xlUp).Row
    MsgBox (count)
    
    If numSeries <> Empty Then
        For i = 0 To count
            valueCellule = range("A" & 5 + i).Value
            If numSeries = valueCellule Then
                valueCell = i + 5
                formUpdate.Show
                Exit For
            End If
        Next i
    End If
    MsgBox (valueCell)
End Function

Private Sub UserForm_Click()

End Sub
