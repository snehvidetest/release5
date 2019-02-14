VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm045 
   Caption         =   "Advarsel"
   ClientHeight    =   3936
   ClientLeft      =   48
   ClientTop       =   180
   ClientWidth     =   4884
   OleObjectBlob   =   "frm045.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm045"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub CommandButton1_Click()
Me.Hide
'go back to previously stored form#
Call goBack
End Sub

Public Sub CommandButton2_Click()
Me.Hide
'store current form#
'recHis (Me.Name)
frm035.Hide
SFunc.ShowFunc ("frm036")
'frm036.Show

End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

