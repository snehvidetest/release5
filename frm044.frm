VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm044 
   Caption         =   "Advarsel"
   ClientHeight    =   3936
   ClientLeft      =   72
   ClientTop       =   216
   ClientWidth     =   4884
   OleObjectBlob   =   "frm044.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm044"
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

frm034.Hide
SFunc.ShowFunc ("frm035")
'frm035.Show

End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()

End Sub
