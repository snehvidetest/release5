VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMsg 
   Caption         =   "Besked"
   ClientHeight    =   2904
   ClientLeft      =   24
   ClientTop       =   108
   ClientWidth     =   5460
   OleObjectBlob   =   "frmMsg.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub CommandButton1_Click()

dFunc.msgError = ""
Unload Me

End Sub

Private Sub lblMsg_Click()

End Sub

Private Sub UserForm_Initialize()

lblMsg.caption = dFunc.msgError

End Sub

