VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UdviklerAdgang 
   Caption         =   "Udvikler Adgang"
   ClientHeight    =   1212
   ClientLeft      =   84
   ClientTop       =   372
   ClientWidth     =   4404
   OleObjectBlob   =   "UdviklerAdgang.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UdviklerAdgang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OK_Click()
'If Password.Value = "UFST123" Then
'    Dim xWs As Worksheet
'    For Each xWs In Application.ActiveWorkbook.Worksheets
'        xWs.Visible = True
'    Next
'    Me.Hide
'Else
'    Forkert_password.Visible = True
'End If
Dim xWs As Worksheet
Application.ScreenUpdating = True
For Each xWs In Application.ActiveWorkbook.Worksheets
    xWs.Visible = True
Next
Me.Hide
Worksheets("SpmSvar").Activate
Application.WindowState = xlNormal
End Sub

Private Sub Tilbage_Click()
Me.Hide
End Sub
