VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm043 
   Caption         =   "Advarsel"
   ClientHeight    =   3672
   ClientLeft      =   96
   ClientTop       =   372
   ClientWidth     =   4152
   OleObjectBlob   =   "frm043.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm043"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub CommandButton1_Click()
Me.Hide
'store current form#
recHis ("frm043")

If frm004.ActiveControl Is Nothing Then
    ' ingen værdi
Else
        frm004.Hide
        SFunc.ShowFunc ("frm005")
        GoTo ending
End If

If frm002.ActiveControl Is Nothing Then
    ' ingen værdi
Else
    frm002.Hide
        If frm002.forkertData = False Then
            SFunc.ShowFunc ("frm003")
        Else
            SFunc.ShowFunc ("frm005")
        End If
        
        GoTo ending
End If











ending:
End Sub

Public Sub CommandButton2_Click()
Me.Hide
'store current form#
recHis ("frm043")
' frm002.txtModtStart.Value = ""
' frm002.txtModtSlut.Value = ""

End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()

End Sub
