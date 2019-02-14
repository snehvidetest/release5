VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm012 
   Caption         =   "For�ldelseskontrol"
   ClientHeight    =   6936
   ClientLeft      =   36
   ClientTop       =   108
   ClientWidth     =   10980
   OleObjectBlob   =   "frm012.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Public Sub OKButton_Click()
If OptionButton1.Value = False And OptionButton2.Value = False Then
    dFunc.msgError = "V�lg venligst et svar for at fors�tte"
    SFunc.ShowFunc ("frmMsg")
    GoTo ending
End If

'Worksheets("SpmSvar").Range("C22:C22").Value = Controls("Label1").Caption
If OptionButton1.Value = True Then
'Worksheets("SpmSvar").Range("D22:D22").Value = "Ja"
Call writeSpmSvar("9.b.2", Controls("Label1").caption, "Ja")
End If

If OptionButton2.Value = True Then
'Worksheets("SpmSvar").Range("D22:D22").Value = "Nej"
Call writeSpmSvar("9.b.2", Controls("Label1").caption, "Nej")
End If

Me.Hide
'store current form#
recHis ("frm012")


' Worksheets("Konfiguration").Activate
' Activate sheet

If OptionButton1 = True Then
    ' G2 Aktiveres
    Worksheets("Regler").Range("G43:G47").Value = "JA"
    ' Stoler p� RIM
    Worksheets("Population").Range("B16:B16").Value = "JA"
    Worksheets("Regler").Range("G40:G40").Value = "JA"
    ' Gruppe 2 aktiveres
    Worksheets("Gruppering").Range("C3:C3").Value = "JA"
    SFunc.ShowFunc ("frm014")
    'frm014.Show
Else
    SFunc.ShowFunc ("frm013")
    'frm013.Show
End If

ending:
End Sub


Public Sub Tilbage_Click()
Me.Hide
'go back to previously stored form#
Call goBack
'SFunc.ShowFunc ("frm011")
'frm011.Show

End Sub

Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeClip

' Indl�s tidligere svar 9b2
If findPreviousAns(findTopSpm("F"), "9.b.2", 1) = "Ja" Then
    OptionButton1.Value = True
ElseIf findPreviousAns(findTopSpm("F"), "9.b.2", 1) = "Nej" Then
    OptionButton2.Value = True
Else
    OptionButton1.Value = False
    OptionButton2.Value = False
End If

Call drawProgressBar(Me, Me.Name)
End Sub
