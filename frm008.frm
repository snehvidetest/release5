VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm008 
   Caption         =   "For�ldelseskontrol"
   ClientHeight    =   6936
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   10980
   OleObjectBlob   =   "frm008.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label1_Click()

End Sub

Public Sub OKButton_Click()
'Worksheets("SpmSvar").Range("C18:C18").Value = Controls("Label1").Caption
If OptionButton1.Value = False And OptionButton2.Value = False Then
    dFunc.msgError = "V�lg venligst et svar for at fors�tte"
    SFunc.ShowFunc ("frmMsg")
    GoTo ending
End If

If OptionButton1.Value = True Then
    'Worksheets("SpmSvar").Range("D18:D18").Value = "Ja"
    Call writeSpmSvar("9.a", Controls("Label1").caption, "Ja")
    
    Worksheets("Gruppering").Range("C2:C2").Value = "NEJ"
    Worksheets("Gruppering").Range("C3:C3").Value = "JA"
    
    Worksheets("Population").Range("B16:B16").Value = "JA"
    Worksheets("Population").Range("B17:B17").Value = "NEJ"
ElseIf OptionButton1.Value = False Then
    'Worksheets("SpmSvar").Range("D18:D18").Value = "Nej"
    Call writeSpmSvar("9.a", Controls("Label1").caption, "Nej")
End If


Me.Hide
'store current form#
recHis ("frm008")

' Worksheets("Konfiguration").Activate
' Activate sheet

If OptionButton1 = True Then
    SFunc.ShowFunc ("frm039")
    'frm039.Show

ElseIf OptionButton2 = True Then
    SFunc.ShowFunc ("frm009")
    'frm009.Show
End If


ending:
End Sub



Private Sub OptionButton1_Click()

End Sub

Private Sub OptionButton2_Click()

End Sub

Public Sub Tilbage_Click()
Me.Hide

'go back to previously stored form#
Call goBack
'SFunc.ShowFunc ("frm007")
'frm007.Show
End Sub

Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeClip

' Indl�s tidligere svar 9a
If findPreviousAns(findTopSpm("F"), "9.a", 1) = "Ja" Then
    OptionButton1.Value = True
    
ElseIf findPreviousAns(findTopSpm("F"), "9.a", 1) = "Nej" Then
    OptionButton2.Value = True
Else
    OptionButton1.Value = False
    OptionButton2.Value = False
End If

Call drawProgressBar(Me, Me.Name)
End Sub
