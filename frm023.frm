VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm023 
   Caption         =   "Frasortering"
   ClientHeight    =   6936
   ClientLeft      =   60
   ClientTop       =   180
   ClientWidth     =   10980
   OleObjectBlob   =   "frm023.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm023"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub OKButton_Click()
If OptionButton1.Value = False And OptionButton2.Value = False Then
    dFunc.msgError = "Vælg venligst et svar for at forsætte"
    SFunc.ShowFunc ("frmMsg")
    GoTo ending
End If
    
'Worksheets("SpmSvar").Range("C41:C41").Value = Controls("Label1").caption
    
' Worksheets("Konfiguration").Activate
' Activate sheet

If OptionButton2.Value = True Then
    Worksheets("Regler").Range("J24:J24").Value = "-1825"
    Worksheets("Regler").Range("J25:J25").Value = "-1825"
    Worksheets("Regler").Range("J26:J26").Value = "-1825"
    Worksheets("Regler").Range("J27:J27").Value = "-1825"
    Worksheets("Regler").Range("J28:J28").Value = "-1825"
    
    Worksheets("Regler").Range("M24:M24").Value = "1"
    Worksheets("Regler").Range("M25:M25").Value = "1"
    Worksheets("Regler").Range("M26:M26").Value = "1"
    Worksheets("Regler").Range("M27:M27").Value = "1"
    Worksheets("Regler").Range("M28:M28").Value = "1"
End If

If OptionButton1.Value = True Then

    Call writeSpmSvar("14", Controls("Label1").caption, Controls("OptionButton1").caption)
    Me.Hide
    'store current form#
    recHis ("frm023")
    SFunc.ShowFunc ("frm017")
    'frm017.Show
    
ElseIf frm005.OptionButton1.Value = True Then

    
    Call writeSpmSvar("14", Controls("Label1").caption, Controls("OptionButton2").caption)
    Me.Hide
    'store current form#
    recHis ("frm023")
    SFunc.ShowFunc ("frm024")
    'frm024.Show
'ElseIf frm027.OptionButton1.Value = True Then
'    Me.Hide
'    'store current form#
'    recHis ("frm023")
'    SFunc.ShowFunc ("frm025")
'    'frm025.Show
End If

ending:
End Sub


Public Sub Tilbage_Click()

    Me.Hide
    'go back to previously stored form#
    Call goBack
    'SFunc.ShowFunc ("frm022")
    'frm022.Show

End Sub

Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeClip

' Indlæs tidligere svar 14
If findPreviousAns(findTopSpm("F"), "14", 1) = "Ja" Then
    OptionButton1.Value = True
ElseIf findPreviousAns(findTopSpm("F"), "14", 1) = "Nej" Then
    OptionButton2.Value = True
Else
    OptionButton1.Value = False
    OptionButton2.Value = False
End If

'If Worksheets("SpmSvar").Range("D41:D41").Value = "Ja" Then
'    OptionButton1.Value = True
'ElseIf Worksheets("SpmSvar").Range("D41:D41").Value = "Nej" Then
'    OptionButton2.Value = True
'Else
'    OptionButton1.Value = False
'    OptionButton2.Value = False
'End If

Call drawProgressBar(Me, Me.Name)
End Sub
