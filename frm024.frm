VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm024 
   Caption         =   "Frasortering"
   ClientHeight    =   6048
   ClientLeft      =   84
   ClientTop       =   372
   ClientWidth     =   9024.001
   OleObjectBlob   =   "frm024.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm024"
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

If OptionButton2.Value = True Then
    Worksheets("Regler").Range("J29:J29").Value = "-1825"
    Worksheets("Regler").Range("J30:J30").Value = "-1825"
    Worksheets("Regler").Range("J31:J31").Value = "-1825"
    Worksheets("Regler").Range("J32:J32").Value = "-1825"
    Worksheets("Regler").Range("J33:J33").Value = "-1825"
    
    Worksheets("Regler").Range("M29:M29").Value = "1"
    Worksheets("Regler").Range("M30:M30").Value = "1"
    Worksheets("Regler").Range("M31:M31").Value = "1"
    Worksheets("Regler").Range("M32:M32").Value = "1"
    Worksheets("Regler").Range("M33:M33").Value = "1"
End If
    
    
    If OptionButton1.Value = True Then
        'Worksheets("SpmSvar").Range("D42:D42").Value = "Ja"
        Call writeSpmSvar("15", Controls("Label1").caption, Controls("OptionButton1").caption)
        Me.Hide
        'store current form#
        recHis ("frm024")
        SFunc.ShowFunc ("frm019")
        'frm019.Show
        
    ElseIf OptionButton2.Value = True Then
        'Worksheets("SpmSvar").Range("D42:D42").Value = "Nej"
        Call writeSpmSvar("15", Controls("Label1").caption, Controls("OptionButton2").caption)
        Me.Hide
        'store current form#
        recHis ("frm024")
        SFunc.ShowFunc ("frm025")
        'frm025.Show
    End If


ending:
End Sub


Public Sub Tilbage_Click()
    Me.Hide
    'go back to previously stored form#
    Call goBack
    'SFunc.ShowFunc ("frm023")
    'frm023.Show
End Sub

Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeStretch

' Indlæs tidligere svar 15
If findPreviousAns(findTopSpm("F"), "15", 1) = "Ja" Then
    OptionButton1.Value = True
ElseIf findPreviousAns(findTopSpm("F"), "15", 1) = "Nej" Then
    OptionButton2.Value = True
Else
    OptionButton1.Value = False
    OptionButton2.Value = False
End If
'If Worksheets("SpmSvar").Range("D42:D42").Value = "Ja" Then
'    OptionButton1.Value = True
'ElseIf Worksheets("SpmSvar").Range("D42:D42").Value = "Nej" Then
'    OptionButton2.Value = True
'Else
'    OptionButton1.Value = False
'    OptionButton2.Value = False
'End If



End Sub
