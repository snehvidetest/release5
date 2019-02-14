VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm006 
   Caption         =   "Populationsafgrænsning"
   ClientHeight    =   6936
   ClientLeft      =   36
   ClientTop       =   180
   ClientWidth     =   10980
   OleObjectBlob   =   "frm006.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label1_Click()

End Sub

Public Sub OKButton_Click()
    ' Gem svar i "SpmSvar"
'    Worksheets("SpmSvar").Range("C14:C14").Value = Controls("Label1").Caption
'    Worksheets("SpmSvar").Range("C15:C15").Value = Controls("Label2").Caption
'    Worksheets("SpmSvar").Range("C16:C16").Value = Controls("Label3").Caption
    
    ' Valider svar
    If ((OptionButton1 = False And OptionButton2 = False) Or (OptionButton3 = False And OptionButton4 = False) Or (OptionButton5 = False And OptionButton6 = False)) Then
        dFunc.msgError = "Besvar venligst alle spørgsmål."
        'frmMsg.Show
        SFunc.ShowFunc ("frmMsg")
        'MsgBox ("Besvar venligst alle spørgsmål.")
        GoTo ending
    End If
    
    ' Gem svar
    If OptionButton1 Then
'        Worksheets("SpmSvar").Range("D14:D14").Value = "Ja"
        Call writeSpmSvar("6", Controls("Label1").caption, "Ja")
    Else
        'Worksheets("SpmSvar").Range("D14:D14").Value = "Nej"
        Call writeSpmSvar("6", Controls("Label1").caption, "Nej")
    End If
        
    If OptionButton3 Then
        'Worksheets("SpmSvar").Range("D15:D15").Value = "Ja"
        Call writeSpmSvar("7", Controls("Label1").caption, "Ja")
    Else
        'Worksheets("SpmSvar").Range("D15:D15").Value = "Nej"
        Call writeSpmSvar("7", Controls("Label1").caption, "Nej")
    End If
    
    If OptionButton5 Then
        'Worksheets("SpmSvar").Range("D16:D16").Value = "Ja"
        Call writeSpmSvar("8", Controls("Label1").caption, "Ja")
    Else
        'Worksheets("SpmSvar").Range("D16:D16").Value = "Nej"
        Call writeSpmSvar("8", Controls("Label1").caption, "Nej")
    End If
    
    ' Gå tilbage hvis der svares "JA" på et eller flere af de tre spørgsmål
    If (OptionButton1 Or OptionButton3 Or OptionButton5) Then
        dFunc.msgError = "FlexFilter konfiguration kan ikke udføres pba. denne besvarelse."
        SFunc.ShowFunc ("frmMsg")
        'MsgBox ("FlexFilter konfiguration kan ikke udføres pba. denne besvarelse.")
    Else
        Me.Hide
        'store current form#
        recHis ("frm006")
        SFunc.ShowFunc ("frm007")
        'frm007.Show
    
    End If
    
ending:

End Sub


Private Sub OptionButton3_Click()

End Sub

Public Sub Tilbage_Click()
    Me.Hide
    'go back to previously stored form#
    Call goBack
    'SFunc.ShowFunc ("frm005")
    'frm005.Show
End Sub

Private Sub UserForm_Initialize()

    Image1.PictureSizeMode = fmPictureSizeModeClip
    ' Indlæs tidligere svar
    If findPreviousAns(findTopSpm("F"), "6", 1) = "Ja" Then
        OptionButton1.Value = True
    ElseIf findPreviousAns(findTopSpm("F"), "6", 1) = "Nej" Then
        OptionButton2.Value = True
    End If
    
        If findPreviousAns(findTopSpm("F"), "7", 1) = "Ja" Then
        OptionButton3.Value = True
    ElseIf findPreviousAns(findTopSpm("F"), "7", 1) = "Nej" Then
        OptionButton4.Value = True
    End If
    
        If findPreviousAns(findTopSpm("F"), "8", 1) = "Ja" Then
        OptionButton5.Value = True
    ElseIf findPreviousAns(findTopSpm("F"), "8", 1) = "Nej" Then
        OptionButton6.Value = True
    End If

Call drawProgressBar(Me, Me.Name)
End Sub
