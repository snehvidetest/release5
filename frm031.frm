VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm031 
   Caption         =   "For�ldelseskontrol"
   ClientHeight    =   8940.001
   ClientLeft      =   36
   ClientTop       =   204
   ClientWidth     =   12984
   OleObjectBlob   =   "frm031.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm031"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public Sub CheckBox1_Click()
If CheckBox1.Value = True Then
    TextBox1.Value = ""
    TextBox1.Enabled = False
ElseIf CheckBox1.Value = False Then
    TextBox1.Enabled = True
End If

End Sub

Public Sub CheckBox2_Click()
If CheckBox2.Value = True Then
    TextBox2.Value = ""
    TextBox2.Enabled = False
    CheckBox3.Enabled = False
ElseIf CheckBox2.Value = False Then
    TextBox2.Enabled = True
    CheckBox3.Enabled = True
End If
End Sub

Public Sub CheckBox3_Click()
If CheckBox3.Value = True Then
    CheckBox2.Enabled = False
    TextBox2.Value = "1095"
    TextBox2.Enabled = False
ElseIf CheckBox3.Value = False Then
    CheckBox2.Enabled = True
    TextBox2.Value = ""
    TextBox2.Enabled = True
End If

End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label9_Click()

End Sub

Public Sub OKButton_Click()
' Validering - Hvis "Aldrig" i frm007 og "Ved ikke" er valgt
If (CheckBox1.Value = True Or CheckBox2.Value = True) And frm007.OptionButton3.Value = True Then
    dFunc.msgError = "Sp�rgeskemaet kan ikke anvendes p� baggrund af indtastede oplysninger"
    SFunc.ShowFunc ("frmMsg")
    Me.Hide
    'store current form#
    recHis ("frm031")
    SFunc.ShowFunc ("frm002")
    GoTo ending
End If

' Validering af ingen optionbuttons valgt
If OptionButton1.Value = False And OptionButton2.Value = False Then
dFunc.msgError = "V�lg venligst �n af svar mulighederne for at g� videre."
    SFunc.ShowFunc ("frmMsg")
    GoTo ending
End If

' Validering - Negative v�rdier
If TextBox1.Value < 0 Or TextBox2.Value < 0 Then
    dFunc.msgError = "Der kan ikke indtastes negative v�rdier i antal dage"
    SFunc.ShowFunc ("frmMsg")
    GoTo ending
End If

' Validering - gyldig tal v�rdi i antal dage tekstfelter
If IsNumeric(TextBox1.Value) = False And CheckBox1.Value = False Then
    dFunc.msgError = "Inds�t en gyldig v�rdi i antal dage"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Inds�t en gyldig v�rdi i antal dage")
    GoTo ending
End If

If IsNumeric(TextBox2.Value) = False And CheckBox2.Value = False Then
    dFunc.msgError = "Inds�t en gyldig v�rdi i antal dage"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Inds�t en gyldig v�rdi i antal dage")
    GoTo ending
End If

' Validering - Antal dage tekstboks felter skal udfyldes
If IsEmpty(TextBox1.Value) = True And CheckBox1.Value = False Then
    dFunc.msgError = "Inds�t en v�rdi i antal dage"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Inds�t en v�rdi i antal dage")
    GoTo ending
End If

If IsEmpty(TextBox2.Value) = True And CheckBox2.Value = False Then
    dFunc.msgError = "Inds�t en v�rdi i antal dage"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Inds�t en v�rdi i antal dage")
    GoTo ending
End If


' Validering p� mindst �n af OptionButtons skal v�re udfyldt
If OptionButton1.Value = False And OptionButton2.Value = False Then
    dFunc.msgError = "V�lg venligst hvor begyndelsestidspunktet for for�ldelsesfristen skal beregnes"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("V�lg venligst hvor begyndelsestidspunktet for for�ldelsesfristen skal beregnes")
    GoTo ending
End If

' Antal dage skrives i Varighed_X for periodeslutdatoen
If OptionButton1.Value = True And (CheckBox1.Value = False And CheckBox2.Value = False) Then
    Worksheets("Regler").Range("J60:J60").Value = CInt(TextBox2.Value) - CInt(TextBox1.Value)
    Worksheets("Regler").Range("J61:J61").Value = CInt(TextBox2.Value) - CInt(TextBox1.Value)
    Worksheets("Regler").Range("J62:J62").Value = CInt(TextBox2.Value) - CInt(TextBox1.Value)
    Worksheets("Regler").Range("J63:J63").Value = CInt(TextBox2.Value) - CInt(TextBox1.Value)
    Worksheets("Regler").Range("J71:J71").Value = CInt(TextBox2.Value) - CInt(TextBox1.Value)
ElseIf OptionButton2.Value = True And (CheckBox1.Value = False And CheckBox2.Value = False) Then
    Worksheets("Regler").Range("J60:J60").Value = CInt(TextBox1.Value) + CInt(TextBox2.Value)
    Worksheets("Regler").Range("J61:J61").Value = CInt(TextBox1.Value) + CInt(TextBox2.Value)
    Worksheets("Regler").Range("J62:J62").Value = CInt(TextBox1.Value) + CInt(TextBox2.Value)
    Worksheets("Regler").Range("J63:J63").Value = CInt(TextBox1.Value) + CInt(TextBox2.Value)
    Worksheets("Regler").Range("J71:J71").Value = CInt(TextBox2.Value) + CInt(TextBox1.Value)
End If

If CheckBox1.Value = False Or CheckBox2.Value = False Then
' Populations arket �ndres
Worksheets("Population").Range("B17:B17").Value = "NEJ"

' Reglerne aktiveres
Worksheets("Regler").Range("G60:G60").Value = "NEJ"
Worksheets("Regler").Range("G61:G61").Value = "NEJ"
Worksheets("Regler").Range("G62:G62").Value = "NEJ"
Worksheets("Regler").Range("G63:G63").Value = "NEJ"
Worksheets("Regler").Range("G71:G71").Value = "NEJ"

Else
' Populations arket �ndres
Worksheets("Population").Range("B17:B17").Value = "JA"

' Reglerne deaktiveres
Worksheets("Regler").Range("G60:G60").Value = "JA"
Worksheets("Regler").Range("G61:G61").Value = "JA"
Worksheets("Regler").Range("G62:G62").Value = "JA"
Worksheets("Regler").Range("G63:G63").Value = "JA"
Worksheets("Regler").Range("G71:G71").Value = "JA"

End If

' Grupper aktiveres
If OptionButton1.Value = True Or OptionButton2.Value = True Then
Worksheets("Gruppering").Range("C2:C2").Value = "JA"
End If

' Indsaet vaerdier i SpmSvar

If OptionButton1.Value = True Then
    Call writeSpmSvar("10.a_5", Controls("Label3").caption, Controls("OptionButton1").caption)
Else
    Call writeSpmSvar("10.a_5", Controls("Label3").caption, Controls("OptionButton2").caption)
End If


If TextBox1.Value <> "" Then
    'Worksheets("SpmSvar").Range("D72:D72").Value = CInt(TextBox1.Value)
    If OptionButton1.Value = True Then
        Call writeSpmSvar("10.a.1_5", Controls("Label1").caption, CInt(TextBox1.Value))
    ElseIf OptionButton2.Value = True Then
        Call writeSpmSvar("10.a.2_5", Controls("Label7").caption, CInt(TextBox1.Value))
    End If
End If
If TextBox2.Value <> "" Then
    'Worksheets("SpmSvar").Range("D73:D73").Value = CInt(TextBox2.Value)
    If OptionButton1.Value = True Then
        Call writeSpmSvar("10.a.1.1_5", Controls("Label4").caption, CInt(TextBox2.Value))
    ElseIf OptionButton2.Value = True Then
        Call writeSpmSvar("10.a.2.1_5", Controls("Label4").caption, CInt(TextBox2.Value))
    End If
End If

' "Ved ikke" skrives ned i arket
If CheckBox1.Value = True Then
    'Worksheets("SpmSvar").Range("D72:D72").Value = "Ved ikke"
    If OptionButton1.Value = True Then
        Call writeSpmSvar("10.a.1_5", Controls("Label1").caption, "Ved ikke")
    ElseIf OptionButton2.Value = True Then
        Call writeSpmSvar("10.a.2_5", Controls("Label7").caption, "Ved ikke")
    End If
End If

If CheckBox2.Value = True Then
    'Worksheets("SpmSvar").Range("D73:D73").Value = "Ved ikke"
    If OptionButton1.Value = True Then
        Call writeSpmSvar("10.a.1.1_5", Controls("Label4").caption, "Ved ikke")
    ElseIf OptionButton2.Value = True Then
        Call writeSpmSvar("10.a.2.1_5", Controls("Label4").caption, "Ved ikke")
    End If
End If


If CheckBox1.Value = True Or CheckBox2.Value = True Then
    dFunc.msgError = "RIM kan ikke beregne et tidligst muligt for�ldelsestidspunkt den del af populationen, hvor der ikke er indsendt FOKO. Den f�lgende konfiguration ang�r derfor kun fordringer, hvor der er indsendt FOKO"
    SFunc.ShowFunc ("frmMsg")
End If

' Tjek om gruppe 1 skal deaktiveres
If frm014.Forfaldsdato.Value = True And (frm028.CheckBox1.Value = True Or frm028.CheckBox2.Value = True) Then
    Worksheets("Gruppering").Range("C2:C2").Value = "NEJ"
End If
    
If frm014.SRB.Value = True And (frm032.CheckBox1.Value = True Or frm032.CheckBox2.Value = True) Then
    Worksheets("Gruppering").Range("C2:C2").Value = "NEJ"
End If
    
If frm014.Stiftelsesdato.Value = True And (frm029.CheckBox1.Value = True Or frm029.CheckBox2.Value = True) Then
    Worksheets("Gruppering").Range("C2:C2").Value = "NEJ"
End If
    
If frm014.PeriodeStartdato.Value = True And (frm030.CheckBox1.Value = True Or frm030.CheckBox2.Value = True) Then
    Worksheets("Gruppering").Range("C2:C2").Value = "NEJ"
End If
    
If frm014.PeriodeSlutdato.Value = True And (frm031.CheckBox1.Value = True Or frm031.CheckBox2.Value = True) Then
    Worksheets("Gruppering").Range("C2:C2").Value = "NEJ"
End If

    Me.Hide
    'store current form#
    recHis ("frm031")
    SFunc.ShowFunc ("frm039")
    'frm039.Show


ending:

End Sub

Private Sub OptionButton1_Click()
If OptionButton1.Value = True Then
    Label1.Visible = True
    Label2.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = False
    Label7.Visible = False
    CheckBox1.Visible = True
    CheckBox2.Visible = True
    CheckBox3.Visible = True
    TextBox1.Enabled = True
    TextBox2.Enabled = True
    TextBox1.Visible = True
    TextBox2.Visible = True
End If
If CheckBox3.Value = True Then
    TextBox2.Enabled = False
    CheckBox2.Enabled = False
End If
End Sub

Private Sub OptionButton2_Click()
If OptionButton2.Value = True Then
    Label1.Visible = False
    Label2.Visible = False
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    CheckBox1.Visible = True
    CheckBox2.Visible = True
    CheckBox3.Visible = True
    TextBox1.Enabled = True
    TextBox2.Enabled = True
    TextBox1.Visible = True
    TextBox2.Visible = True
End If
If CheckBox3.Value = True Then
    TextBox2.Enabled = False
    CheckBox2.Enabled = False
End If
End Sub


Public Sub Tilbage_Click()
Me.Hide
'go back to previously stored form#
Call goBack
'SFunc.ShowFunc ("frm014")
'frm014.Show
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeClip


If findPreviousAns(findTopSpm("F"), "10.a_5", 1) = "Samme dag eller senere end det valgte stamdatafelt" Then
    OptionButton2.Value = True
ElseIf findPreviousAns(findTopSpm("F"), "10.a_5", 1) = "F�r det valgte stamdatafelt" Then
    OptionButton1.Value = True
End If
If OptionButton1.Value Then
    If findPreviousAns(findTopSpm("F"), "10.a.1_5", 1) = "Ved ikke" Then
        CheckBox1.Value = True
    Else
        TextBox1.Value = findPreviousAns(findTopSpm("F"), "10.a.1_5", 1)
    End If
    If findPreviousAns(findTopSpm("F"), "10.a.1_5", 1) = "Ved ikke" Then
        CheckBox1.Value = True
    Else
        TextBox1.Value = findPreviousAns(findTopSpm("F"), "10.a.1_5", 1)
    End If
    If findPreviousAns(findTopSpm("F"), "10.a.1.1_5", 1) = "Ved ikke" Then
        CheckBox2.Value = True
    Else
        TextBox2.Value = findPreviousAns(findTopSpm("F"), "10.a.1.1_5", 1)
        If findPreviousAns(findTopSpm("F"), "10.a.1.1_5", 1) = "1095" Then
            CheckBox3.Value = True
            TextBox1.Enabled = False
        End If
    End If
    
ElseIf OptionButton2.Value Then
    If findPreviousAns(findTopSpm("F"), "10.a.2_5", 1) = "Ved ikke" Then
        CheckBox1.Value = True
    Else
        TextBox1.Value = findPreviousAns(findTopSpm("F"), "10.a.2_5", 1)
    End If
    If findPreviousAns(findTopSpm("F"), "10.a.2.1_5", 1) = "Ved ikke" Then
        CheckBox2.Value = True
    Else
        TextBox2.Value = findPreviousAns(findTopSpm("F"), "10.a.2.1_5", 1)
        If findPreviousAns(findTopSpm("F"), "10.a.2.1_5", 1) = "1095" Then
            CheckBox3.Value = True
            TextBox2.Enabled = False
        End If
    End If
End If

If OptionButton1.Value = False And OptionButton2.Value = False Then
    Label1.Visible = False
    Label2.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    CheckBox1.Visible = False
    CheckBox2.Visible = False
    CheckBox3.Visible = False
    TextBox1.Visible = False
    TextBox2.Visible = False
Else
    Label1.Visible = True
    Label2.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    CheckBox1.Visible = True
    CheckBox2.Visible = True
    CheckBox3.Visible = True
    TextBox1.Visible = True
    TextBox2.Visible = True
    CheckBox1.Visible = True
    CheckBox2.Visible = True
    TextBox1.Visible = True
    TextBox2.Visible = True
End If

    Label12.Font.size = 15
Call drawProgressBar(Me, Me.Name)
End Sub
