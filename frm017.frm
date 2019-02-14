VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm017 
   Caption         =   "Frasortering"
   ClientHeight    =   6936
   ClientLeft      =   60
   ClientTop       =   276
   ClientWidth     =   10980
   OleObjectBlob   =   "frm017.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm017"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub OKButton_Click()
    ' Gem svar
'    Worksheets("SpmSvar").Range("C43:C43").Value = Controls("Label1").caption
'    Worksheets("SpmSvar").Range("D43:D43").Value = CheckBox1.caption & " " & CheckBox1.Value
'    Worksheets("SpmSvar").Range("E43:E43").Value = "SRB" & " " & CheckBox2.Value
'    Worksheets("SpmSvar").Range("F43:F43").Value = CheckBox3.caption & " " & CheckBox3.Value
'    Worksheets("SpmSvar").Range("G43:G43").Value = "PeriodeStart" & " " & CheckBox4.Value
'    Worksheets("SpmSvar").Range("H43:H43").Value = "PeriodeSlut" & " " & CheckBox5.Value
Call writeSpmSvar("14.b", Controls("Label1").caption, "")
If CheckBox1.Value = True Then
    Call writeSpmSvar("14.b_1", Controls("CheckBox1").caption, "")
End If
If CheckBox2.Value = True Then
    Call writeSpmSvar("14.b_2", Controls("CheckBox2").caption, "")
End If
If CheckBox3.Value = True Then
    Call writeSpmSvar("14.b_3", Controls("CheckBox3").caption, "")
End If
If CheckBox4.Value = True Then
    Call writeSpmSvar("14.b_4", Controls("CheckBox4").caption, "")
End If
If CheckBox5.Value = True Then
    Call writeSpmSvar("14.b_5", Controls("CheckBox5").caption, "")
End If

If CheckBox1.Value = False Then
    Worksheets("Regler").Range("J24:J24").Value = "-1825"
    Worksheets("Regler").Range("M24:M24").Value = "-1"
ElseIf CheckBox2.Value = False Then
    Worksheets("Regler").Range("J25:J25").Value = "-1825"
    Worksheets("Regler").Range("M25:M25").Value = "-1"
ElseIf CheckBox3.Value = False Then
    Worksheets("Regler").Range("J26:J26").Value = "-1825"
    Worksheets("Regler").Range("M26:M26").Value = "-1"
ElseIf CheckBox4.Value = False Then
    Worksheets("Regler").Range("J27:J27").Value = "-1825"
    Worksheets("Regler").Range("M27:M27").Value = "-1"
ElseIf CheckBox5.Value = False Then
    Worksheets("Regler").Range("J28:J28").Value = "-1825"
    Worksheets("Regler").Range("M28:M28").Value = "-1"
End If


If (CheckBox1.Value = True Or CheckBox2.Value = True Or CheckBox3.Value = True Or CheckBox4.Value = True Or CheckBox5.Value) Then

    
        If CheckBox1.Value = True Then
            frm041.TextBox1.Enabled = True
            frm041.ComboBox1.Enabled = True
            frm041.Label4.ForeColor = RGB(0, 0, 0)
        ElseIf frm017.CheckBox1.Value = False Then
            frm041.TextBox1.Enabled = False
            frm041.ComboBox1.Enabled = False
            frm041.TextBox1.Value = ""
            frm041.ComboBox1.Value = ""
            frm041.Label4.ForeColor = RGB(169, 169, 169)
        End If
        
        If CheckBox2.Value = True Then
            frm041.TextBox2.Enabled = True
            frm041.ComboBox2.Enabled = True
            frm041.Label5.ForeColor = RGB(0, 0, 0)
        Else
            frm041.TextBox2.Enabled = False
            frm041.ComboBox2.Enabled = False
            frm041.TextBox2.Value = ""
            frm041.ComboBox2.Value = ""
            frm041.Label5.ForeColor = RGB(169, 169, 169)
        End If
        
        If CheckBox3.Value = True Then
            frm041.TextBox3.Enabled = True
            frm041.ComboBox3.Enabled = True
            frm041.Label3.ForeColor = RGB(0, 0, 0)
        Else
            frm041.TextBox3.Enabled = False
            frm041.ComboBox3.Enabled = False
            frm041.TextBox3.Value = ""
            frm041.ComboBox3.Value = ""
            frm041.Label3.ForeColor = RGB(169, 169, 169)
        End If
        
        If CheckBox4.Value = True Then
            frm041.TextBox4.Enabled = True
            frm041.ComboBox4.Enabled = True
            frm041.Label2.ForeColor = RGB(0, 0, 0)
        Else
            frm041.TextBox4.Enabled = False
            frm041.ComboBox4.Enabled = False
            frm041.TextBox4.Value = ""
            frm041.ComboBox4.Value = ""
            frm041.Label2.ForeColor = RGB(169, 169, 169)
        End If
    
        If CheckBox5.Value = True Then
            frm041.TextBox5.Enabled = True
            frm041.ComboBox5.Enabled = True
            frm041.Label8.ForeColor = RGB(0, 0, 0)
        Else
            frm041.TextBox5.Enabled = False
            frm041.ComboBox5.Enabled = False
            frm041.TextBox5.Value = ""
            frm041.ComboBox5.Value = ""
            frm041.Label8.ForeColor = RGB(169, 169, 169)
        End If
    
    Me.Hide
    'store current form#
    recHis ("frm017")
    SFunc.ShowFunc ("frm041")
    
    'frm041.Show
ElseIf frm005.OptionButton1.Value = True Then
    Me.Hide
    'store current form#
    recHis ("frm017")
    dFunc.msgError = "Det skal overvejes, hvornår RIM vil tillade, at fordringer, der sendes til inddrivelse inden udløbet af de fem stamdatafelter, lukkes igennem FLEX-filteret."
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Det skal overvejes, hvornår RIM vil tillade, at fordringer, der sendes til inddrivelse inden udløbet af de fem stamdatafelter, lukkes igennem FLEX-filteret.")
    SFunc.ShowFunc ("frm024")
    'frm024.Show
ElseIf frm027.OptionButton1.Value = True Then
    Me.Hide
    'store current form#
    recHis ("frm017")
    dFunc.msgError = "Det skal overvejes, hvornår RIM vil tillade, at fordringer, der sendes til inddrivelse inden udløbet af de fem stamdatafelter, lukkes igennem FLEX-filteret."
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Det skal overvejes, hvornår RIM vil tillade, at fordringer, der sendes til inddrivelse inden udløbet af de fem stamdatafelter, lukkes igennem FLEX-filteret.")
    SFunc.ShowFunc ("frm025")
    'frm025.Show
End If



End Sub

Public Sub Tilbage_Click()
    Me.Hide
    'go back to previously stored form#
    Call goBack
    'ShowFunc ("frm023")
    'frm023.Show
End Sub

Private Sub UserForm_Initialize()

    Image1.PictureSizeMode = fmPictureSizeModeClip

    ' Indlæs tidligere svar
    
If findPreviousAns(findTopSpm("F"), "14.b_1", 0) <> "" Then
    CheckBox1.Value = True
End If
If findPreviousAns(findTopSpm("F"), "14.b_2", 0) <> "" Then
    CheckBox2 = True
End If
If findPreviousAns(findTopSpm("F"), "14.b_3", 0) <> "" Then
    CheckBox3.Value = True
End If
If findPreviousAns(findTopSpm("F"), "14.b_4", 0) <> "" Then
    CheckBox4.Value = True
End If
If findPreviousAns(findTopSpm("F"), "14.b_5", 0) <> "" Then
    CheckBox5.Value = True
End If

'    If Not IsEmpty(Worksheets("SpmSvar").Range("D43:D43").Value) Then
'
'        If vaArray = Split(Worksheets("SpmSvar").Range("D43:D43").Value, " ")(1) = True Then
'            CheckBox1.Value = True
'        End If
'
'        If vaArray = Split(Worksheets("SpmSvar").Range("E43:E43").Value, " ")(1) = True Then
'            CheckBox2.Value = True
'        End If
'
'        If vaArray = Split(Worksheets("SpmSvar").Range("F43:F43").Value, " ")(1) = True Then
'            CheckBox3.Value = True
'        End If
'
'        If vaArray = Split(Worksheets("SpmSvar").Range("G43:G43").Value, " ")(1) = True Then
'            CheckBox4.Value = True
'        End If
'
'        If vaArray = Split(Worksheets("SpmSvar").Range("H43:H43").Value, " ")(1) = True Then
'            CheckBox5.Value = True
'        End If
'    End If
Call drawProgressBar(Me, Me.Name)
End Sub
