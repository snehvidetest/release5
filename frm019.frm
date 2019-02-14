VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm019 
   Caption         =   "Frasortering"
   ClientHeight    =   6936
   ClientLeft      =   60
   ClientTop       =   264
   ClientWidth     =   10980
   OleObjectBlob   =   "frm019.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm019"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public Sub OKButton_Click()
    ' Gem svar i "SpmSvar"
'    Worksheets("SpmSvar").Range("C45:C45").Value = Controls("Label1").caption
'    Worksheets("SpmSvar").Range("D45:D45").Value = CheckBox1.caption & " " & CheckBox1.Value
'    Worksheets("SpmSvar").Range("E45:E45").Value = "SRB" & " " & CheckBox2.Value
'    Worksheets("SpmSvar").Range("F45:F45").Value = CheckBox3.caption & " " & CheckBox3.Value
'    Worksheets("SpmSvar").Range("G45:G45").Value = "PeriodeStart" & " " & CheckBox4.Value
'    Worksheets("SpmSvar").Range("H45:H45").Value = "PeriodeSlut" & " " & CheckBox5.Value
Call writeSpmSvar("15.b", Controls("Label1").caption, "")
If CheckBox1.Value = True Then
    Call writeSpmSvar("15.b_1", Controls("CheckBox1").caption, "")
End If
If CheckBox2.Value = True Then
    Call writeSpmSvar("15.b_2", Controls("CheckBox2").caption, "")
End If
If CheckBox3.Value = True Then
    Call writeSpmSvar("15.b_3", Controls("CheckBox3").caption, "")
End If
If CheckBox4.Value = True Then
    Call writeSpmSvar("15.b_4", Controls("CheckBox4").caption, "")
End If
If CheckBox5.Value = True Then
    Call writeSpmSvar("15.b_5", Controls("CheckBox5").caption, "")
End If
    
If CheckBox1.Value = False Then
    Worksheets("Regler").Range("J29:J29").Value = "-1825"
    Worksheets("Regler").Range("M29:M29").Value = "-1"
ElseIf CheckBox2.Value = False Then
    Worksheets("Regler").Range("J30:J30").Value = "-1825"
    Worksheets("Regler").Range("M30:M30").Value = "-1"
ElseIf CheckBox3.Value = False Then
    Worksheets("Regler").Range("J31:J31").Value = "-1825"
    Worksheets("Regler").Range("M31:M31").Value = "-1"
ElseIf CheckBox4.Value = False Then
    Worksheets("Regler").Range("J32:J32").Value = "-1825"
    Worksheets("Regler").Range("M32:M32").Value = "-1"
ElseIf CheckBox5.Value = False Then
    Worksheets("Regler").Range("J33:J33").Value = "-1825"
    Worksheets("Regler").Range("M33:M33").Value = "-1"
End If
    
    
    ' Vælg form
    If (CheckBox1.Value = True Or CheckBox2.Value = True Or CheckBox3.Value = True Or CheckBox4.Value = True Or CheckBox5.Value = True) Then
        
        
        If CheckBox1.Value = True Then
            frm042.TextBox1.Enabled = True
            frm042.ComboBox1.Enabled = True
            frm042.Label8.ForeColor = RGB(0, 0, 0)
        ElseIf frm017.CheckBox1.Value = False Then
            frm042.TextBox1.Enabled = False
            frm042.ComboBox1.Enabled = False
            frm042.TextBox1.Value = ""
            frm042.ComboBox1.Value = ""
            frm042.Label8.ForeColor = RGB(169, 169, 169)
        End If
        
        If CheckBox2.Value = True Then
            frm042.TextBox2.Enabled = True
            frm042.ComboBox2.Enabled = True
            frm042.Label11.ForeColor = RGB(0, 0, 0)
        Else
            frm042.TextBox2.Enabled = False
            frm042.ComboBox2.Enabled = False
            frm042.TextBox2.Value = ""
            frm042.ComboBox2.Value = ""
            frm042.Label11.ForeColor = RGB(169, 169, 169)
        End If
        
        If CheckBox3.Value = True Then
            frm042.TextBox3.Enabled = True
            frm042.ComboBox3.Enabled = True
            frm042.Label10.ForeColor = RGB(0, 0, 0)
        Else
            frm042.TextBox3.Enabled = False
            frm042.ComboBox3.Enabled = False
            frm042.TextBox3.Value = ""
            frm042.ComboBox3.Value = ""
            frm042.Label10.ForeColor = RGB(169, 169, 169)
        End If
        
        If CheckBox4.Value = True Then
            frm042.TextBox4.Enabled = True
            frm042.ComboBox4.Enabled = True
            frm042.Label9.ForeColor = RGB(0, 0, 0)
        Else
            frm042.TextBox4.Enabled = False
            frm042.ComboBox4.Enabled = False
            frm042.TextBox4.Value = ""
            frm042.ComboBox4.Value = ""
            frm042.Label9.ForeColor = RGB(169, 169, 169)
        End If
    
        If CheckBox5.Value = True Then
            frm042.TextBox5.Enabled = True
            frm042.ComboBox5.Enabled = True
            frm042.Label12.ForeColor = RGB(0, 0, 0)
        Else
            frm042.TextBox5.Enabled = False
            frm042.ComboBox5.Enabled = False
            frm042.TextBox5.Value = ""
            frm042.ComboBox5.Value = ""
            frm042.Label12.ForeColor = RGB(169, 169, 169)
        End If
        
        Me.Hide
        'store current form#
        recHis ("frm019")
        SFunc.ShowFunc ("frm042")
        
        'frm042.Show
    Else
        dFunc.msgError = "Det skal overvejes, hvornår RIM vil tillade, at fordringer, der oprettes til modregning inden udløbet af de fem stamdatafelter, lukkes igennem FLEX-filteret."
        SFunc.ShowFunc ("frmMsg")
        'MsgBox ("Det skal overvejes, hvornår RIM vil tillade, at fordringer, der oprettes til modregning inden udløbet af de fem stamdatafelter, lukkes igennem FLEX-filteret.")
        Me.Hide
        'store current form#
        recHis ("frm019")
        SFunc.ShowFunc ("frm025")
        'frm025.Show
    End If



End Sub



Public Sub Tilbage_Click()

    Me.Hide
    'go back to previously stored form#
    Call goBack
    'SFunc.TestMode ("frm024")
    'frm024.Show

End Sub

Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeClip

    ' Indlæs tidligere svar
If findPreviousAns(findTopSpm("F"), "15.b_1", 0) <> "" Then
    CheckBox1.Value = True
End If
If findPreviousAns(findTopSpm("F"), "15.b_2", 0) <> "" Then
    CheckBox2.Value = True
End If
If findPreviousAns(findTopSpm("F"), "15.b_3", 0) <> "" Then
    CheckBox3.Value = True
End If
If findPreviousAns(findTopSpm("F"), "15.b_4", 0) <> "" Then
    CheckBox4.Value = True
End If
If findPreviousAns(findTopSpm("F"), "15.b_5", 0) <> "" Then
    CheckBox5.Value = True
End If

'    If Not IsEmpty(Worksheets("SpmSvar").Range("D45:D45").Value) Then
'
'        If vaArray = Split(Worksheets("SpmSvar").Range("D45:D45").Value, " ")(1) = True Then
'            CheckBox1.Value = True
'        End If
'
'        If vaArray = Split(Worksheets("SpmSvar").Range("E45:E45").Value, " ")(1) = True Then
'            CheckBox2.Value = True
'        End If
'
'        If vaArray = Split(Worksheets("SpmSvar").Range("F45:F45").Value, " ")(1) = True Then
'            CheckBox3.Value = True
'        End If
'
'        If vaArray = Split(Worksheets("SpmSvar").Range("G45:G45").Value, " ")(1) = True Then
'            CheckBox4.Value = True
'        End If
'
'        If vaArray = Split(Worksheets("SpmSvar").Range("H45:H45").Value, " ")(1) = True Then
'            CheckBox5.Value = True
'        End If
'    End If
Call drawProgressBar(Me, Me.Name)
End Sub
