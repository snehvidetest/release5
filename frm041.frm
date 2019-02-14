VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm041 
   ClientHeight    =   6936
   ClientLeft      =   60
   ClientTop       =   288
   ClientWidth     =   10980
   OleObjectBlob   =   "frm041.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm041"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub OKButton_Click()
    'Worksheets("SpmSvar").Range("C101:C101").Value = Controls("Label1").caption
Call writeSpmSvar("14.b.2", Controls("Label1").caption, "")
' indsæt forfaldsdato i Regler (INDR) og SpmSvar
    If IsNumeric(TextBox1.Value) Then
        Call writeSpmSvar("14.b.2_3", Controls("Label4").caption, TextBox1.Value, ComboBox1.Value)
    End If
    
    If ComboBox1.Value = "Dage" And IsNumeric(TextBox1.Value) Then
        Worksheets("Regler").Range("J24:J24").Value = "-1095"
        Worksheets("Regler").Range("M24:M24").Value = TextBox1.Value
'        Worksheets("SpmSvar").Range("D102:D102").Value = TextBox1.Value
'        Worksheets("SpmSvar").Range("E102:E102").Value = ComboBox1.Value
        
      
    ElseIf ComboBox1.Value = "Måneder" And IsNumeric(TextBox1.Value) Then
        Worksheets("Regler").Range("J24:J24").Value = "-1095"
        Worksheets("Regler").Range("N24:N24").Value = TextBox1.Value
'        Worksheets("SpmSvar").Range("D102:D102").Value = TextBox1.Value
'        Worksheets("SpmSvar").Range("E102:E102").Value = ComboBox1.Value
        
    ElseIf ComboBox1.Value = "År" And IsNumeric(TextBox1.Value) Then
        Worksheets("Regler").Range("J24:J24").Value = "-1095"
        Worksheets("Regler").Range("O24:O24").Value = TextBox1.Value
'        Worksheets("SpmSvar").Range("D102:D102").Value = TextBox1.Value
'        Worksheets("SpmSvar").Range("E102:E102").Value = ComboBox1.Value
    End If
    
        
' indsæt SRB i Regler (INDR) og SpmSvar
    If IsNumeric(TextBox2.Value) Then
        Call writeSpmSvar("14.b.2_4", Controls("Label5").caption, TextBox2.Value, ComboBox2.Value)
    End If
    
    If ComboBox2.Value = "Dage" And IsNumeric(TextBox2.Value) Then
        Worksheets("Regler").Range("J25:J25").Value = "-1095"
        Worksheets("Regler").Range("M25:M25").Value = TextBox2.Value
'        Worksheets("SpmSvar").Range("D103:D103").Value = TextBox2.Value
'        Worksheets("SpmSvar").Range("E103:E103").Value = ComboBox2.Value
      
    ElseIf ComboBox2.Value = "Måneder" And IsNumeric(TextBox2.Value) Then
        Worksheets("Regler").Range("J25:J25").Value = "-1095"
        Worksheets("Regler").Range("N25:N25").Value = TextBox2.Value
'        Worksheets("SpmSvar").Range("D103:D103").Value = TextBox2.Value
'        Worksheets("SpmSvar").Range("E103:E103").Value = ComboBox2.Value
        
    ElseIf ComboBox2.Value = "År" And IsNumeric(TextBox2.Value) Then
        Worksheets("Regler").Range("J25:J25").Value = "-1095"
        Worksheets("Regler").Range("O25:O25").Value = TextBox2.Value
'        Worksheets("SpmSvar").Range("D103:D103").Value = TextBox2.Value
'        Worksheets("SpmSvar").Range("E103:E103").Value = ComboBox2.Value
        
    End If
    
' indsæt stiftelse (INDR) i Regler og SpmSvar
    If IsNumeric(TextBox3.Value) Then
        Call writeSpmSvar("14.b.2_2", Controls("Label3").caption, TextBox3.Value, ComboBox3.Value)
    End If
    
  If ComboBox3.Value = "Dage" And IsNumeric(TextBox3.Value) Then
        Worksheets("Regler").Range("J26:J26").Value = "-1095"
        Worksheets("Regler").Range("M26:M26").Value = TextBox3.Value
'        Worksheets("SpmSvar").Range("D104:D104").Value = TextBox3.Value
'        Worksheets("SpmSvar").Range("E104:E104").Value = ComboBox3.Value
    
    ElseIf ComboBox3.Value = "Måneder" And IsNumeric(TextBox3.Value) Then
        Worksheets("Regler").Range("J26:J26").Value = "-1095"
        Worksheets("Regler").Range("N26:N26").Value = TextBox3.Value
'        Worksheets("SpmSvar").Range("D104:D104").Value = TextBox3.Value
'        Worksheets("SpmSvar").Range("E104:E104").Value = ComboBox3.Value
        
    ElseIf ComboBox3.Value = "År" And IsNumeric(TextBox3.Value) Then
        Worksheets("Regler").Range("J26:J26").Value = "-1095"
        Worksheets("Regler").Range("O26:O26").Value = TextBox3.Value
'        Worksheets("SpmSvar").Range("D104:D104").Value = TextBox3.Value
'        Worksheets("SpmSvar").Range("E104:E104").Value = ComboBox3.Value
    End If

' indsæt periodestart (INDR) i Regler og SpmSvar
    If IsNumeric(TextBox4.Value) Then
        Call writeSpmSvar("14.b.2_1", Controls("Label2").caption, TextBox4.Value, ComboBox4.Value)
    End If
    
If ComboBox4.Value = "Dage" And IsNumeric(TextBox4.Value) Then
        Worksheets("Regler").Range("J27:J27").Value = "-1095"
        Worksheets("Regler").Range("M27:M27").Value = TextBox4.Value
'        Worksheets("SpmSvar").Range("D105:D105").Value = TextBox4.Value
'        Worksheets("SpmSvar").Range("E105:E105").Value = ComboBox4.Value
    
    ElseIf ComboBox4.Value = "Måneder" And IsNumeric(TextBox4.Value) Then
        Worksheets("Regler").Range("J27:J27").Value = "-1095"
        Worksheets("Regler").Range("N27:N27").Value = TextBox4.Value
'        Worksheets("SpmSvar").Range("D105:D105").Value = TextBox4.Value
'        Worksheets("SpmSvar").Range("E105:E105").Value = ComboBox4.Value
        
    ElseIf ComboBox4.Value = "År" And IsNumeric(TextBox4.Value) Then
        Worksheets("Regler").Range("J27:J27").Value = "-1095"
        Worksheets("Regler").Range("O27:O27").Value = TextBox4.Value
'        Worksheets("SpmSvar").Range("D105:D105").Value = TextBox4.Value
'        Worksheets("SpmSvar").Range("E105:E105").Value = ComboBox4.Value
    End If

' indsæt periodeslut (INDR) i Regler og SpmSvar
    If IsNumeric(TextBox5.Value) Then
        Call writeSpmSvar("14.b.1", Controls("Label8").caption, TextBox5.Value, ComboBox5.Value)
    End If
    
If ComboBox5.Value = "Dage" And IsNumeric(TextBox5.Value) Then
        Worksheets("Regler").Range("J28:J28").Value = "-1095"
        Worksheets("Regler").Range("M28:M28").Value = TextBox5.Value
'        Worksheets("SpmSvar").Range("D111:D111").Value = TextBox5.Value
'        Worksheets("SpmSvar").Range("E111:E111").Value = ComboBox5.Value
    
    ElseIf ComboBox5.Value = "Måneder" And IsNumeric(TextBox5.Value) Then
        Worksheets("Regler").Range("J28:J28").Value = "-1095"
        Worksheets("Regler").Range("N28:N28").Value = TextBox5.Value
'        Worksheets("SpmSvar").Range("D111:D111").Value = TextBox5.Value
'        Worksheets("SpmSvar").Range("E111:E111").Value = ComboBox5.Value
        
    ElseIf ComboBox5.Value = "År" And IsNumeric(TextBox5.Value) Then
        Worksheets("Regler").Range("J28:J28").Value = "-1095"
        Worksheets("Regler").Range("O28:O28").Value = TextBox5.Value
'        Worksheets("SpmSvar").Range("D111:D111").Value = TextBox5.Value
'        Worksheets("SpmSvar").Range("E111:E111").Value = ComboBox5.Value
    End If


If TextBox1.Enabled = True And TextBox1.Value = "" Then
    dFunc.msgError = "Udfyld venligst antallet"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst antallet")
    GoTo ending
End If

If TextBox2.Enabled = True And TextBox2.Value = "" Then
    dFunc.msgError = "Udfyld venligst antallet"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst antallet")
    GoTo ending
End If

If TextBox3.Enabled = True And TextBox3.Value = "" Then
    dFunc.msgError = "Udfyld venligst antallet"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst antallet")
    GoTo ending
End If

If TextBox4.Enabled = True And TextBox4.Value = "" Then
    dFunc.msgError = "Udfyld venligst antallet"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst antallet")
    GoTo ending
End If

If TextBox5.Enabled = True And TextBox5.Value = "" Then
    dFunc.msgError = "Udfyld venligst antallet"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst antallet")
    GoTo ending
End If

If ComboBox1.Enabled = True And ComboBox1.Value = "" Then
    dFunc.msgError = "Udfyld venligst Dag/Måneder/År"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst Dag/Måneder/År")
    GoTo ending
End If

If ComboBox2.Enabled = True And ComboBox2.Value = "" Then
    dFunc.msgError = "Udfyld venligst Dag/Måneder/År"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst Dag/Måneder/År")
    GoTo ending
End If

If ComboBox3.Enabled = True And ComboBox3.Value = "" Then
    dFunc.msgError = "Udfyld venligst Dag/Måneder/År"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst Dag/Måneder/År")
    GoTo ending
End If

If ComboBox4.Enabled = True And ComboBox4.Value = "" Then
    dFunc.msgError = "Udfyld venligst Dag/Måneder/År"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst Dag/Måneder/År")
    GoTo ending
End If

If ComboBox5.Enabled = True And ComboBox5.Value = "" Then
    dFunc.msgError = "Udfyld venligst Dag/Måneder/År"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyld venligst Dag/Måneder/År")
    GoTo ending
End If

    If ComboBox1.Value = "" Then
        Worksheets("Regler").Range("G24:G24").Value = "NEJ"
    End If

    If ComboBox2.Value = "" Then
        Worksheets("Regler").Range("G25:G25").Value = "NEJ"
    End If

    If ComboBox3.Value = "" Then
        Worksheets("Regler").Range("G26:G26").Value = "NEJ"
    End If

    If ComboBox4.Value = "" Then
        Worksheets("Regler").Range("G27:G27").Value = "NEJ"
    End If

    If ComboBox5.Value = "" Then
        Worksheets("Regler").Range("G28:G28").Value = "NEJ"
    End If
    
    If ComboBox1.Value <> "" Then
        Worksheets("Regler").Range("G24:G24").Value = "JA"
    End If

    If ComboBox2.Value <> "" Then
        Worksheets("Regler").Range("G25:G25").Value = "JA"
    End If

    If ComboBox3.Value <> "" Then
        Worksheets("Regler").Range("G26:G26").Value = "JA"
    End If

    If ComboBox4.Value <> "" Then
        Worksheets("Regler").Range("G27:G27").Value = "JA"
    End If

    If ComboBox5.Value <> "" Then
        Worksheets("Regler").Range("G28:G28").Value = "JA"
    End If

If frm005.OptionButton1.Value = True Then
    dFunc.msgError = "Det skal overvejes, hvornår RIM vil tillade, at fordringer, der sendes til inddrivelse inden udløbet af de fem stamdatafelter, lukkes igennem FLEX-filteret."
    SFunc.ShowFunc ("frmMsg")
     Me.Hide
    'store current form#
    recHis ("frm041")
    'MsgBox ("Det skal overvejes, hvornår RIM vil tillade, at fordringer, der sendes til inddrivelse inden udløbet af de fem stamdatafelter, lukkes igennem FLEX-filteret.")
    SFunc.ShowFunc ("frm024")
    'frm024.Show
Else 'If frm027.OptionButton1.Value = True Then
    Me.Hide
    'store current form#
    recHis ("frm041")
    dFunc.msgError = "Det skal overvejes, hvornår RIM vil tillade, at fordringer, der sendes til inddrivelse inden udløbet af de fem stamdatafelter, lukkes igennem FLEX-filteret."
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Det skal overvejes, hvornår RIM vil tillade, at fordringer, der sendes til inddrivelse inden udløbet af de fem stamdatafelter, lukkes igennem FLEX-filteret.")
    SFunc.ShowFunc ("frm025")
    'frm025.Show
End If

ending:
End Sub

Public Sub Tilbage_Click()
'TextBox1.Value = ""
'TextBox2.Value = ""
'TextBox3.Value = ""
'TextBox4.Value = ""
'TextBox5.Value = ""
'
'ComboBox1.Value = ""
'ComboBox2.Value = ""
'ComboBox3.Value = ""
'ComboBox4.Value = ""
'ComboBox5.Value = ""
'
'Label2.ForeColor = RGB(0, 0, 0)
'Label3.ForeColor = RGB(0, 0, 0)
'Label4.ForeColor = RGB(0, 0, 0)
'Label5.ForeColor = RGB(0, 0, 0)
'Label8.ForeColor = RGB(0, 0, 0)



Me.Hide
'go back to previously stored form#
Call goBack
'SFunc.ShowFunc ("frm023")
'frm023.Show
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
Image1.PictureSizeMode = fmPictureSizeModeClip
    
    With ComboBox1
        .AddItem ("Dage")
        .AddItem ("Måneder")
        .AddItem ("År")
    End With
    
    With ComboBox2
        .AddItem ("Dage")
        .AddItem ("Måneder")
        .AddItem ("År")
    End With

    With ComboBox3
        .AddItem ("Dage")
        .AddItem ("Måneder")
        .AddItem ("År")
    End With
    
    With ComboBox4
        .AddItem ("Dage")
        .AddItem ("Måneder")
        .AddItem ("År")
    End With

    With ComboBox5
        .AddItem ("Dage")
        .AddItem ("Måneder")
        .AddItem ("År")
    End With
' Indlæs forfaldsdato dato fra tidligere besvarelse
    If findPreviousAns(findTopSpm("F"), "14.b.2_3", 2) = "Dage" Then
        TextBox1.Value = findPreviousAns(findTopSpm("F"), "14.b.2_3", 1)
        ComboBox1.Value = "Dage"
    
    ElseIf findPreviousAns(findTopSpm("F"), "14.b.2_3", 2) = "Måneder" Then
        TextBox1.Value = findPreviousAns(findTopSpm("F"), "14.b.2_3", 1)
        ComboBox1.Value = "Måneder"
        
    ElseIf findPreviousAns(findTopSpm("F"), "14.b.2_3", 2) = "År" Then
        TextBox1.Value = findPreviousAns(findTopSpm("F"), "14.b.2_3", 1)
        ComboBox1.Value = "År"
    End If

' Indlæs SRB dato fra tidligere besvarelse
    If findPreviousAns(findTopSpm("F"), "14.b.2_4", 2) = "Dage" Then
        TextBox2.Value = findPreviousAns(findTopSpm("F"), "14.b.2_4", 1)
        ComboBox2.Value = "Dage"
        
    ElseIf findPreviousAns(findTopSpm("F"), "14.b.2_4", 2) = "Måneder" Then
        TextBox2.Value = findPreviousAns(findTopSpm("F"), "14.b.2_4", 1)
        ComboBox2.Value = "Måneder"
        
    ElseIf findPreviousAns(findTopSpm("F"), "14.b.2_4", 2) = "År" Then
        TextBox2.Value = findPreviousAns(findTopSpm("F"), "14.b.2_4", 1)
        ComboBox2.Value = "År"
    End If
    
' Indlæs stiftelsesdato fra tidligere besvarelse
    If findPreviousAns(findTopSpm("F"), "14.b.2_2", 2) = "Dage" Then
        TextBox3.Value = findPreviousAns(findTopSpm("F"), "14.b.2_2", 1)
        ComboBox3.Value = "Dage"
        
    ElseIf findPreviousAns(findTopSpm("F"), "14.b.2_2", 2) = "Måneder" Then
        TextBox3.Value = findPreviousAns(findTopSpm("F"), "14.b.2_2", 1)
        ComboBox3.Value = "Måneder"
        
    ElseIf findPreviousAns(findTopSpm("F"), "14.b.2_2", 2) = "År" Then
        TextBox3.Value = findPreviousAns(findTopSpm("F"), "14.b.2_2", 1)
        ComboBox3.Value = "År"
    End If

' Indlæs periodestartdato fra tidligere besvarelse
    If findPreviousAns(findTopSpm("F"), "14.b.2_1", 2) = "Dage" Then
        TextBox4.Value = findPreviousAns(findTopSpm("F"), "14.b.2_1", 1)
        ComboBox4.Value = "Dage"
        
    ElseIf findPreviousAns(findTopSpm("F"), "14.b.2_1", 2) = "Måneder" Then
        TextBox4.Value = findPreviousAns(findTopSpm("F"), "14.b.2_1", 1)
        ComboBox4.Value = "Måneder"
        
    ElseIf findPreviousAns(findTopSpm("F"), "14.b.2_1", 2) = "År" Then
        TextBox4.Value = findPreviousAns(findTopSpm("F"), "14.b.2_1", 1)
        ComboBox4.Value = "År"
    End If
    
    ' Indlæs periodeslutdato fra tidligere besvarelse
    If findPreviousAns(findTopSpm("F"), "14.b.1", 2) = "Dage" Then
        TextBox5.Value = findPreviousAns(findTopSpm("F"), "14.b.1", 1)
        ComboBox5.Value = "Dage"
        
    ElseIf findPreviousAns(findTopSpm("F"), "14.b.1", 2) = "Måneder" Then
        TextBox5.Value = findPreviousAns(findTopSpm("F"), "14.b.1", 1)
        ComboBox5.Value = "Måneder"
        
    ElseIf findPreviousAns(findTopSpm("F"), "14.b.1", 2) = "År" Then
        TextBox5.Value = findPreviousAns(findTopSpm("F"), "14.b.1", 1)
        ComboBox5.Value = "År"
    End If
    

    If frm017.CheckBox1.Value = True Then
        TextBox1.Enabled = True
        ComboBox1.Enabled = True
    ElseIf frm017.CheckBox1.Value = False Then
        TextBox1.Enabled = False
        ComboBox1.Enabled = False
        TextBox1.Value = ""
        ComboBox1.Value = ""
        Label4.ForeColor = RGB(169, 169, 169)
    End If
    
    If frm017.CheckBox2.Value = True Then
        TextBox2.Enabled = True
        ComboBox2.Enabled = True
    Else
        TextBox2.Enabled = False
        ComboBox2.Enabled = False
        TextBox2.Value = ""
        ComboBox2.Value = ""
        Label5.ForeColor = RGB(169, 169, 169)
    End If
    
    If frm017.CheckBox3.Value = True Then
        TextBox3.Enabled = True
        ComboBox3.Enabled = True
    Else
        TextBox3.Enabled = False
        ComboBox3.Enabled = False
        TextBox3.Value = ""
        ComboBox3.Value = ""
        Label3.ForeColor = RGB(169, 169, 169)
    End If
    
    If frm017.CheckBox4.Value = True Then
        TextBox4.Enabled = True
        ComboBox4.Enabled = True
    Else
        TextBox4.Enabled = False
        ComboBox4.Enabled = False
        TextBox4.Value = ""
        ComboBox4.Value = ""
        Label2.ForeColor = RGB(169, 169, 169)
    End If

    If frm017.CheckBox5.Value = True Then
        TextBox5.Enabled = True
        ComboBox5.Enabled = True
    Else
        TextBox5.Enabled = False
        ComboBox5.Enabled = False
        TextBox5.Value = ""
        ComboBox5.Value = ""
        Label8.ForeColor = RGB(169, 169, 169)
    End If

Call drawProgressBar(Me, Me.Name)
End Sub
