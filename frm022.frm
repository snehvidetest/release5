VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm022 
   Caption         =   "Indledende spørgsmål"
   ClientHeight    =   10932
   ClientLeft      =   60
   ClientTop       =   264
   ClientWidth     =   10980
   OleObjectBlob   =   "frm022.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm022"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




' Dag i måneden

Private Sub CheckBox1_Click()
    If CheckBox1.Value = True Then
        'Datotype
        Forfaldsdato.Enabled = True
        SRB.Enabled = True
        Stiftelsesdato.Enabled = True
        PeriodeStartdato.Enabled = True
        PeriodeSlutdato.Enabled = True
        
    ElseIf CheckBox1.Value = False Then
        'Datotype
        Forfaldsdato.Enabled = False
        SRB.Enabled = False
        Stiftelsesdato.Enabled = False
        PeriodeStartdato.Enabled = False
        PeriodeSlutdato.Enabled = False
        
        Forfaldsdato.Value = False
        SRB.Value = False
        Stiftelsesdato.Value = False
        PeriodeStartdato.Value = False
        PeriodeSlutdato.Value = False
        
        'TxtBoxes
        txtFFStart.Enabled = False
        txtFFSlut.Enabled = False
        txtSRBstart.Enabled = False
        txtSRBslut.Enabled = False
        txtSTIstart.Enabled = False
        txtSTIslut.Enabled = False
        txtPSTstart.Enabled = False
        txtPSTslut.Enabled = False
        txtPSLstart.Enabled = False
        txtPSLslut.Enabled = False
        
        'EOM
        CheckBox4.Enabled = False
        CheckBox5.Enabled = False
        CheckBox6.Enabled = False
        CheckBox7.Enabled = False
        CheckBox8.Enabled = False
        
        'TxtBoxes
        txtFFStart.Value = ""
        txtFFSlut.Value = ""
        txtSRBstart.Value = ""
        txtSRBslut.Value = ""
        txtSTIstart.Value = ""
        txtSTIslut.Value = ""
        txtPSTstart.Value = ""
        txtPSTslut.Value = ""
        txtPSLstart.Value = ""
        txtPSLslut.Value = ""
        
        'EOM
        CheckBox4.Value = False
        CheckBox5.Value = False
        CheckBox6.Value = False
        CheckBox7.Value = False
        CheckBox8.Value = False
        
        
    End If
    
End Sub


Private Sub CheckBox3_Click()
If CheckBox3.Value = True Then
    CheckBox1.Enabled = False
    CheckBox2.Enabled = False
    CheckBox1.Value = False
    CheckBox2.Value = False
    txtFFStart.Value = ""
    txtFFSlut.Value = ""
    txtSRBstart.Value = ""
    txtSRBslut.Value = ""
    txtSTIstart.Value = ""
    txtSTIslut.Value = ""
    txtPSTstart.Value = ""
    txtPSTslut.Value = ""
    txtPSLstart.Value = ""
    txtPSLslut.Value = ""
    
    TextBox1.Value = ""
    TextBox2.Value = ""
    TextBox3.Value = ""
    TextBox4.Value = ""
    TextBox5.Value = ""
    TextBox6.Value = ""
    TextBox7.Value = ""
    TextBox8.Value = ""
    TextBox9.Value = ""
    TextBox10.Value = ""

    CheckBox4.Value = False
    CheckBox5.Value = False
    CheckBox6.Value = False
    CheckBox7.Value = False
    CheckBox8.Value = False
    
End If

If CheckBox3.Value = False Then
    CheckBox1.Enabled = True
    CheckBox2.Enabled = True
End If

End Sub

Private Sub Forfaldsdato_Click()
    If Forfaldsdato.Value = True Then
        txtFFStart.Enabled = True
        txtFFSlut.Enabled = True
        CheckBox4.Enabled = True
    ElseIf Forfaldsdato.Value = False Then
        txtFFStart.Enabled = False
        txtFFSlut.Enabled = False
        CheckBox4.Enabled = False
        CheckBox4.Value = False
    End If
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Public Sub OKButton_Click()

' Validering - tjek at dag er valid

Dim msg As String
msg = "Dag ikke udfyldt korrekt"

If check_day_month(txtFFStart, msg, "1") Then
    GoTo ending
End If

If check_day_month(txtFFSlut, msg, "1") Then
    GoTo ending
End If

If check_day_month(txtSRBstart, msg, "1") Then
    GoTo ending
End If

If check_day_month(txtSRBslut, msg, "1") Then
    GoTo ending
End If

If check_day_month(txtSTIstart, msg, "1") Then
    GoTo ending
End If

If check_day_month(txtSTIslut, msg, "1") Then
    GoTo ending
End If

If check_day_month(txtPSTstart, msg, "1") Then
    GoTo ending
End If

If check_day_month(txtPSTslut, msg, "1") Then
    GoTo ending
End If

If check_day_month(txtPSLstart, msg, "1") Then
    GoTo ending
End If

If check_day_month(txtPSLslut, msg, "1") Then
    GoTo ending
End If

' Validering - tjek at måned er valid

msg = "Måned ikke udfyldt korrekt"

If check_day_month(TextBox1, msg, "2") Then
   GoTo ending
End If

If check_day_month(TextBox2, msg, "2") Then
    GoTo ending
End If

If check_day_month(TextBox3, msg, "2") Then
    GoTo ending
End If

If check_day_month(TextBox4, msg, "2") Then
   GoTo ending
End If

If check_day_month(TextBox5, msg, "2") Then
   GoTo ending
End If

If check_day_month(TextBox6, msg, "2") Then
    GoTo ending
End If

If check_day_month(TextBox7, msg, "2") Then
    GoTo ending
End If

If check_day_month(TextBox8, msg, "2") Then
    GoTo ending
End If

If check_day_month(TextBox9, msg, "2") Then
    GoTo ending
   
End If

If check_day_month(TextBox10, msg, "2") Then
    GoTo ending
End If

If CheckBox1.Value = True And (Forfaldsdato.Value = False And SRB.Value = False And Stiftelsesdato.Value = False And PeriodeStartdato = False And PeriodeSlutdato = False) Then
    'dFunc.FOKO_Retracer = "Vælg mindst én af datotyperne"
    'SFunc.ShowFunc ("frmMsg")
    dFunc.msgError = "Vælg mindst én af datotyperne"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Vælg mindst én af datotyperne")
    GoTo ending
End If

If CheckBox2.Value = True And (CheckBox10.Value = False And CheckBox11.Value = False And CheckBox12.Value = False And CheckBox13.Value = False And CheckBox14.Value = False) Then
    'dFunc.FOKO_Retracer = "Vælg mindst én af datotyperne"
    'SFunc.ShowFunc ("frmMsg")
    dFunc.msgError = "Vælg mindst én af datotyperne"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Vælg mindst én af datotyperne")
    GoTo ending
End If

If Forfaldsdato.Value = True And (txtFFStart.Value = "" And txtFFSlut.Value = "" And CheckBox4.Value = False) Then
    'dFunc.FOKO_Retracer = "Vælg mindst én af datotyperne"
    'SFunc.ShowFunc ("frmMsg")
    dFunc.msgError = "Vælg mindst én af datotyperne"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Vælg mindst én dag i måneden")
    GoTo ending
End If

If SRB.Value = True And (txtSRBstart.Value = "" And txtSRBslut.Value = "" And CheckBox5.Value = False) Then
    'dFunc.FOKO_Retracer = "Vælg mindst én af datotyperne"
    'SFunc.ShowFunc ("frmMsg")
    dFunc.msgError = "Vælg mindst én af datotyperne"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Vælg mindst én dag i måneden")
    GoTo ending
End If

If Stiftelsesdato.Value = True And (txtSTIstart.Value = "" And txtSTIslut.Value = "" And CheckBox6.Value = False) Then
    'dFunc.FOKO_Retracer = "Vælg mindst én af datotyperne"
    'SFunc.ShowFunc ("frmMsg")
    dFunc.msgError = "Vælg mindst én af datotyperne"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Vælg mindst én dag i måneden")
    GoTo ending
End If

If PeriodeStartdato.Value = True And (txtPSTstart.Value = "" And txtPSTslut.Value = "" And CheckBox7.Value = False) Then
    'dFunc.FOKO_Retracer = "Vælg mindst én af datotyperne"
    'SFunc.ShowFunc ("frmMsg")
    dFunc.msgError = "Vælg mindst én af datotyperne"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Vælg mindst én dag i måneden")
    GoTo ending
End If

If PeriodeSlutdato.Value = True And (txtPSLstart.Value = "" And txtPSLslut.Value = "" And CheckBox8.Value = False) Then
    'dFunc.FOKO_Retracer = "Vælg mindst én af datotyperne"
    'SFunc.ShowFunc ("frmMsg")
    dFunc.msgError = "Vælg mindst én af datotyperne"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Vælg mindst én dag i måneden")
    GoTo ending
End If

If CheckBox10.Value = True And (TextBox1.Value = "" And TextBox2.Value = "") Then
    'dFunc.FOKO_Retracer = "Vælg mindst én af datotyperne"
    'SFunc.ShowFunc ("frmMsg")
    dFunc.msgError = "Vælg mindst én af datotyperne"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Vælg mindst én måned i året")
    GoTo ending
End If

If CheckBox11.Value = True And (TextBox3.Value = "" And TextBox4.Value = "") Then
    dFunc.msgError = "Vælg mindst én måned i året"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Vælg mindst én måned i året")
    GoTo ending
End If

If CheckBox12.Value = True And (TextBox5.Value = "" And TextBox6.Value = "") Then
    dFunc.msgError = "Vælg mindst én måned i året"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Vælg mindst én måned i året")
    GoTo ending
End If

If CheckBox13.Value = True And (TextBox7.Value = "" And TextBox8.Value = "") Then
    dFunc.msgError = "Vælg mindst én måned i året"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Vælg mindst én måned i året")
    GoTo ending
End If

If CheckBox14.Value = True And (TextBox9.Value = "" And TextBox10.Value = "") Then
    dFunc.msgError = "Vælg mindst én måned i året"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Vælg mindst én måned i året")
    GoTo ending
End If

' SLUT på validering







'Worksheets("SpmSvar").Range("C30:C30").Value = Controls("Label1").caption
        
If CheckBox1.Value = True Then
    'Worksheets("SpmSvar").Range("D30:D30").Value = "Samme dag i måneden"
End If

If CheckBox2.Value = True Then
    'Worksheets("SpmSvar").Range("D30:D30").Value = "Samme måned i året"
End If

If CheckBox3.Value = True Then
    'Worksheets("SpmSvar").Range("D30:D30").Value = "Nej/Ved ikke"
    Call writeSpmSvar("13", Controls("Label1").caption, Controls("CheckBox3").caption)
End If

' Data overførsel Samme dag i måneden

'Worksheets("SpmSvar").Range("C31:C31").Value = Controls("Label21").caption
If CheckBox1.Value = True Then
    Call writeSpmSvar("13.a", Controls("Label21").caption, "")
End If

If Forfaldsdato.Value = True Then
'    Worksheets("SpmSvar").Range("D31:D31").Value = "Forfaldsdato"
'    Worksheets("SpmSvar").Range("E31:E31").Value = txtFFStart
'    Worksheets("SpmSvar").Range("F31:F31").Value = txtFFSlut
    If CheckBox4 = True Then
        If txtFFStart = "" Then
            Call writeSpmSvar("13.a_1", Controls("Forfaldsdato").caption, "Sidste dag i måneden")
        Else
            Call writeSpmSvar("13.a_1", Controls("Forfaldsdato").caption, txtFFStart, "Sidste dag i måneden")
        End If
    Else
        Call writeSpmSvar("13.a_1", Controls("Forfaldsdato").caption, txtFFStart, txtFFSlut)
    End If
End If


If SRB.Value = True Then
'Worksheets("SpmSvar").Range("D32:D32").Value = "SRB Dato"
'Worksheets("SpmSvar").Range("E32:E32").Value = txtSRBstart
'Worksheets("SpmSvar").Range("F32:F32").Value = txtSRBslut
    If CheckBox5 = True Then
        If txtSRBstart = "" Then
            Call writeSpmSvar("13.a_2", Controls("SRB").caption, "Sidste dag i måneden")
        Else
            Call writeSpmSvar("13.a_2", Controls("SRB").caption, txtSRBstart, "Sidste dag i måneden")
        End If
    Else
        Call writeSpmSvar("13.a_2", Controls("SRB").caption, txtSRBstart, txtSRBslut)
    End If
End If

If Stiftelsesdato.Value = True Then
'Worksheets("SpmSvar").Range("D33:D33").Value = "Stiftelsesdato"
'Worksheets("SpmSvar").Range("E33:E33").Value = txtSTIstart
'Worksheets("SpmSvar").Range("F33:F33").Value = txtSTIslut
    If CheckBox6 = True Then
        If txtSTIstart = "" Then
            Call writeSpmSvar("13.a_3", Controls("Stiftelsesdato").caption, "Sidste dag i måneden")
        Else
            Call writeSpmSvar("13.a_3", Controls("Stiftelsesdato").caption, txtSTIstart, "Sidste dag i måneden")
        End If
    Else
        Call writeSpmSvar("13.a_3", Controls("Stiftelsesdato").caption, txtSTIstart, txtSTIslut)
    End If
End If

If PeriodeStartdato.Value = True Then
'Worksheets("SpmSvar").Range("D34:D34").Value = "PeriodeStartdato"
'Worksheets("SpmSvar").Range("E34:E34").Value = txtPSTstart
'Worksheets("SpmSvar").Range("F34:F34").Value = txtPSTslut
    If CheckBox7 = True Then
        If txtPSTstart = "" Then
            Call writeSpmSvar("13.a_4", Controls("PeriodeStartdato").caption, "Sidste dag i måneden")
        Else
            Call writeSpmSvar("13.a_4", Controls("PeriodeStartdato").caption, txtPSTstart, "Sidste dag i måneden")
        End If
    Else
        Call writeSpmSvar("13.a_4", Controls("PeriodeStartdato").caption, txtPSTstart, txtPSTslut)
    End If
End If

If PeriodeSlutdato.Value = True Then
'Worksheets("SpmSvar").Range("D35:D35").Value = "PeriodeSlutdato"
'Worksheets("SpmSvar").Range("E35:E35").Value = txtPSLstart
'Worksheets("SpmSvar").Range("F35:F35").Value = txtPSLslut
    If CheckBox8 = True Then
        If txtPSLstart = "" Then
            Call writeSpmSvar("13.a_5", Controls("PeriodeSlutdato").caption, "Sidste dag i måneden")
        Else
            Call writeSpmSvar("13.a_5", Controls("PeriodeSlutdato").caption, txtPSLstart, "Sidste dag i måneden")
        End If
    Else
        Call writeSpmSvar("13.a_5", Controls("PeriodeSlutdato").caption, txtPSLstart, txtPSLslut)
    End If
End If


' data overførsel "Samme måned i året"
If CheckBox2.Value = True Then
    Call writeSpmSvar("13.b", Controls("Label22").caption, "")
End If
'Worksheets("SpmSvar").Range("C36:C36").Value = Controls("Label22").caption

If CheckBox10.Value = True Then
    'Worksheets("SpmSvar").Range("D36:D36").Value = "Forfaldsdato"
    'Worksheets("SpmSvar").Range("E36:E36").Value = TextBox1
    'Worksheets("SpmSvar").Range("F36:F36").Value = TextBox2
    If TextBox1 = "" Then
        Call writeSpmSvar("13.b_1", Controls("CheckBox10").caption, TextBox2)
    Else
        Call writeSpmSvar("13.b_1", Controls("CheckBox10").caption, TextBox1, TextBox2)
    End If
End If

If CheckBox11.Value = True Then
    'Worksheets("SpmSvar").Range("D37:D37").Value = "SRB Dato"
    'Worksheets("SpmSvar").Range("E37:E37").Value = TextBox3
    'Worksheets("SpmSvar").Range("F37:F37").Value = TextBox4
    If TextBox3 = "" Then
        Call writeSpmSvar("13.b_2", Controls("CheckBox11").caption, TextBox4)
    Else
        Call writeSpmSvar("13.b_2", Controls("CheckBox11").caption, TextBox3, TextBox4)
    End If
End If

If CheckBox12.Value = True Then
    'Worksheets("SpmSvar").Range("D38:D38").Value = "Stiftelsesdato"
    'Worksheets("SpmSvar").Range("E38:E38").Value = TextBox5
    'Worksheets("SpmSvar").Range("F38:F38").Value = TextBox6
    If TextBox5 = "" Then
        Call writeSpmSvar("13.b_3", Controls("CheckBox12").caption, TextBox6)
    Else
        Call writeSpmSvar("13.b_3", Controls("CheckBox12").caption, TextBox5, TextBox6)
    End If
End If

If CheckBox13.Value = True Then
    'Worksheets("SpmSvar").Range("D39:D39").Value = "PeriodeStartdato"
    'Worksheets("SpmSvar").Range("E39:E39").Value = TextBox7
    'Worksheets("SpmSvar").Range("F39:F39").Value = TextBox8
    If TextBox7 = "" Then
        Call writeSpmSvar("13.b_4", Controls("CheckBox13").caption, TextBox8)
    Else
        Call writeSpmSvar("13.b_4", Controls("CheckBox13").caption, TextBox7, TextBox8)
    End If
End If

If CheckBox14.Value = True Then
    'Worksheets("SpmSvar").Range("D40:D40").Value = "PeriodeSlutdato"
    'Worksheets("SpmSvar").Range("E40:E40").Value = TextBox9
    'Worksheets("SpmSvar").Range("F40:F40").Value = TextBox10
    Call writeSpmSvar("13.b_5", Controls("CheckBox14").caption, TextBox9, TextBox10)
    If TextBox9 = "" Then
        Call writeSpmSvar("13.b_5", Controls("CheckBox14").caption, TextBox10)
    Else
        Call writeSpmSvar("13.b_5", Controls("CheckBox14").caption, TextBox9, TextBox10)
    End If
End If


Dim fields As Variant
fields = Array("P77:P78", "P81:P82", "P85:P86", "P89:P90", "P93:P94", "Q79:Q80", "Q83:Q84", "Q87:Q88", "Q91:Q92", "Q95:Q96")
For i = 0 To 9
    Call Insert_to_sheet("Regler", fields(i), "")
Next

If Forfaldsdato.Value = True Then
    Call Insert_to_sheet("Regler", "P77:P77", txtFFStart.Value)
    If CheckBox4.Value = True Then
        Call Insert_to_sheet("Regler", "P78:P78", "EOM")
    Else
        Call Insert_to_sheet("Regler", "P78:P78", txtFFSlut.Value)
    End If
End If

If SRB.Value = True Then
    Call Insert_to_sheet("Regler", "P85:P85", txtSRBstart.Value)
    If CheckBox5.Value = True Then
        Call Insert_to_sheet("Regler", "P86:P86", "EOM")
    Else
        Call Insert_to_sheet("Regler", "P86:P86", txtSRBslut.Value)
    End If
End If

If Stiftelsesdato.Value = True Then
    Call Insert_to_sheet("Regler", "P81:P81", txtSTIstart.Value)
    If CheckBox6.Value = True Then
        Call Insert_to_sheet("Regler", "P82:P82", "EOM")
    Else
        Call Insert_to_sheet("Regler", "P82:P82", txtSTIslut.Value)
    End If
End If

If PeriodeStartdato.Value = True Then
    Call Insert_to_sheet("Regler", "P89:P89", txtPSTstart.Value)
    If CheckBox7.Value = True Then
        Call Insert_to_sheet("Regler", "P90:P90", "EOM")
    Else
        Call Insert_to_sheet("Regler", "P90:P90", txtPSTslut.Value)
    End If
End If

If PeriodeSlutdato.Value = True Then
    Call Insert_to_sheet("Regler", "P93:P93", txtPSLstart.Value)
    If CheckBox8.Value = True Then
        Call Insert_to_sheet("Regler", "P94:P94", "EOM")
    Else
        Call Insert_to_sheet("Regler", "P94:P94", txtPSLslut.Value)
    End If
End If

' Måned i året

If CheckBox10.Value = True Then
    Call Insert_to_sheet("Regler", "Q79:Q79", TextBox1.Value)
    Call Insert_to_sheet("Regler", "Q80:Q80", TextBox2.Value)
End If

If CheckBox11.Value = True Then
    Call Insert_to_sheet("Regler", "Q87:Q87", TextBox3.Value)
    Call Insert_to_sheet("Regler", "Q88:Q88", TextBox4.Value)
End If

If CheckBox12.Value = True Then
    Call Insert_to_sheet("Regler", "Q83:Q83", TextBox5.Value)
    Call Insert_to_sheet("Regler", "Q84:Q84", TextBox6.Value)
End If

If CheckBox13.Value = True Then
    Call Insert_to_sheet("Regler", "Q91:Q91", TextBox7.Value)
    Call Insert_to_sheet("Regler", "Q92:Q92", TextBox8.Value)
End If

If CheckBox14.Value = True Then
    Call Insert_to_sheet("Regler", "Q95:Q95", TextBox9.Value)
    Call Insert_to_sheet("Regler", "Q96:Q96", TextBox10.Value)
End If

If CheckBox1.Value = False And CheckBox2.Value = False And CheckBox3.Value = False Then
    dFunc.msgError = "Udfyldt venligst datofelter, eller 'Nej/Ved ikke'"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Udfyldt venligst datofelter, eller 'Nej/Ved ikke'")
    GoTo ending
End If

Me.Hide
'store current form#
recHis ("frm022")
SFunc.ShowFunc ("frm023")

ending:
End Sub

Private Sub SRB_Click()
    If SRB.Value = True Then
        txtSRBstart.Enabled = True
        txtSRBslut.Enabled = True
        CheckBox5.Enabled = True
    ElseIf SRB.Value = False Then
        txtSRBstart.Enabled = False
        txtSRBslut.Enabled = False
        CheckBox5.Enabled = False
        CheckBox5.Value = False
    End If
End Sub

Private Sub Stiftelsesdato_Click()
    If Stiftelsesdato.Value = True Then
        txtSTIstart.Enabled = True
        txtSTIslut.Enabled = True
        CheckBox6.Enabled = True
    ElseIf Stiftelsesdato.Value = False Then
        txtSTIstart.Enabled = False
        txtSTIslut.Enabled = False
        CheckBox6.Enabled = False
        CheckBox6.Value = False
    End If
End Sub

Private Sub PeriodeStartdato_Click()
    If PeriodeStartdato.Value = True Then
        txtPSTstart.Enabled = True
        txtPSTslut.Enabled = True
        CheckBox7.Enabled = True
    ElseIf PeriodeStartdato.Value = False Then
        txtPSTstart.Enabled = False
        txtPSTslut.Enabled = False
        CheckBox7.Enabled = False
        CheckBox7.Value = False
    End If
End Sub

Private Sub PeriodeSlutdato_Click()
    If PeriodeSlutdato.Value = True Then
        txtPSLstart.Enabled = True
        txtPSLslut.Enabled = True
        CheckBox8.Enabled = True
    ElseIf PeriodeSlutdato.Value = False Then
        txtPSLstart.Enabled = False
        txtPSLslut.Enabled = False
        CheckBox8.Enabled = False
        CheckBox8.Value = False
    End If
End Sub


Private Sub CheckBox4_Click()
    If CheckBox4.Value = True Then
        txtFFSlut.Enabled = False
        txtFFSlut.Value = ""
    ElseIf CheckBox4.Value = False Then
        txtFFSlut.Enabled = True
        txtFFSlut.Value = ""
    End If
End Sub
Private Sub CheckBox5_Click()
    If CheckBox5.Value = True Then
        txtSRBslut.Enabled = False
        txtSRBslut.Value = ""
    ElseIf CheckBox5.Value = False Then
        txtSRBslut.Enabled = True
        txtSRBslut.Value = ""
    End If

End Sub
Private Sub CheckBox6_Click()
    If CheckBox6.Value = True Then
        txtSTIslut.Enabled = False
        txtSTIslut.Value = ""
    ElseIf CheckBox6.Value = False Then
        txtSTIslut.Enabled = True
        txtSTIslut.Value = ""
    End If

End Sub
Private Sub CheckBox7_Click()
    If CheckBox7.Value = True Then
        txtPSTslut.Enabled = False
        txtPSTslut.Value = ""
    ElseIf CheckBox7.Value = False Then
        txtPSTslut.Enabled = True
        txtPSTslut.Value = ""
    End If

End Sub
Private Sub CheckBox8_Click()
    If CheckBox8.Value = True Then
        txtPSLslut.Enabled = False
        txtPSLslut.Value = ""
    ElseIf CheckBox8.Value = False Then
        txtPSLslut.Enabled = True
        txtPSLslut.Value = ""
    End If

End Sub

' Måned i året


Private Sub CheckBox2_Click()
    If CheckBox2.Value = True Then
        'Datotype
        CheckBox10.Enabled = True
        CheckBox11.Enabled = True
        CheckBox12.Enabled = True
        CheckBox13.Enabled = True
        CheckBox14.Enabled = True
        
    ElseIf CheckBox2.Value = False Then
        'Datotype
        CheckBox10.Enabled = False
        CheckBox11.Enabled = False
        CheckBox12.Enabled = False
        CheckBox13.Enabled = False
        CheckBox14.Enabled = False
        
        'TxtBoxes
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        TextBox10.Enabled = False


        'Datotype
        CheckBox10.Value = False
        CheckBox11.Value = False
        CheckBox12.Value = False
        CheckBox13.Value = False
        CheckBox14.Value = False
        
        'TxtBoxes
        TextBox1.Value = ""
        TextBox2.Value = ""
        TextBox3.Value = ""
        TextBox4.Value = ""
        TextBox5.Value = ""
        TextBox6.Value = ""
        TextBox7.Value = ""
        TextBox8.Value = ""
        TextBox9.Value = ""
        TextBox10.Value = ""
        
    End If
    
End Sub

Private Sub CheckBox10_Click()
    If CheckBox10.Value = True Then
        TextBox1.Enabled = True
        TextBox2.Enabled = True

    ElseIf CheckBox10.Value = False Then
        TextBox1.Enabled = False
        TextBox2.Enabled = False

    End If
End Sub

Private Sub CheckBox11_Click()
    If CheckBox11.Value = True Then
        TextBox3.Enabled = True
        TextBox4.Enabled = True

    ElseIf CheckBox11.Value = False Then
        TextBox3.Enabled = False
        TextBox4.Enabled = False

    End If
End Sub

Private Sub CheckBox12_Click()
    If CheckBox12.Value = True Then
        TextBox5.Enabled = True
        TextBox6.Enabled = True

    ElseIf CheckBox12.Value = False Then
        TextBox5.Enabled = False
        TextBox6.Enabled = False

    End If
End Sub

Private Sub CheckBox13_Click()
    If CheckBox13.Value = True Then
        TextBox7.Enabled = True
        TextBox8.Enabled = True

    ElseIf CheckBox13.Value = False Then
        TextBox7.Enabled = False
        TextBox8.Enabled = False

    End If
End Sub

Private Sub CheckBox14_Click()
    If CheckBox14.Value = True Then
        TextBox9.Enabled = True
        TextBox10.Enabled = True

    ElseIf CheckBox14.Value = False Then
        TextBox9.Enabled = False
        TextBox10.Enabled = False

    End If
End Sub

Public Sub Tilbage_Click()
Me.Hide
'go back to previously stored form#
Call goBack
'SFunc.ShowFunc ("frm021")
'frm021.Show
End Sub

Private Sub txtFFSlut_Change()

End Sub

Private Sub txtFFStart_Change()

End Sub

Private Sub txtPSLstart_Change()

End Sub

Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeClip

Dim valgMd As Boolean
Dim valgAr As Boolean

valgMd = False
valgAr = False


' Indlæs tidligere svar 13
' Samme dag i måneden

If findPreviousAns(findTopSpm("F"), "13", 1) = "Nej/Ved ikke" Then
    CheckBox3.Value = True
End If
If findPreviousAns(findTopSpm("F"), "13.a", 0) <> "" Then
    CheckBox1.Value = True
End If
'Forfaldsdato
If findPreviousAns(findTopSpm("F"), "13.a_1", 0) <> "" Then
    Forfaldsdato.Value = True
    txtFFStart = findPreviousAns(findTopSpm("F"), "13.a_1", 1)
    txtFFSlut = findPreviousAns(findTopSpm("F"), "13.a_1", 2)
    If txtFFStart = "Sidste dag i måneden" Then
        CheckBox4.Value = True
        txtFFStart = ""
    End If
    If txtFFSlut = "Sidste dag i måneden" Then
        CheckBox4.Value = True
        txtFFSlut = ""
    End If
End If
'If Worksheets("SpmSvar").Range("D31:D31").Value <> "" Then
'    Forfaldsdato.Value = True
'    txtFFStart = Worksheets("SpmSvar").Range("E31:E31").Value
'    txtFFSlut = Worksheets("SpmSvar").Range("F31:F31").Value
'
'    If txtFFSlut = "Sidste dag i måneden" Then
'        CheckBox4.Value = True
'        txtFFSlut = ""
'    End If
'
'    valgMd = True
'End If

'SRB
If findPreviousAns(findTopSpm("F"), "13.a_2", 0) <> "" Then
    SRB.Value = True
    txtSRBstart = findPreviousAns(findTopSpm("F"), "13.a_2", 1)
    txtSRBslut = findPreviousAns(findTopSpm("F"), "13.a_2", 2)
    If txtSRBstart = "Sidste dag i måneden" Then
        CheckBox5.Value = True
        txtSRBstart = ""
    End If
    If txtSRBslut = "Sidste dag i måneden" Then
        CheckBox5.Value = True
        txtSRBslut = ""
    End If
End If
'If Worksheets("SpmSvar").Range("D32:D32").Value <> "" Then
'    SRB.Value = True
'    txtSRBstart = Worksheets("SpmSvar").Range("E32:E32").Value
'    txtSRBslut = Worksheets("SpmSvar").Range("F32:F32").Value
'
'    If txtSRBslut = "Sidste dag i måneden" Then
'        CheckBox5.Value = True
'        txtSRBslut = ""
'    End If
'
'    valgMd = True
'End If

'Stiftelsesdato
If findPreviousAns(findTopSpm("F"), "13.a_3", 0) <> "" Then
    Stiftelsesdato.Value = True
    txtSTIstart = findPreviousAns(findTopSpm("F"), "13.a_3", 1)
    txtSTIslut = findPreviousAns(findTopSpm("F"), "13.a_3", 2)
    If txtSTIstart = "Sidste dag i måneden" Then
        CheckBox6.Value = True
        txtSTIstart = ""
    End If
    If txtSTIslut = "Sidste dag i måneden" Then
        CheckBox6.Value = True
        txtSTIslut = ""
    End If
End If
'If Worksheets("SpmSvar").Range("D33:D33").Value <> "" Then
'    Stiftelsesdato.Value = True
'    txtSTIstart = Worksheets("SpmSvar").Range("E33:E33").Value
'    txtSTIslut = Worksheets("SpmSvar").Range("F33:F33").Value
'
'    If txtSTIslut = "Sidste dag i måneden" Then
'        CheckBox6.Value = True
'        txtSTIslut = ""
'    End If
'
'    valgMd = True
'End If

'PeriodeStartdato
If findPreviousAns(findTopSpm("F"), "13.a_4", 0) <> "" Then
    PeriodeStartdato.Value = True
    txtPSTstart = findPreviousAns(findTopSpm("F"), "13.a_4", 1)
    txtPSTslut = findPreviousAns(findTopSpm("F"), "13.a_4", 2)
    If txtPSTstart = "Sidste dag i måneden" Then
        CheckBox7.Value = True
        txtPSTstart = ""
    End If
    If txtPSTslut = "Sidste dag i måneden" Then
        CheckBox7.Value = True
        txtPSTslut = ""
    End If
End If
'If Worksheets("SpmSvar").Range("D34:D34").Value <> "" Then
'    PeriodeStartdato.Value = True
'    txtPSTstart = Worksheets("SpmSvar").Range("E34:E34").Value
'    txtPSTslut = Worksheets("SpmSvar").Range("F34:F34").Value
'
'    If txtPSTslut = "Sidste dag i måneden" Then
'        CheckBox7.Value = True
'        txtPSTslut = ""
'    End If
'
'    valgMd = True
'End If

If findPreviousAns(findTopSpm("F"), "13.a_5", 0) <> "" Then
    PeriodeSlutdato.Value = True
    txtPSLstart = findPreviousAns(findTopSpm("F"), "13.a_5", 1)
    txtPSLslut = findPreviousAns(findTopSpm("F"), "13.a_5", 2)
    If txtPSLstart = "Sidste dag i måneden" Then
        CheckBox8.Value = True
        txtPSLstart = ""
    End If
    If txtPSLslut = "Sidste dag i måneden" Then
        CheckBox8.Value = True
        txtPSLslut = ""
    End If
End If
'If Worksheets("SpmSvar").Range("D35:D35").Value <> "" Then
'    PeriodeSlutdato.Value = True
'    txtPSLstart = Worksheets("SpmSvar").Range("E35:E35").Value
'    txtPSLslut = Worksheets("SpmSvar").Range("F35:F35").Value
'
'    If txtPSLslut = "Sidste dag i måneden" Then
'        CheckBox8.Value = True
'        txtPSLslut = ""
'    End If
'
'    valgMd = True
'End If

'If valgMd = True Then
'    CheckBox1.Value = True
'End If

'Samme måned i året:
If findPreviousAns(findTopSpm("F"), "13.b", 0) <> "" Then
    CheckBox2.Value = True
End If
'Forfaldsdato
If findPreviousAns(findTopSpm("F"), "13.b_1", 0) <> "" Then
    CheckBox10.Value = True
    TextBox1 = findPreviousAns(findTopSpm("F"), "13.b_1", 1)
    TextBox2 = findPreviousAns(findTopSpm("F"), "13.b_1", 2)
End If
'If Worksheets("SpmSvar").Range("D36:D36").Value <> "" Then
'    CheckBox10.Value = True
'    TextBox1 = Worksheets("SpmSvar").Range("E36:E36").Value
'    TextBox2 = Worksheets("SpmSvar").Range("F36:F36").Value
'
'    valgAr = True
'End If

'SRB
If findPreviousAns(findTopSpm("F"), "13.b_2", 0) <> "" Then
    CheckBox11.Value = True
    TextBox3 = findPreviousAns(findTopSpm("F"), "13.b_2", 1)
    TextBox4 = findPreviousAns(findTopSpm("F"), "13.b_2", 2)
End If
'If Worksheets("SpmSvar").Range("D37:D37").Value <> "" Then
'    CheckBox11.Value = True
'    TextBox3 = Worksheets("SpmSvar").Range("E37:E37").Value
'    TextBox4 = Worksheets("SpmSvar").Range("F37:F37").Value
'
'    valgAr = True
'End If

'Stiftelsesdato
If findPreviousAns(findTopSpm("F"), "13.b_3", 0) <> "" Then
    CheckBox12.Value = True
    TextBox5 = findPreviousAns(findTopSpm("F"), "13.b_3", 1)
    TextBox6 = findPreviousAns(findTopSpm("F"), "13.b_3", 2)
End If
'If Worksheets("SpmSvar").Range("D38:D38").Value <> "" Then
'    CheckBox12.Value = True
'    TextBox5 = Worksheets("SpmSvar").Range("E38:E38").Value
'    TextBox6 = Worksheets("SpmSvar").Range("F38:F38").Value
'
'    valgAr = True
'End If

'PeriodeStartdato
If findPreviousAns(findTopSpm("F"), "13.b_4", 0) <> "" Then
    CheckBox13.Value = True
    TextBox7 = findPreviousAns(findTopSpm("F"), "13.b_4", 1)
    TextBox8 = findPreviousAns(findTopSpm("F"), "13.b_4", 2)
End If
'If Worksheets("SpmSvar").Range("D39:D39").Value <> "" Then
'    CheckBox13.Value = True
'    TextBox7 = Worksheets("SpmSvar").Range("E39:E39").Value
'    TextBox8 = Worksheets("SpmSvar").Range("F39:F39").Value
'
'    valgAr = True
'End If

If findPreviousAns(findTopSpm("F"), "13.b_5", 0) <> "" Then
    CheckBox14.Value = True
    TextBox9 = findPreviousAns(findTopSpm("F"), "13.b_5", 1)
    TextBox10 = findPreviousAns(findTopSpm("F"), "13.b_5", 2)
End If
'If Worksheets("SpmSvar").Range("D40:D40").Value <> "" Then
'    CheckBox14.Value = True
'    TextBox9 = Worksheets("SpmSvar").Range("E40:E40").Value
'    TextBox10 = Worksheets("SpmSvar").Range("F40:F40").Value
'
'    valgAr = True
'End If

'If valgAr = True Then
'    CheckBox2.Value = True
'End If

Call drawProgressBar(Me, Me.Name)
End Sub
