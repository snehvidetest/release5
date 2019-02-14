VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm039 
   Caption         =   "Frasortering"
   ClientHeight    =   6936
   ClientLeft      =   36
   ClientTop       =   192
   ClientWidth     =   10980
   OleObjectBlob   =   "frm039.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm039"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CheckBox4_Click()

If CheckBox4.Value = True Then
    CheckBox5.Value = False
    CheckBox5.Enabled = False
Else
    CheckBox5.Value = True
    CheckBox5.Enabled = True
End If

End Sub

Private Sub CheckBox5_Click()

If CheckBox5.Value = True Then
    CheckBox4.Value = False
    CheckBox4.Enabled = False
Else
    CheckBox4.Value = True
    CheckBox4.Enabled = True
End If

End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label1_Click()

End Sub

Public Sub OKButton_Click()

    If CheckBox4.Value = False And CheckBox5.Value = False Then
        dFunc.msgError = "Vælg venligst en relation for 'stiftelsesdato'."
        SFunc.ShowFunc ("frmMsg")
        'MsgBox "Vælg venligst en relation for 'stiftelsesdato'.", vbExclamation, "Relation mangler"
        GoTo ending
    End If
    
    'Worksheets("SpmSvar").Range("D63:D63").Value = CheckBox4.Value
    'Worksheets("SpmSvar").Range("D64:D64").Value = CheckBox5.Value
    If CheckBox4.Value = True Then
        Call writeSpmSvar("11", Controls("Label1").caption, Controls("CheckBox4").caption)
    ElseIf CheckBox5.Value = True Then
        Call writeSpmSvar("11", Controls("Label1").caption, Controls("CheckBox5").caption)
    End If
    Me.Hide
    'store current form#
    recHis ("frm039")
    SFunc.ShowFunc ("frm034")
    'frm034.Show
    
ending:
End Sub

Public Sub Tilbage_Click()

    Me.Hide
    'go back to previously stored form#
    Call goBack
    
'    If frm007.OptionButton1 = True Then
'       If frm008.OptionButton1 = True Then
'            SFunc.ShowFunc ("frm008")
'        Else
'            SFunc.ShowFunc ("frm009")
'        End If
'    Else
'        If frm014.PeriodeSlutdato.Value = True Then
'                SFunc.ShowFunc ("frm031")
'        ElseIf frm014.PeriodeStartdato.Value = True Then
'                SFunc.ShowFunc ("frm030")
'        ElseIf frm014.Stiftelsesdato.Value = True Then
'                SFunc.ShowFunc ("frm029")
'        ElseIf frm014.SRB.Value = True Then
'                SFunc.ShowFunc ("frm032")
'        ElseIf frm014.Forfaldsdato.Value = True Then
'                SFunc.ShowFunc ("frm028")
'
'        Else
'                SFunc.ShowFunc ("frm014")
'        End If
'    End If
    
    'frm014.Show

End Sub
Private Sub UserForm_Initialize()

    Image1.PictureSizeMode = fmPictureSizeModeClip
    
    ' Activate sheet
    Worksheets("SpmSvar").Activate
    ActiveWindow.Zoom = 80
    Worksheets("SpmSvar").Range("I1").Select
    
    CheckBox1.Visible = False
    CheckBox2.Visible = False
    CheckBox3.Visible = False

    If findPreviousAns(findTopSpm("F"), "11", 1) = CheckBox4.caption Then
        CheckBox4.Value = True
    End If
    
    If findPreviousAns(findTopSpm("F"), "11", 1) = CheckBox5.caption Then
        CheckBox5.Value = True
    End If
Call drawProgressBar(Me, Me.Name)
End Sub
