VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm001 
   Caption         =   "Population"
   ClientHeight    =   6936
   ClientLeft      =   36
   ClientTop       =   180
   ClientWidth     =   10980
   OleObjectBlob   =   "frm001.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub CommandButton1_Click()

    dFunc.msgYesNoTxt = "Er du sikker? Dette vil slette den tidligere besvarelse, hvis en sådan eksisterer."
    SFunc.ShowFunc ("frmMsgYesNo")
        
    If dFunc.msgYesNo = "NEJ" Then
       'bliv på siden
    Else
       'start forfra
        Worksheets("SpmSvar").Range("A2:I150").Value = ""
        Worksheets("Form_Log").Range("A2:A500").Value = ""
        frm002.lblFtypeTxt.caption = ""
        frm002.lblFhaverTxt.caption = ""
        frm002.UserForm_Initialize
        Me.Hide
        dFunc.msgYesNoTxt = ""
        recHis ("frm001")
        SFunc.ShowFunc ("frm002")
        'frm002.Show
    End If
    'Call YesNoMessageBox
    
End Sub

Private Sub CommandButton2_Click()

End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    
End Sub
Public Sub OKButton_Click()
    Call savePreviousAns
    Worksheets("Form_Log").Range("A2:A500").Value = ""
    Me.Hide
    recHis ("frm001")
    ShowFunc ("frm002")
    'frm002.Show
End Sub

Sub YesNoMessageBox()
 
Dim Answer As String
Dim MyNote As String
 
    'Place your text here
    MyNote = "Er du sikker? Dette vil slette den tidligere besvarelse, hvis en sådan eksisterer."
 
    'Display MessageBox
    Answer = MsgBox(MyNote, vbQuestion + vbOKCancel, "Ny Besvarelse")
 
    If Answer = vbOK Then
        Worksheets("SpmSvar").Range("D2:I150").Value = ""
        frm002.UserForm_Initialize
        Me.Hide
        recHis ("frm001")
        frm002.Show
    End If
 
End Sub

Private Sub Udvikler_Click()
    UdviklerAdgang.Show
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

    Image1.PictureSizeMode = fmPictureSizeModeClip
    Worksheets("SpmSvar").Activate

End Sub
