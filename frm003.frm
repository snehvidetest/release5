VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm003 
   Caption         =   "Populationsafgrænsning"
   ClientHeight    =   6936
   ClientLeft      =   60
   ClientTop       =   204
   ClientWidth     =   10980
   OleObjectBlob   =   "frm003.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub OKButton_Click()

If OptionButton1 = False And OptionButton2 = False And OptionButton3 = False Then
    dFunc.msgError = "Vælg venligst et svar"
    SFunc.ShowFunc ("frmMsg")
    'frmMsg.Show
    'MsgBox ("Vælg venligst et svar")
    GoTo ending
End If

If OptionButton1 Then
Call writeSpmSvar("4.a", Controls("Label1").caption, OptionButton1.caption)
    Me.Hide
    
    'store current form#
    recHis ("frm003")
    SFunc.ShowFunc ("frm004")
    'frm004.Show
ElseIf OptionButton2 Then
Call writeSpmSvar("4.a", Controls("Label1").caption, OptionButton2.caption)

    Me.Hide
    'temporarily store current form#
    recHis ("frm003")
    SFunc.ShowFunc ("frm026")
    'frm026.Show
Else
Call writeSpmSvar("4.a", Controls("Label1").caption, OptionButton3.caption)
    dFunc.msgError = "Populationen skal afgrænses på ny, hvis motorvejen skal kunne anvendes"
    SFunc.ShowFunc ("frmMsg")
    'frmMsg.Show
    'MsgBox "Populationen skal afgrænses på ny, hvis motorvejen skal kunne anvendes"
    Me.Hide
    'temporarily store current form#
    recHis ("frm003")
    SFunc.ShowFunc ("frm002")
    'frm002.Show
End If

'Worksheets("SpmSvar").Range("C6:C6").Value = Label1.Caption
If OptionButton1.Value = True Then
'    Worksheets("SpmSvar").Range("D6:D6").Value = OptionButton1.Caption

ElseIf OptionButton2.Value = True Then
'    Worksheets("SpmSvar").Range("D6:D6").Value = OptionButton2.Caption
    frm002.txtModtStart.Value = "01-09-2013"
'    Worksheets("SpmSvar").Range("D4:D4").Value = frm002.txtModtStart.Value
    Worksheets("Population").Range("B4:B4").Value = frm002.txtModtStart.Value
    frm002.txtModtSlut.Value = ""
'    Worksheets("SpmSvar").Range("E4:E4").Value = frm002.txtModtSlut.Value
    Worksheets("Population").Range("B5:B5").Value = frm002.txtModtSlut.Value
    
    
    
ElseIf OptionButton3.Value = True Then
'    Worksheets("SpmSvar").Range("D6:D6").Value = OptionButton3.Caption

End If



ending:
End Sub



Private Sub OptionButton1_Click()

End Sub

Private Sub OptionButton2_Click()

End Sub

Public Sub Tilbage_Click()

Me.Hide

'go back to previously stored form#
'go back to previously stored form#
Call goBack
'SFunc.ShowFunc ("frm002")
'frm002.Show

End Sub

Private Sub UserForm_Initialize()
Dim myAns As String

OptionButton1.Value = False
OptionButton2.Value = False
OptionButton3.Value = False
myAns = findPreviousAns(findTopSpm("F"), "4.a", 1)
'Fill JA/NEJ ComboBox
Image1.PictureSizeMode = fmPictureSizeModeClip
' Activate sheet
' Worksheets("Population").Activate

' Indlæs tidligere svar 4a
If myAns = OptionButton1.caption Then 'Worksheets("SpmSvar").Range("D6:D6").Value = OptionButton1.Caption Then
    OptionButton1.Value = True
ElseIf myAns = OptionButton2.caption Then
    OptionButton2.Value = True
ElseIf myAns = OptionButton3.caption Then
    OptionButton3.Value = True
End If

Call drawProgressBar(Me, Me.Name)

End Sub
