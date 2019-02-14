VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm036 
   Caption         =   "Frasortering"
   ClientHeight    =   6936
   ClientLeft      =   60
   ClientTop       =   276
   ClientWidth     =   10980
   OleObjectBlob   =   "frm036.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm036"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Public Sub OKButton_Click()
      
    ' Validering for numeriske værdier
    
    Dim cControl As Control
        
    For Each cControl In Me.Controls
        
        control_type = UCase(Left(cControl.Name, 4))
            
        If control_type = "TEXT" Then
           If cControl.Text = "" Then
              cControl.SetFocus
              dFunc.msgError = "Felt skal udfyldes med tal."
              SFunc.ShowFunc ("frmMsg")
              GoTo ending
           End If
           If cControl.Text <> "" Then
              If IsNumeric(cControl.Text) = False Then
                 cControl.SetFocus
                 dFunc.msgError = "Felt skal udfyldes med tal."
                 SFunc.ShowFunc ("frmMsg")
                 GoTo ending
              End If
           End If
        End If
        
    Next cControl
    
     ' Validering for forkert anvendelse af før/efter
    
    If ComboBox2.Value = "efter" And ComboBox4.Value = "før" Then
        dFunc.msgError = "Forkert anvendelse af før/efter"
        SFunc.ShowFunc ("frmMsg")
        GoTo ending
    End If
    
    ' Validering for 'efter'
    
    If ComboBox2.Value = "efter" Then
        If Int(TextBox1.Value) > Int(TextBox2.Value) Then
            dFunc.msgError = "Værdien i 'Fra' skal være mindre end værdien i 'Til'."
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    Dim antal As Integer
    
    Dim x1 As Variant
    Dim x2 As Variant
    
    ' Reset values
    
    Call Insert_to_sheet("Regler", "J23:O23", "")
    
    'Relationen mellem "periode slut" og "periode start"
    
    x1 = TextBox1.Value
    x2 = TextBox2.Value
    
     ' 'Før' fra foranstilles med minus
    If ComboBox2.Value = "før" Then
        x1 = "-" + x1
    End If
    
    ' 'Før' fra foranstilles med minus
    If ComboBox4.Value = "før" Then
        x2 = "-" + x2
    End If
    
    ' Validering for 'før'
    
    If ComboBox2.Value = "før" Then
        If Int(x1) > Int(x2) Then
            dFunc.msgError = "Værdien i 'Fra' skal være mindre end værdien i 'Til'."
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    ' Indsæt værdier i regler
    Call Insert_to_sheet("Regler", "J23:J23", x1)
    Call Insert_to_sheet("Regler", "M23:M23", x2)
       
    ' Aktiver regler
    Call Insert_to_sheet("Regler", "G23:G23", "JA")
    
    ' Validering af 'Periode start' kan ligge samme dag som eller op til 732 dage efter 'Periode slut'.
    
    If ComboBox2.Value = "før" And ComboBox4.Value = "før" Then
        If (Int(TextBox1.Value) - Int(TextBox2.Value) > 732) Then
            dFunc.msgError = "Antal dage mellem 'Periode start' og 'Periode slut' kan maksimalt være 732 dage. "
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    If ComboBox2.Value = "før" And ComboBox4.Value = "efter" Then
        If (Int(TextBox2.Value) + Int(TextBox1.Value) > 732) Then
            dFunc.msgError = "Antal dage mellem 'Periode start' og 'Periode slut' kan maksimalt være 732 dage. "
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    If ComboBox2.Value = "efter" And ComboBox4.Value = "efter" Then
        If (Int(TextBox2.Value) - Int(TextBox1.Value) > 732) Then
            dFunc.msgError = "Antal dage mellem 'Periode start' og 'Periode slut' kan maksimalt være 732 dage."
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
        
    ' Skriv svar ned i 'SpmSvar'
    
    ' Relationen mellem "periode slut" og "periode start"
    a = "Periode slut"
    b = "Periode start"
    VisuTitle = a & " i forhold til " & b
    
'    Worksheets("SpmSvar").Range("C62:C62").Value = VisuTitle
'    Worksheets("SpmSvar").Range("D62:D62").Value = TextBox1.Value
'    Worksheets("SpmSvar").Range("E62:E62").Value = "dage"
'    Worksheets("SpmSvar").Range("F62:F62").Value = ComboBox2.Value
'    Worksheets("SpmSvar").Range("G62:G62").Value = TextBox2.Value
'    Worksheets("SpmSvar").Range("H62:H62").Value = "dage"
'    Worksheets("SpmSvar").Range("I62:I62").Value = ComboBox4.Value
    Call writeSpmSvar("11.a_4", Controls("Label47").caption, Controls("Label46").caption & " " & TextBox1.Value & " " & Controls("Label49").caption & " " & ComboBox2.Text, Controls("Label40").caption & " " & TextBox2.Value & " " & Controls("Label50").caption & " " & ComboBox4.Text)
    
    ' Hvis fordringshaver svarer, at "periodeslut" kan ligge før "periode start"
    ' skal der komme en advarsel om, at dette ikke er "normalt".
    
    
    Me.Hide
    'store current form#
    recHis ("frm036")
    
    If ComboBox2.Value = "før" Or ComboBox4.Value = "før" Then
        SFunc.ShowFunc ("frm046")
        GoTo ending
    End If
    
    If frm039.CheckBox4.Value = True Then
        SFunc.ShowFunc ("frm037")
    Else
        SFunc.ShowFunc ("frm038")
    End If
    
ending:
End Sub

Private Sub TextBox1_Change()
    
    Count = TextBox1.Value
    DMY = "dage"
    FE = ComboBox2.Value
    
    If Not IsNumeric(Count) Then
        Exit Sub
    End If
       
    X = Count
      
    Worksheets("Grafik_frm036").Range("C2") = X
    
    Worksheets("Grafik_frm036").Range("B2") = Count & " " & DMY & " " & FE
    
    If FE = "efter" Then
    
        Worksheets("Grafik_frm036").Range("C2") = X
        
    ElseIf FE = "før" Then
    
        Worksheets("Grafik_frm036").Range("C2") = -X
        
    End If
    
    Call DrawChart
    
    
End Sub

Private Sub TextBox2_Change()
    
    Count = TextBox2.Value
    DMY = "dage"
    FE = ComboBox4.Value
    
    If Not IsNumeric(Count) Then
        Exit Sub
    End If
    
    X = Count
    
    Worksheets("Grafik_frm036").Range("C4") = X
    
    Worksheets("Grafik_frm036").Range("B4") = Count & " " & DMY & " " & FE
    
    If FE = "efter" Then
    
        Worksheets("Grafik_frm036").Range("C4") = X
        
    ElseIf FE = "før" Then
    
        Worksheets("Grafik_frm036").Range("C4") = -X
        
    End If
    
    Call DrawChart

End Sub


Private Sub DrawChart()

    Dim Fname As String

    Call SaveChart
    Fname = ThisWorkbook.Path & "\temp1.gif"
    With Me.Image2
        .Picture = LoadPicture(Fname)
        .PictureSizeMode = fmPictureSizeModeClip
    End With
    Call DeleteFile
    
End Sub

Private Sub SaveChart()

    Dim MyChart As Chart
    Dim Fname As String

    Set MyChart = Sheets("Grafik_frm036").ChartObjects(1).Chart
    Fname = ThisWorkbook.Path & "\temp1.gif"
    MyChart.Export Filename:=Fname, FilterName:="GIF"
    
End Sub

Sub DeleteFile()

    Dim Fname As String
    On Error Resume Next
    Fname = ThisWorkbook.Path & "\temp1.gif"
    Kill Fname
    On Error GoTo 0
    
End Sub

Public Sub Tilbage_Click()
    Me.Hide
    'go back to previously stored form#
    Call goBack
    'SFunc.ShowFunc ("frm035")
End Sub

Private Sub ComboBox2_Change()
    Call TextBox1_Change
End Sub

Private Sub ComboBox4_Change()
    Call TextBox2_Change
End Sub
Private Sub UserForm_Initialize()

    Image1.PictureSizeMode = fmPictureSizeModeClip

    ' Activate sheet
'    Worksheets("SpmSvar").Activate
'    ActiveWindow.Zoom = 80
'    Worksheets("Grafik_frm036").Range("A1").Select
    
    a = "Periode slut"
    b = "Periode start"
    VisuTitle = a & " i forhold til " & b
    
    Worksheets("Grafik_frm036").Range("B3") = b
    Worksheets("Grafik_frm036").Range("A1") = VisuTitle
    Worksheets("Grafik_frm036").Range("B2") = "20 dage efter"
    Worksheets("Grafik_frm036").Range("C2") = 20
    Worksheets("Grafik_frm036").Range("B4") = "30 dage efter"
    Worksheets("Grafik_frm036").Range("C4") = 30
    
    Call DrawChart
    With ComboBox2
        .AddItem "før"
        .AddItem "efter"
    End With
    
    With ComboBox4
        .AddItem "før"
        .AddItem "efter"
    End With

    ' Activate sheet
    Worksheets("SpmSvar").Activate
    ActiveWindow.Zoom = 80
    Worksheets("SpmSvar").Range("I1").Select
    
    
    ' Indlæs tidligere svar fra 'SpmSvar'

    ' Relationen mellem forfaldsdato" og "sidste rettidige betalingsdato"
    If Not IsEmpty(findPreviousAns(findTopSpm("F"), "11.a_1", 1)) Then
        myLetter = "a"
    ElseIf Not IsEmpty(findPreviousAns(findTopSpm("F"), "11.b_1", 1)) Then
        myLetter = "b"
    Else
        Exit Sub
    End If
    If findPreviousAns(findTopSpm("F"), "11." & myLetter & "_4", 1) <> "" Then
        TextBox1.Value = Split(findPreviousAns(findTopSpm("F"), "11." & myLetter & "_4", 1))(1)
        ComboBox2.Value = Split(findPreviousAns(findTopSpm("F"), "11." & myLetter & "_4", 1))(3)
    End If
    
    If findPreviousAns(findTopSpm("F"), "11." & myLetter & "_4", 2) <> "" Then
        TextBox2.Value = Split(findPreviousAns(findTopSpm("F"), "11." & myLetter & "_4", 2))(1)
        ComboBox4.Value = Split(findPreviousAns(findTopSpm("F"), "11." & myLetter & "_4", 2))(3)
    End If
Call drawProgressBar(Me, Me.Name)
End Sub

