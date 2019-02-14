VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm038 
   Caption         =   "Frasortering"
   ClientHeight    =   6936
   ClientLeft      =   60
   ClientTop       =   168
   ClientWidth     =   10980
   OleObjectBlob   =   "frm038.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm038"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ComboBox2_Change()
    Call TextBox1_Change
End Sub


Private Sub ComboBox4_Change()
    Call TextBox2_Change
End Sub



Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Public Sub OKButton_Click()
           
    ' Validering for forkert anvendelse af før/efter
    
    If ComboBox2.Value = "efter" And ComboBox4.Value = "før" Then
        dFunc.msgError = "Forkert anvendelse af før/efter"
        SFunc.ShowFunc ("frmMsg")
        GoTo ending
    End If
    
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
    
    Call Insert_to_sheet("Regler", "J22:O22", "")
    
    'Relationen mellem "stiftelsesdato" og "periode slut"
    
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
            'MsgBox ("Værdien i 'Fra' skal være mindre end værdien i 'Til'.")
            GoTo ending
        End If
    End If
    
    ' Validering af 'Stiftelsesdato' kan ligge fra 10 dage før til 1081 dage efter 'Periode slut'.
    
    If ComboBox2.Value = "før" Then
        If (Int(TextBox1.Value) > 10) Then
            dFunc.msgError = "'Stiftelsesdato' kan minimalt ligge 10 dage før 'Periode slut'."
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    If ComboBox2.Value = "før" And ComboBox4.Value = "efter" Then
        If (Int(TextBox2.Value) > 1081) Then
            dFunc.msgError = "'Stiftelsesdato' kan maksimalt ligge 1081 dage efter 'Periode slut'."
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    If ComboBox2.Value = "efter" And ComboBox4.Value = "efter" Then
        If (Int(TextBox2.Value) - Int(TextBox1.Value) > 1081) Then
            dFunc.msgError = "'Stiftelsesdato' kan maksimalt ligge 1081 dage efter 'Periode slut'. "
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    ' Indsæt værdier i  regel
    Call Insert_to_sheet("Regler", "J22:J22", x1)
    Call Insert_to_sheet("Regler", "M22:M22", x2)
    
    ' Aktiver regler
    Call Insert_to_sheet("Regler", "G22:G22", "JA")
    
    ' Skriv svar ned i 'SpmSvar'
    
    ' Relationen mellem "stiftelsesdato" og "periode slut"
    a = "Stiftelsesdato"
    b = "Periode slut"
    VisuTitle = a & " i forhold til " & b
'    Worksheets("SpmSvar").Range("C64:C64").Value = VisuTitle
'    Worksheets("SpmSvar").Range("D64:D64").Value = TextBox1.Value
'    Worksheets("SpmSvar").Range("E64:E64").Value = "dage"
'    Worksheets("SpmSvar").Range("F64:F64").Value = ComboBox2.Value
'    Worksheets("SpmSvar").Range("G64:G64").Value = TextBox2.Value
'    Worksheets("SpmSvar").Range("H64:H64").Value = "dage"
'    Worksheets("SpmSvar").Range("I64:I64").Value = ComboBox4.Value
    Call writeSpmSvar("11.b_1", Label47.caption, Label46.caption & " " & TextBox1.Value & " " & Label50.caption & " " & ComboBox2.Text, Label40.caption & " " & TextBox2.Value & " " & Label51.caption & " " & ComboBox4.Text)
    
    Me.Hide
    'store current form#
    recHis ("frm038")
    SFunc.ShowFunc ("frm021")
       
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
        
    Worksheets("Grafik_frm038").Range("C2") = X
    
    Worksheets("Grafik_frm038").Range("B2") = Count & " " & DMY & " " & FE
    
    If FE = "efter" Then
    
        Worksheets("Grafik_frm038").Range("C2") = X
        
    ElseIf FE = "før" Then
    
        Worksheets("Grafik_frm038").Range("C2") = -X
        
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
    
    Worksheets("Grafik_frm038").Range("C4") = X
    
    Worksheets("Grafik_frm038").Range("B4") = Count & " " & DMY & " " & FE
    
    If FE = "efter" Then
    
        Worksheets("Grafik_frm038").Range("C4") = X
        
    ElseIf FE = "før" Then
    
        Worksheets("Grafik_frm038").Range("C4") = -X
        
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

    Set MyChart = Sheets("Grafik_frm038").ChartObjects(1).Chart
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
    'SFunc.ShowFunc ("frm036")
End Sub

Private Sub UserForm_Initialize()

    Image1.PictureSizeMode = fmPictureSizeModeClip
    ' Activate sheet
    Worksheets("SpmSvar").Activate
    ActiveWindow.Zoom = 80
    Worksheets("SpmSvar").Range("A1").Select

    a = "Stiftelsesdato"
    b = "Periode slut"
    VisuTitle = a & " i forhold til " & b
    
    Worksheets("Grafik_frm038").Range("B3") = b
    Worksheets("Grafik_frm038").Range("A1") = VisuTitle
    Worksheets("Grafik_frm038").Range("B2") = "20 dage efter"
    Worksheets("Grafik_frm038").Range("C2") = 20
    Worksheets("Grafik_frm038").Range("B4") = "30 dage efter"
    Worksheets("Grafik_frm038").Range("C4") = 30
    
    Call DrawChart
    
    With ComboBox2
        .AddItem "før"
        .AddItem "efter"
    End With
    
    With ComboBox4
        .AddItem "før"
        .AddItem "efter"
    End With

    
    ' Indlæs tidligere svar fra 'SpmSvar'

    ' Relationen mellem forfaldsdato" og "sidste rettidige betalingsdato"
    If findPreviousAns(findTopSpm("F"), "11.b_1", 1) <> "" Then
        TextBox1.Value = Split(findPreviousAns(findTopSpm("F"), "11.b_1", 1))(1)
        ComboBox2.Value = Split(findPreviousAns(findTopSpm("F"), "11.b_1", 1))(3)
    End If
    
    If findPreviousAns(findTopSpm("F"), "11.b_1", 2) <> "" Then
        TextBox2.Value = Split(findPreviousAns(findTopSpm("F"), "11.b_1", 2))(1)
        ComboBox4.Value = Split(findPreviousAns(findTopSpm("F"), "11.b_1", 2))(3)
    End If
Call drawProgressBar(Me, Me.Name)
End Sub

