VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm037 
   Caption         =   "Frasortering"
   ClientHeight    =   6276
   ClientLeft      =   84
   ClientTop       =   372
   ClientWidth     =   9060.001
   OleObjectBlob   =   "frm037.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm037"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Public Sub OKButton_Click()
       
    ' Validering for numeriske v�rdier
    
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
    
    ' Validering for forkert anvendelse af f�r/efter
    
    If ComboBox2.Value = "efter" And ComboBox4.Value = "f�r" Then
        dFunc.msgError = "Forkert anvendelse af f�r/efter"
        SFunc.ShowFunc ("frmMsg")
        'MsgBox ("Forkert anvendelse af f�r/efter")
        GoTo ending
    End If
       
    ' Validering for 'efter'
    
    If ComboBox2.Value = "efter" Then
        If Int(TextBox1.Value) > Int(TextBox2.Value) Then
            dFunc.msgError = "V�rdien i 'Fra' skal v�re mindre end v�rdien i 'Til'."
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    Dim antal As Integer
    
    Dim x1 As Variant
    Dim x2 As Variant
    
    ' Reset values
    
    Call Insert_to_sheet("Regler", "J21:O21", "")
    
    'Relationen mellem "stiftelsesdato" og "periode start"
    
    x1 = TextBox1.Value
    x2 = TextBox2.Value
    
    ' 'F�r' fra foranstilles med minus
    If ComboBox2.Value = "f�r" Then
        x1 = "-" + x1
    End If
    
    ' 'F�r' fra foranstilles med minus
    If ComboBox4.Value = "f�r" Then
        x2 = "-" + x2
    End If
    
    ' Validering for 'f�r'
    
    If ComboBox2.Value = "f�r" Then
        If Int(x1) > Int(x2) Then
            dFunc.msgError = "V�rdien i 'Fra' skal v�re mindre end v�rdien i 'Til'."
            SFunc.ShowFunc ("frmMsg")
            'MsgBox ("V�rdien i 'Fra' skal v�re mindre end v�rdien i 'Til'.")
            GoTo ending
        End If
    End If
    
    ' Validering af 'Stiftelsesdato' kan ligge samme dag som eller op til 365 dage efter 'Periode start'.
    
    If ComboBox2.Value = "f�r" And ComboBox4.Value = "f�r" Then
        If (Int(TextBox1.Value) - Int(TextBox2.Value) > 365) Then
            dFunc.msgError = "Antal dage mellem 'Stiftelsesdato' og 'Periode start' kan maksimalt v�re 365 dage. "
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    If ComboBox2.Value = "f�r" And ComboBox4.Value = "efter" Then
        If (Int(TextBox2.Value) + Int(TextBox1.Value) > 365) Then
            dFunc.msgError = "Antal dage mellem 'Stiftelsesdato' og 'Periode start' kan maksimalt v�re 365 dage. "
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    If ComboBox2.Value = "efter" And ComboBox4.Value = "efter" Then
        If (Int(TextBox2.Value) - Int(TextBox1.Value) > 365) Then
            dFunc.msgError = "Antal dage mellem 'Stiftelsesdato' og 'Periode start' kan maksimalt v�re 365 dage. "
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    ' Inds�t v�rdier i regler
    Call Insert_to_sheet("Regler", "J21:J21", x1)
    Call Insert_to_sheet("Regler", "M21:M21", x2)
   
    ' Aktiver regler
    Call Insert_to_sheet("Regler", "G21:G21", "JA")
    
    ' Skriv svar ned i 'SpmSvar'
    
    ' Relationen mellem "stiftelsesdato" og "periode start"
    a = "Stiftelsesdato"
    b = "Periode start"
    VisuTitle = a & " i forhold til " & b
'    Worksheets("SpmSvar").Range("C63:C63").Value = VisuTitle
'    Worksheets("SpmSvar").Range("D63:D63").Value = TextBox1.Value
'    Worksheets("SpmSvar").Range("E63:E63").Value = "dage"
'    Worksheets("SpmSvar").Range("F63:F63").Value = ComboBox2.Value
'    Worksheets("SpmSvar").Range("G63:G63").Value = TextBox2.Value
'    Worksheets("SpmSvar").Range("H63:H63").Value = "dage"
'    Worksheets("SpmSvar").Range("I63:I63").Value = ComboBox4.Value
    Call writeSpmSvar("11.a_1", Controls("Label47").caption, Controls("Label46").caption & " " & TextBox1.Value & " " & Controls("Label49").caption & " " & ComboBox2.Text, Controls("Label40").caption & " " & TextBox2.Value & " " & Controls("Label50").caption & " " & ComboBox4.Text)
    
    
    ' Hvis fordringshaver svarer, at "stiftelsesdato" kan ligge f�r "periode start"
    ' skal der komme en advarsel om, at dette ikke er "normalt".
        
    If ComboBox2.Value = "f�r" Or ComboBox4.Value = "f�r" Then
        SFunc.ShowFunc ("frm047")
        GoTo ending
    End If
     
    Me.Hide
    'store current form#
    recHis ("frm037")
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
    
    Worksheets("Grafik_frm037").Range("C2") = X
    
    Worksheets("Grafik_frm037").Range("B2") = Count & " " & DMY & " " & FE
    
    If FE = "efter" Then
    
        Worksheets("Grafik_frm037").Range("C2") = X
        
    ElseIf FE = "f�r" Then
    
        Worksheets("Grafik_frm037").Range("C2") = -X
        
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
        
    Worksheets("Grafik_frm037").Range("C4") = X
    
    Worksheets("Grafik_frm037").Range("B4") = Count & " " & DMY & " " & FE
    
    If FE = "efter" Then
    
        Worksheets("Grafik_frm037").Range("C4") = X
        
    ElseIf FE = "f�r" Then
    
        Worksheets("Grafik_frm037").Range("C4") = -X
        
    End If
    
    Call DrawChart

End Sub

Private Sub DrawChart()

    Dim Fname As String

    Call SaveChart
    Fname = ThisWorkbook.Path & "\temp1.gif"
    With Me.Image2
        .Picture = LoadPicture(Fname)
        .PictureSizeMode = fmPictureSizeModeZoom
    End With
    Call DeleteFile
    
End Sub

Private Sub SaveChart()

    Dim MyChart As Chart
    Dim Fname As String

    Set MyChart = Sheets("Grafik_frm037").ChartObjects(1).Chart
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

Private Sub ComboBox2_Change()
    Call TextBox1_Change
End Sub

Private Sub ComboBox4_Change()
    Call TextBox2_Change
End Sub
Private Sub UserForm_Initialize()

    Image1.PictureSizeMode = fmPictureSizeModeStretch

'    ' Activate sheet
'    Worksheets("SpmSvar").Activate
'    ActiveWindow.Zoom = 80
'    Worksheets("SpmSvar").Range("I1").Select
    
    a = "Stiftelsesdato"
    b = "Periode start"
    VisuTitle = a & " i forhold til " & b
    
    Worksheets("Grafik_frm037").Range("B3") = b
    Worksheets("Grafik_frm037").Range("A1") = VisuTitle
    Worksheets("Grafik_frm037").Range("B2") = "20 dage efter"
    Worksheets("Grafik_frm037").Range("C2") = 20
    Worksheets("Grafik_frm037").Range("B4") = "30 dage efter"
    Worksheets("Grafik_frm037").Range("C4") = 30
    
    Call DrawChart
    
    With ComboBox2
        .AddItem "f�r"
        .AddItem "efter"
    End With
    
    With ComboBox4
        .AddItem "f�r"
        .AddItem "efter"
    End With

    
    ' Activate sheet
    Worksheets("SpmSvar").Activate
    ActiveWindow.Zoom = 80
    Worksheets("SpmSvar").Range("A1").Select
    
    ' Indl�s tidligere svar fra 'SpmSvar'

    ' Relationen mellem forfaldsdato" og "sidste rettidige betalingsdato"
    If findPreviousAns(findTopSpm("F"), "11.a_1", 1) <> "" Then
        TextBox1.Value = Split(findPreviousAns(findTopSpm("F"), "11.a_1", 1))(1)
        ComboBox2.Value = Split(findPreviousAns(findTopSpm("F"), "11.a_1", 1))(3)
    End If
    If findPreviousAns(findTopSpm("F"), "11.a_1", 2) <> "" Then
        TextBox2.Value = Split(findPreviousAns(findTopSpm("F"), "11.a_1", 2))(1)
        ComboBox4.Value = Split(findPreviousAns(findTopSpm("F"), "11.a_1", 2))(3)
    End If
End Sub

