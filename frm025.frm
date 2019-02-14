VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm025 
   Caption         =   "Afslutning"
   ClientHeight    =   4836
   ClientLeft      =   84
   ClientTop       =   264
   ClientWidth     =   7212
   OleObjectBlob   =   "frm025.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm025"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub OKButton_Click()
    Me.Hide
    'store current form#
    recHis ("frm025")
    
    Call formatPDF
    
    Call SavePDF
    
    ' Close all
    'Dim UForm As Object
    'Dim i As Integer
    'i = 0
    'For Each UForm In VBA.UserForms
        'Debug.Print UForm.Name
    '    UForm.Hide
    '    Unload VBA.UserForms(i)
    '   i = i + 1
    'Next
    Dim objLoop As Object
    
    For Each objLoop In VBA.UserForms
        If TypeOf objLoop Is UserForm Then Unload objLoop
    Next objLoop
    
    'dFunc.msgError = "Tak - din besvarelse er nu gemt !"
    'SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Tak - din besvarelse er nu gemt !")
End Sub

Public Sub Tilbage_Click()
    Me.Hide
    'go back to previously stored form#
    Call goBack
    'SFunc.ShowFunc ("frm024")
    'frm024.Show
End Sub

Private Sub SavePDF()
    ' Save PDF
    Dim PathString
    PathString = Application.ActiveWorkbook.Path
    PathString = PathString & "\SpørgeskemaBesvarelse.pdf"
    
    Worksheets("PDF").Activate
    
    With ActiveSheet.PageSetup
        .Orientation = xlLandscape
    End With
    
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=PathString, _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
            :=True, OpenAfterPublish:=True

End Sub

Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeStretch

End Sub

Private Sub formatPDF()
    'Clear PDF sheet
    Sheets("PDF").Range("A1:E200").ClearContents
    Sheets("PDF").Range("A1:E200").ClearFormats
    
    'Copy from Answer to PDF
    Dim maxRow As Integer
    maxRow = findTopSpm("A", "SpmSvar") - 1
    Worksheets("PDF").Range("A1", "C" & maxRow).Value = Worksheets("SpmSvar").Range("A1", "C" & maxRow).Value
    Worksheets("PDF").Range("E1", "E" & maxRow).Value = Worksheets("SpmSvar").Range("D1", "D" & maxRow).Value
    
    'Format PDF
    Worksheets("PDF").Range("A1", "E" & maxRow).WrapText = True
    Dim rng As Range

    Set rng = Worksheets("PDF").Range("A1", "E" & maxRow)

    With rng.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
    
    With rng
        .HorizontalAlignment = xlLeft
    End With
    
    Set rng = Worksheets("PDF").Range("C1", "C" & maxRow)
    With rng
        .HorizontalAlignment = xlRight
    End With
    
    Set rng = Worksheets("PDF").Range("A1", "E1")
    With rng
        .Font.Bold = True
    End With
    
    Set rng = Worksheets("PDF").Range("D1", "D" & maxRow)
    
    Dim edges(1) As Variant
    edges(0) = xlEdgeLeft
    edges(1) = xlEdgeRight
    rng.Borders(edges(1)).LineStyle = xlNone
    rng.Borders(edges(0)).LineStyle = xlDot
End Sub
