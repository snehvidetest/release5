Attribute VB_Name = "Func"
Function check_numeric(X, msg)
    If IsNumeric(X) = False Then
         MsgBox (msg)
    End If
End Function

Function Insert_to_sheet(sheet0, range0, value0)
    Worksheets(sheet0).Range(range0).Value = value0
    If value0 = "JA" Then
        Worksheets(sheet0).Range(range0).Interior.Color = RGB(198, 239, 206) 'Background color
        Worksheets(sheet0).Range(range0).Font.Color = RGB(0, 97, 0)          'Text color
    ElseIf value0 = "NEJ" Then
        Worksheets(sheet0).Range(range0).Interior.Color = RGB(255, 199, 206) 'Background color
        Worksheets(sheet0).Range(range0).Font.Color = RGB(156, 0, 6)         'Text color
    End If


End Function


Function onlyDigits(s As String) As String
    Dim retval As String
    Dim i As Integer

    retval = ""
                        
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            retval = retval + Mid(s, i, 1)
        End If
    Next
                       '
    onlyDigits = retval
End Function

Function maxValueMonths(X As Variant) As String
    maxValueMonths = ""
    If X > 12 Then
    maxValueMonths = "Indsæt en gyldig måned i året"
    End If
End Function

Function maxValueDays(X As Variant) As String
    maxValueDays = ""
    If X > 31 Then
    maxValueDays = "Indsæt en gyldig dag i måneden"
    End If
End Function
Function check_day_month(X As String, msg As String, check As String) As Boolean
    
    Dim a As Long
    check_day_month = False
    
    If X = "" Then
        a = 0
    End If
    
    If X <> "" Then
        If IsNumeric(X) = False Then
            check_day_month = True
            dFunc.msgError = msg + " (1 og 2)"
            SFunc.ShowFunc ("frmMsg")
            'MsgBox (msg + " (1 og 2)")
            GoTo Tilbage
        End If
      
        If check = "1" Then
            a = CLng(X)
            If (a <= 0 Or a > 31) Then
                check_day_month = True
                dFunc.msgError = msg + " (1)"
                SFunc.ShowFunc ("frmMsg")
                'MsgBox (msg + " (1)")
                GoTo Tilbage
            End If
        End If
    
    
        If check = "2" Then
            a = CLng(X)
            If (a <= 0 Or a > 12) Then
                check_day_month = True
                dFunc.msgError = msg + " (2)"
                SFunc.ShowFunc ("frmMsg")
                'MsgBox (msg + " (2)")
            End If
        End If
    End If
    
Tilbage:

End Function

Function check_month(a1 As String, msg) As Boolean

    Dim a As Long
    
    check_month = False
            
    a = CLng(a1)
       
    If (a <= 0 Or a > 12) Then
        check_month = True
        dFunc.msgError = msg
        SFunc.ShowFunc ("frmMsg")
        'MsgBox (msg)
        GoTo ending
    End If
            
ending:

End Function

Function test(X) As Integer
    
    test = X + 2
    
End Function

Sub test2()
a = onlyDigits("hej123he4j")
MsgBox (a)
End Sub

Function executeTest()

End Function

Function writeSpmSvar(spmNum As String, caption As String, ans1 As String, Optional ans2 As String, Optional column As Integer)
' This function writes the answer to the form in the last row unless the question has already been answered before.
' If the question has been answered before it will delete all subsequent answers after the specified question and the write the new answer on the same row.

    Dim myRange As String
    Dim colStr As String
    
    If column = 0 Then
        column = 1
    End If
    
    If column = 6 Then
        colStr = "F"
    Else
        colStr = "A"
    End If
    Call deleteHistory(findTopSpm(colStr), spmNum)
    
    topSpmRow = findTopSpm(colStr)
    
    myRange = "A" & CStr(topSpmRow) & ":A" & CStr(topSpmRow)
    Worksheets("SpmSvar").Cells(topSpmRow, column).Value = spmNum
    myRange = "B" & CStr(topSpmRow) & ":B" & CStr(topSpmRow)
    Worksheets("SpmSvar").Cells(topSpmRow, column + 1).Value = caption
    myRange = "C" & CStr(topSpmRow) & ":C" & CStr(topSpmRow)
    Worksheets("SpmSvar").Cells(topSpmRow, column + 2).Value = ans1
    myRange = "D" & CStr(topSpmRow) & ":D" & CStr(topSpmRow)
    Worksheets("SpmSvar").Cells(topSpmRow, column + 3).Value = ans2
    
    
End Function

Function findTopSpm(column As String, Optional mySheet As String) As Integer
' This function finds the total number of rows filled in rows from the top, in the specified column
    Dim maxSpm As Integer
    Dim i As Integer
    Dim myRange As String
    
    'Set default sheet
    If mySheet = "" Then mySheet = "SpmSvar"
    
    maxSpm = 500
    
    For i = 1 To maxSpm
        myRange = column & CStr(i) & ":" & column & CStr(i)
        If Worksheets(mySheet).Range(myRange).Value = "" Then
            GoTo ending
        End If
    Next i
ending:
findTopSpm = i
End Function

Function deleteHistory(topSpmRow As Integer, spmNum As String)
' This function finds the first occurance of the question and then deletes all questions that might follow.
' This is to ensure that even though the user goes back and changes previous answers, only the answers in the final run will be saved.
    
    Dim i As Integer
    Dim spmRange As String
    Dim myRange As String
    Dim maxRange As String
    
    maxRange = "E" & CStr(topSpmRow)
    
    For i = 1 To topSpmRow
        spmRange = "A" & CStr(i)
        myRange = spmRange & ":" & spmRange
        If Worksheets("SpmSvar").Range(myRange).Value = spmNum Then
            myRange = spmRange & ":" & maxRange
            Worksheets("SpmSvar").Range(myRange).Value = ""
            GoTo ending
        End If
    Next i
ending:
End Function

Function findPreviousAns(topSpmRow As Integer, spmNum As String, ansNum As Integer, Optional column As Integer) As String
' This function searches the columns with the previous answers to the Spørgeskema.
' If it can find the question asked, it will then output the answer.

    Dim i As Integer
    Dim spmRange As String
    Dim myRange As String
    Dim maxRange As String
    Dim var As Integer
    Dim ansColumn As Integer
     
    findPreviousAns = ""
     
    If column = 0 Then
        column = 6
    End If
    'maxRange = "J" & CStr(topSpmRow)
    maxRange = CStr(column) & "," & CStr(topSpmRow)
    
    For i = 1 To topSpmRow
        'spmRange = "F" & CStr(i)
        'spmRange = CStr(column) & "," & CStr(i)
        'myRange = spmRange & ":" & spmRange
        If Worksheets("SpmSvar").Cells(i, column).Value = spmNum Then
            ansColumn = column + ansNum - 1
            findPreviousAns = Worksheets("SpmSvar").Cells(i, ansColumn + 2).Value
            GoTo ending
        End If
    Next i
ending:
End Function

Function savePreviousAns()
' This function stores the previous answers to the "Spørgeskema" in seperate columns.
' These answers will then be used in the initiliaze process of opening the forms.
    Dim spmRange As String
    Dim myRange As String
    Dim deleteRange As String
    
    topSpmRow = findTopSpm("A")
    topSaveRow = findTopSpm("F")
    deleteRange = "F1:J" & CStr(topSaveRow)
    
    Worksheets("SpmSvar").Range(deleteRange).Value = ""
    
    spmRange = "A1"
    myRange = "A1:E" & CStr(topSpmRow)
    saveRange = "F1:J" & CStr(topSpmRow)
    Worksheets("SpmSvar").Range(saveRange).Value = Worksheets("SpmSvar").Range(myRange).Value
ending:
End Function

Function findRow(topSpmRow As Integer, spmNum As String, Optional column As Integer, Optional sheet As String) As Integer
' This function outputs the first row which includes the chosen string in the specified column

    Dim i As Integer
    Dim spmRange As String
    Dim myRange As String
    Dim maxRange As String
    
    If column = 0 Then
        column = 1
    End If
    'maxRange = "E" & CStr(topSpmRow)
    If sheet = "" Then sheet = "SpmSvar"
    
    For i = 1 To topSpmRow
        'spmRange = "A" & CStr(i)
        'myRange = spmRange & ":" & spmRange
        If Worksheets(sheet).Cells(i, column).Value = spmNum Then
            findRow = i
            'myRange = spmRange & ":" & maxRange
            'Worksheets("SpmSvar").Range(myRange).Value = ""
            GoTo ending
        End If
    Next i
ending:
End Function

Function goBack()
' This function goes to to the previously recorded form and deletes the current form from the log
    Dim topSpmRow As Integer
    Dim prevForm As String
    
    topSpmRow = findTopSpm("A", "Form_Log")
    prevForm = Worksheets("Form_Log").Cells(topSpmRow - 1, 1).Value
    Worksheets("Form_Log").Cells(topSpmRow - 1, 1).Value = ""
    
    ShowFunc (prevForm)
End Function

Function recHis(currentForm As String)
' This function records the name of the current form and saves it in a log to be used in the goBack function
    Dim topSpmRow As Integer
    
    topSpmRow = findTopSpm("A", "Form_Log")
    Worksheets("Form_Log").Cells(topSpmRow, 1).Value = currentForm
    
End Function

Function drawProgressBar(Form As UserForm, frm As String)
'This function draws the progress bar using a bar chart in the spreadsheet "ProgressBar"

'DrawChart
    Dim totalFrmNr As Integer
    Dim Fname As String
    Dim MyChart As Chart
    totalFrmNr = 37
    frmNr = WorksheetFunction.VLookup(frm, Sheets("ProgressBar").Range("A1:B41"), 2, False)
'SaveChart
    Sheets("ProgressBar").Cells(2, 3).Value = frmNr
    Set MyChart = Sheets("ProgressBar").ChartObjects(1).Chart
        Fname = ThisWorkbook.Path & "\pBar.gif"
        MyChart.Export Filename:=Fname, FilterName:="GIF"
    
'LoadChart
    With Form.pBar
        .Picture = LoadPicture(Fname)
        .PictureSizeMode = fmPictureSizeModeStretch
    End With

'Delete File
    On Error Resume Next
    Kill Fname
    On Error GoTo 0

    
End Function
