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
Dim longArray() As Variant
myarray1 = fillArray(1, "1")
myarray2 = fillArray(2, "1")
myarray3 = fillArray(3, "1")
myarray4 = fillArray(4, "1")

ReDim longArray(UBound(myarray1) * 4)
Dim i As Integer

For i = 0 To (UBound(myarray1) - 1)
    longArray(i * 4) = myarray1(i + 1)
    longArray(i * 4 + 1) = myarray2(i + 1)
    longArray(i * 4 + 2) = myarray3(i + 1)
    longArray(i * 4 + 3) = myarray4(i + 1)
Next i



End Function

Function writeSpmSvar(spmNum As String, caption As String, ans1 As String, Optional ans2 As String, Optional column As Integer)

'Skriver til
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

Function fillArray(spmNum As String, Optional column As Integer) As Variant
Dim returnVal(4) As String
Dim i As Integer

If column = 0 Then column = 1

row = findRow(findTopSpm("A"), spmNum)
For i = column To column + 3
    returnVal(i) = Chr(i + 64) & row
Next i
fillArray = returnVal
End Function

Function goBack()
    Dim topSpmRow As Integer
    Dim prevForm As String
    
    topSpmRow = findTopSpm("A", "Form_Log")
    prevForm = Worksheets("Form_Log").Cells(topSpmRow - 1, 1).Value
    Worksheets("Form_Log").Cells(topSpmRow - 1, 1).Value = ""
    
    ShowFunc (prevForm)
End Function

Function recHis(currentForm As String)
    Dim topSpmRow As Integer
    
    topSpmRow = findTopSpm("A", "Form_Log")
    Worksheets("Form_Log").Cells(topSpmRow, 1).Value = currentForm
    
End Function
