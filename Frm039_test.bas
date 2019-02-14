Attribute VB_Name = "Frm039_test"
Private result As String
Private formID As Integer
Private formName As String
Private parameters As Scripting.Dictionary
Private parametersAndCols As Scripting.Dictionary
'Private spmCells As Scripting.Dictionary
Private popCells As Scripting.Dictionary
Private rulCells As Scripting.Dictionary
Private groCells As Scripting.Dictionary
Sub RunTests()
    'On Error GoTo Error_handler
    formName = "frm039"
    formID = 39
    
    Set parametersAndCols = Global_Test_Func.getParamtersAndTheirCols(formID)
    Set spmCells = New Scripting.Dictionary
    Set popCells = New Scripting.Dictionary
    Set groCells = New Scripting.Dictionary
    Set rulCells = New Scripting.Dictionary
    
    Dim nrTC As Integer, i As Integer
    nrTC = Application.WorksheetFunction.CountIf(testWS.Range("A:A"), formID)
    
    For i = 1 To nrTC
        Set parameters = New Scripting.Dictionary
        Testcase i
    Next i
    
    Exit Sub
Error_handler:
    Global_Test_Func.PrintTestResults CStr(formID) + "." + CStr(i), "crash", "False"
    Resume Next
End Sub


Private Function Testcase(tc As Integer)
    Dim review As Boolean, tcid As String
    
    'Reset spørgeskema workbook
    Global_Test_Func.resetSheets ThisWorkbook
    
    'Create the TCID
    tcid = Global_Test_Func.GetTCID(tc, formID)
    If logging Then
        Write #1, tcid
    End If
    
    Set parameters = New Scripting.Dictionary 'Resets the testcase parameters
    Set parameters = Global_Test_Func.getData(tcid, parametersAndCols)
    
    'Clear all fields related to spørskema
    'ClearAllFields ThisWorkbook

    ThisWorkbook.Activate
    
    If parameters("run") = 0 Then
         Exit Function
    End If
    
    Select Case parameters("testSubject")
    
        Case "printsToGroSheet"
            SetFields
            frm039.OKButton_Click 'Click on Videre button
            CheckFields "Gruppering"
            
        Case "printsToPopSheet"
            SetFields
            frm039.OKButton_Click 'Click on Videre button
            CheckFields "Population"
            
        Case "printsToSpmSheet"
            SetFields
            frm039.OKButton_Click 'Click on Videre button
            'CheckFields "SpmSvar"
            result = findPreviousAns(findTopSpm("A"), "11", 1, 1)
            
        Case "errorMessage"
            SetFields
            frm039.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.errorMessage
            
        Case "nextStep"
            SetFields
            frm039.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "backButton"
            'ThisWorkbook.Sheets("SpmSvar").Range("D24:H24").Value = "" 'Prevents crashing when frm010 initialises frm014
            recHis ("frm008")
            frm039.Tilbage_Click
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "tidligereBesvarelse"
            DataIsSaved "SpmSvar"
            
        Case "noExtraPrints"
            Sheet1.recordChangingCells = True
            SetFields
            'ThisWorkbook.Sheets("SpmSvar").Range("D24:H24").Value = "" 'Prevents crashing when frm010 initialises frm014
            If parameters("testParameter") = "noChangeWhenBackButton" Then
                frm039.Tilbage_Click 'Click back button
            Else
                frm039.OKButton_Click 'Click on Videre button
            End If
            CheckNoExtraPrints
            Sheet1.recordChangingCells = False
         
        Case Else
            MsgBox "Error in 'testsubject' input: tcid " & tcid 'Msgbox to stop the code because you made a mistake in the inputs..
    End Select
    
    If result = parameters("expected") Then
        review = True
    Else:
        review = False
    End If

    KillForms
    
     'Print results
    Global_Test_Func.PrintTestResults tcid, result, review

End Function


Private Function SetFields()
    
    'The folowing code inserts the inputs into the actual form
    frm039.CheckBox4.Value = parameters("checkbox4")
    frm039.CheckBox5.Value = parameters("checkbox5")

End Function

Private Function CheckFields(sheet As String)

    'Check results
    If (sheet = "SpmSvar") Then
        result = ThisWorkbook.Sheets(sheet).Range("D63").Text
    End If
    
End Function


Function DataIsSaved(sheet As String)
    
    'ThisWorkbook.Sheets(sheet).Range("D63").Value = parameters("checkbox4")
    'ThisWorkbook.Sheets(sheet).Range("D64").Value = parameters("checkbox5")
    
    Select Case parameters("testParameter")
        Case "checkbox4" '
            If parameters("checkbox4") = "True" Then
                Call writeSpmSvar("11", "", "Stiftelsesdato og Periodestart", "", 6)
            End If
        Case "checkbox5"
            If parameters("checkbox5") = "True" Then
                Call writeSpmSvar("11", "", "Stiftelsesdato og Periodeslut", "", 6)
            End If
    End Select
            
    ShowFunc (formName)
    
    Select Case parameters("testParameter")
        Case "checkbox4"
            result = CStr(frm039.CheckBox4.Value)
        Case "checkbox5"
            result = CStr(frm039.CheckBox5.Value)
    End Select
    
End Function
Private Function CheckNoExtraPrints()
    'Check Which configuration to choose
    Select Case parameters("testParameter")
        'Test different cases were different cells should be changed
        Case "noChangeWhenError"
        Case "noChangeWhenBackButton"
        Case "config1"
        Case "config2"
    End Select
    
    'returns a string which shows either true or has the input of the cells that changed that shouldn't have been changed
    result = Global_Test_Func.CheckPrintsInAllSheets(spmCells, popCells, rulCells, groCells)
    
    'Cleans up all arrays and dictionaries
    Sheet9.spmChangedCells.RemoveAll
    Sheet5.groChangedCells.RemoveAll
    Sheet3.rulChangedCells.RemoveAll
    Sheet1.popChangedCells.RemoveAll
    popCells.RemoveAll
    rulCells.RemoveAll
    groCells.RemoveAll
    spmCells.RemoveAll
End Function
Function KillForms()
    'Closes forms
    ThisWorkbook.Activate
    If IsLoaded("frm014") Then
        Unload frm014
    End If
    If IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
    If IsLoaded("frm039") Then
        Unload frm039
    End If
    If IsLoaded("frm034") Then
        Unload frm034
    End If
End Function




