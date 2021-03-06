Attribute VB_Name = "Frm006_test"
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
    On Error GoTo Error_handler
    formName = "frm006"
    formID = 6
    
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
    ClearAllFields ThisWorkbook

    ThisWorkbook.Activate
    
    If parameters("run") = 0 Then
         Exit Function
    End If
        
    Select Case parameters("testSubject")
        Case "printsToSpmSheet"
            SetFields
            frm006.OKButton_Click 'Click on Videre button
            CheckFields "SpmSvar"
            
        Case "errorMessage"
            SetFields
            frm006.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.errorMessage
            
        Case "nextStep"
            SetFields
            frm006.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "backButton"
            recHis ("frm005")
            frm006.Tilbage_Click
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "tidligereBesvarelse"
            DataIsSaved "SpmSvar"
            
        Case "noExtraPrints"
            Sheet1.recordChangingCells = True
            SetFields
            If parameters("testParameter") = "noChangeWhenBackButton" Then
                frm006.Tilbage_Click 'Click back button
            Else
                frm006.OKButton_Click 'Click on Videre button
            End If
            CheckNoExtraPrints
            Sheet1.recordChangingCells = False
        Case Else
            MsgBox "Error in 'testsubject' input: tcid " & tcid 'Msgbox to stop the code because you made a mistake in the inputs..
    End Select
    
    'Comparison
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
   
    frm006.OptionButton1.Value = parameters("optionButton1")
    frm006.OptionButton2.Value = parameters("optionButton2")
    frm006.OptionButton3.Value = parameters("optionButton3")
    frm006.OptionButton4.Value = parameters("optionButton4")
    frm006.OptionButton5.Value = parameters("optionButton5")
    frm006.OptionButton6.Value = parameters("optionButton6")
    
End Function
Private Function CheckFields(sheet As String)
    'Check results
    If parameters("optionButton1") = "True" Or parameters("optionButton2") = "True" Then
        result = findPreviousAns(findTopSpm("A"), "6", 1, 1) 'ThisWorkbook.Sheets(sheet).Range("D14").Text
    ElseIf parameters("optionButton3") = "True" Or parameters("optionButton4") = "True" Then
        result = findPreviousAns(findTopSpm("A"), "7", 1, 1) 'ThisWorkbook.Sheets(sheet).Range("D15").Text
    ElseIf parameters("optionButton5") = "True" Or parameters("optionButton6") = "True" Then
        result = findPreviousAns(findTopSpm("A"), "8", 1, 1) ' ThisWorkbook.Sheets(sheet).Range("D16").Text
    End If

End Function
Private Function DataIsSaved(sheet As String)
   
   If parameters("expected") = True Then
        Select Case parameters("testParameter")
           Case "optionButton1"
               Call writeSpmSvar("6", "", "Ja", "", 6) 'ThisWorkbook.Sheets(sheet).Range("D14").Value = "Ja"
               SFunc.ShowFunc (formName)
               result = CStr(frm006.OptionButton1.Value)
           Case "optionButton2"
               Call writeSpmSvar("6", "", "Nej", "", 6) 'ThisWorkbook.Sheets(sheet).Range("D14").Value = "Nej"
               SFunc.ShowFunc (formName)
               result = CStr(frm006.OptionButton2.Value)
           Case "optionButton3"
               Call writeSpmSvar("7", "", "Ja", "", 6) 'ThisWorkbook.Sheets(sheet).Range("D15").Value = "Ja"
               SFunc.ShowFunc (formName)
               result = CStr(frm006.OptionButton3.Value)
            Case "optionButton4"
               Call writeSpmSvar("7", "", "Nej", "", 6) 'ThisWorkbook.Sheets(sheet).Range("D15").Value = "Nej"
               SFunc.ShowFunc (formName)
               result = CStr(frm006.OptionButton4.Value)
           Case "optionButton5"
               Call writeSpmSvar("8", "", "Ja", "", 6) 'ThisWorkbook.Sheets(sheet).Range("D16").Value = "Ja"
               SFunc.ShowFunc (formName)
               result = CStr(frm006.OptionButton5.Value)
           Case "optionButton6"
               Call writeSpmSvar("8", "", "Nej", "", 6) 'ThisWorkbook.Sheets(sheet).Range("D16").Value = "Nej"
               SFunc.ShowFunc (formName)
               result = CStr(frm006.OptionButton6.Value)
        End Select
    Else
        Select Case parameters("testParameter")
           Case "optionButton1"
               ThisWorkbook.Sheets(sheet).Range("D14").Value = ""
               SFunc.ShowFunc (formName)
               result = CStr(frm006.OptionButton1.Value)
           Case "optionButton2"
               ThisWorkbook.Sheets(sheet).Range("D14").Value = ""
               SFunc.ShowFunc (formName)
               result = CStr(frm006.OptionButton2.Value)
           Case "optionButton3"
               ThisWorkbook.Sheets(sheet).Range("D15").Value = ""
               SFunc.ShowFunc (formName)
               result = CStr(frm006.OptionButton3.Value)
            Case "optionButton4"
               ThisWorkbook.Sheets(sheet).Range("D15").Value = ""
               SFunc.ShowFunc (formName)
               result = CStr(frm006.OptionButton4.Value)
           Case "optionButton5"
               ThisWorkbook.Sheets(sheet).Range("D16").Value = ""
               SFunc.ShowFunc (formName)
               result = CStr(frm006.OptionButton5.Value)
           Case "optionButton6"
               ThisWorkbook.Sheets(sheet).Range("D16").Value = ""
               SFunc.ShowFunc (formName)
               result = CStr(frm006.OptionButton6.Value)
        End Select
    End If
    
End Function
'Private Function CheckNoExtraPrints()
'    Select Case parameters("testParameter")
'        'Test different cases were different cells should be changed
'        Case "noChangeWhenError"
'            popCells = Array()
'            rulCells = Array()
'            groCells = Array()
'            spmCells = Array()
'        Case "noChangeWhenBackButton"
'            popCells = Array()
'            rulCells = Array()
'            groCells = Array()
'            spmCells = Array()
'        Case "config1"
'            popCells = Array()
'            rulCells = Array()
'            groCells = Array()
'            spmCells = Array("D14", "D15", "D16", "C14", "C15", "C16")
'    End Select
'
'    'returns a string which shows either true or has the input of the cells that changed that shouldn't have been changed
'    result = Global_Test_Func.CheckPrintsInAllSheets(spmCells, popCells, rulCells, groCells)
'
'    'Cleans up all arrays and dictionaries
'    Erase popCells, rulCells, groCells, spmCells
'    Sheet9.spmChangedCells.RemoveAll
'    Sheet5.groChangedCells.RemoveAll
'    Sheet3.rulChangedCells.RemoveAll
'    Sheet1.popChangedCells.RemoveAll
'End Function
Private Function CheckNoExtraPrints()
    'Check Which configuration to choose
    Select Case parameters("testParameter")
        'Test different cases were different cells should be changed
        Case "noChangeWhenError"
        Case "noChangeWhenBackButton"
        Case "config1"
            Call addSpm("6", "Nej")
            Call addSpm("7", "Nej")
            Call addSpm("8", "Nej")
        Case "config2"
            Call addSpm("6", "Ja")
            Call addSpm("7", "Nej")
            Call addSpm("8", "Nej")
        Case "config3"
            Call addSpm("6", "Ja")
            Call addSpm("7", "Ja")
            Call addSpm("8", "Nej")
        Case "config4"
            Call addSpm("6", "Ja")
            Call addSpm("7", "Nej")
            Call addSpm("8", "Ja")
        Case "config5"
            Call addSpm("6", "Nej")
            Call addSpm("7", "Ja")
            Call addSpm("8", "Nej")
        Case "config6"
            Call addSpm("6", "Nej")
            Call addSpm("7", "Ja")
            Call addSpm("8", "Ja")
        Case "config7"
            Call addSpm("6", "Nej")
            Call addSpm("7", "Nej")
            Call addSpm("8", "Ja")
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
Private Function KillForms()
    'Closes forms
    ThisWorkbook.Activate
    If Global_Test_Func.IsLoaded("frm005") Then
        Unload frm005
    End If
    If Global_Test_Func.IsLoaded("frm006") Then
        Unload frm006
    End If
    If Global_Test_Func.IsLoaded("frm007") Then
        Unload frm007
    End If
    If Global_Test_Func.IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
End Function



