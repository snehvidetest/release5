Attribute VB_Name = "Frm035_test"
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
    formName = "frm035"
    formID = 35
    
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
    
        Case "printsToSpmSheet"
            SetFields
            frm035.OKButton_Click 'Click on Videre button
            'CheckFields "SpmSvar"
            Select Case parameters("testParameter")
                Case "textbox1"
                    result = findPreviousAns(findTopSpm("A"), "11.a_3", 1, 1)
                Case "textbox2"
                    result = findPreviousAns(findTopSpm("A"), "11.a_3", 2, 1)
            End Select
            
        Case "printsToRulSheet"
            SetFields
            frm035.OKButton_Click 'Click on Videre button
            CheckFields "Regler"
            
        Case "errorMessage"
            SetFields

            frm035.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.errorMessage
            
        Case "nextStep"
            SetFields
            frm035.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "backButton"
            recHis ("frm039")
            frm035.Tilbage_Click
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "tidligereBesvarelse"
            DataIsSaved "SpmSvar"
            
        Case "noExtraPrints"
            Sheet1.recordChangingCells = True
            SetFields
            If parameters("testParameter") = "noChangeWhenBackButton" Then
                frm035.Tilbage_Click 'Click back button
            Else
                frm035.OKButton_Click 'Click on Videre button
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
    frm035.TextBox1.Value = parameters("textbox1")
    frm035.TextBox2.Value = parameters("textbox2")
    frm035.ComboBox2.Value = parameters("combobox2")
    frm035.ComboBox4.Value = parameters("combobox4")
    
End Function

Private Function CheckFields(sheet As String)
    Select Case parameters("testParameter")
        Case "textbox1"
            result = ThisWorkbook.Sheets(sheet).Range("D61").Text
        Case "textbox2"
            result = ThisWorkbook.Sheets(sheet).Range("G61").Text
        Case "combobox2"
            result = ThisWorkbook.Sheets(sheet).Range("F61").Text
        Case "combobox4"
            result = ThisWorkbook.Sheets(sheet).Range("I61").Text
        Case "ruleActivation"
            result = ThisWorkbook.Sheets(sheet).Range("G15").Text
        Case "ruleXDays"
            result = ThisWorkbook.Sheets(sheet).Range("J15").Text
        Case "ruleYDays"
            result = ThisWorkbook.Sheets(sheet).Range("M15").Text
    End Select
    
    
End Function


Function DataIsSaved(sheet As String)
    If parameters("expected") = True Then
        Select Case parameters("testParameter")
           Case "textbox1"
               'ThisWorkbook.Sheets(sheet).Range("D61").Value = "10"
                Call writeSpmSvar("11.a_3", "", "fra 10 dage efter", "til 100 dage efter", 6)
               ShowFunc (formName)
               result = CStr(frm035.TextBox1.Value)
           Case "textbox2"
               'ThisWorkbook.Sheets(sheet).Range("G61").Value = "100"
                Call writeSpmSvar("11.a_3", "", "fra 10 dage efter", "til 100 dage efter", 6)
               result = CStr(frm035.TextBox2.Value)
           Case "combobox2"
               'ThisWorkbook.Sheets(sheet).Range("F61").Value = "efter"
                Call writeSpmSvar("11.a_3", "", "fra 10 dage efter", "til 100 dage efter", 6)
               ShowFunc (formName)
               result = CStr(frm035.ComboBox2.Value)
            Case "combobox4"
               'ThisWorkbook.Sheets(sheet).Range("I61").Value = "efter"
                Call writeSpmSvar("11.a_3", "", "fra 10 dage efter", "til 100 dage efter", 6)
               ShowFunc (formName)
               result = CStr(frm035.ComboBox4.Value)
        End Select
    Else
        Select Case parameters("testParameter")
           Case "textbox1"
               'ThisWorkbook.Sheets(sheet).Range("D61").Value = ""
               ShowFunc (formName)
               result = CStr(frm035.TextBox1.Value)
           Case "textbox2"
               'ThisWorkbook.Sheets(sheet).Range("G61").Value = ""
               result = CStr(frm035.TextBox2.Value)
           Case "combobox2"
               'ThisWorkbook.Sheets(sheet).Range("F61").Value = ""
               ShowFunc (formName)
               result = CStr(frm035.ComboBox2.Value)
            Case "ombobox4"
               'ThisWorkbook.Sheets(sheet).Range("I61").Value = ""
               ShowFunc (formName)
               result = CStr(frm035.ComboBox4.Value)
        End Select
    End If
    
End Function
Private Function CheckNoExtraPrints()
    'Check Which configuration to choose
    Select Case parameters("testParameter")
        'Test different cases were different cells should be changed
        Case "noChangeWhenError"
        Case "noChangeWhenBackButton"
        Case "config1"
            rulCells.Add "G15", "JA"
            rulCells.Add "J15", parameters("textbox1")
            rulCells.Add "M15", parameters("textbox2")
            
            Call addSpm("11.a_3", parameters("textbox1"), parameters("textbox2"))
        Case "config2"
            rulCells.Add "G15", "JA"
            rulCells.Add "J15", parameters("textbox1")
            rulCells.Add "M15", parameters("textbox2")
            
            Call addSpm("11.a_3", parameters("textbox1"), parameters("textbox2"))
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
'Private Function CheckNoExtraPrints()
'    Select Case parameters("testParameter")
'        'Test different cases were different cells should be changed
'        Case "noChangeWhenError"
'            popCells = Array()
'            rulCells = Array()
'            groCells = Array()
'            spmCells = Array()
'        Case "config1"
'            popCells = Array()
'            rulCells = Array("G15", "J15", "M15")
'            groCells = Array()
'            spmCells = Array("D61", "F61", "G61", "I61", "C61")
'    End Select
'
'    'returns a string which shows either true or has the input of the cells that changed that shouldn't have been changed
'    result = Global_Test_Func.CheckPrintsInAllSheets(spmCells, popCells, rulCells, groCells)
'
'     'Cleans up all arrays and dictionaries
'    Erase popCells, rulCells, groCells, spmCells
'    Sheet9.spmChangedCells.RemoveAll
'    Sheet5.groChangedCells.RemoveAll
'    Sheet3.rulChangedCells.RemoveAll
'    Sheet1.popChangedCells.RemoveAll
'End Function
Function KillForms()
    'Closes forms
    ThisWorkbook.Activate
    If IsLoaded("frm035") Then
        Unload frm035
    End If
    If IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
    If IsLoaded("frm036") Then
        Unload frm036
    End If
    If IsLoaded("frm034") Then
        Unload frm034
    End If
    If IsLoaded("frm045") Then
        Unload frm045
    End If
End Function
