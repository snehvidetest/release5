Attribute VB_Name = "Frm037_test"
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
    formName = "frm037"
    formID = 37
    
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
            frm037.OKButton_Click 'Click on Videre button
            'CheckFields "SpmSvar"
            Select Case parameters("testParameter")
                Case "textbox1"
                    result = findPreviousAns(findTopSpm("A"), "11.a_1", 1, 1)
                Case "textbox2"
                    result = findPreviousAns(findTopSpm("A"), "11.a_1", 2, 1)
            End Select
            
        Case "printsToRulSheet"
            SetFields
            frm037.OKButton_Click 'Click on Videre button
            CheckFields "Regler"
            
        Case "errorMessage"
            SetFields
            Debug.Print (frm037.ComboBox2.Value)
            frm037.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.errorMessage
            
        Case "nextStep"
            SetFields
            frm037.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "backButton"
            recHis ("frm039")
            frm037.Tilbage_Click
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "tidligereBesvarelse"
            DataIsSaved "SpmSvar"
            
        Case "noExtraPrints"
            Sheet1.recordChangingCells = True
            SetFields
            If parameters("testParameter") = "noChangeWhenBackButton" Then
                frm037.Tilbage_Click 'Click back button
            Else
                frm037.OKButton_Click 'Click on Videre button
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
    frm037.TextBox1.Value = parameters("textbox1")
    frm037.TextBox2.Value = parameters("textbox2")
    frm037.ComboBox2.Value = parameters("combobox2")
    frm037.ComboBox4.Value = parameters("combobox4")
End Function

Private Function CheckFields(sheet As String)
    Select Case parameters("testParameter")
        Case "textbox1"
            result = ThisWorkbook.Sheets(sheet).Range("D63").Text
        Case "textbox2"
            result = ThisWorkbook.Sheets(sheet).Range("G63").Text
        Case "combobox2"
            result = ThisWorkbook.Sheets(sheet).Range("F63").Text
        Case "combobox4"
            result = ThisWorkbook.Sheets(sheet).Range("I63").Text
        Case "ruleActivation"
            result = ThisWorkbook.Sheets(sheet).Range("G21").Text
        Case "ruleXDays"
            result = ThisWorkbook.Sheets(sheet).Range("J21").Text
        Case "ruleYDays"
            result = ThisWorkbook.Sheets(sheet).Range("M21").Text
    End Select
    
    
End Function


Function DataIsSaved(sheet As String)
    If parameters("expected") = True Then
        Select Case parameters("testParameter")
           Case "textbox1"
               'ThisWorkbook.Sheets(sheet).Range("D63").Value = "10"
                Call writeSpmSvar("11.a_5", "", "fra 10 dage efter", "til 100 dage efter", 6)
               ShowFunc (formName)
               result = CStr(frm037.TextBox1.Value)
           Case "textbox2"
               'ThisWorkbook.Sheets(sheet).Range("G63").Value = "100"
                Call writeSpmSvar("11.a_5", "", "fra 10 dage efter", "til 100 dage efter", 6)
               result = CStr(frm037.TextBox2.Value)
           Case "combobox2"
               'ThisWorkbook.Sheets(sheet).Range("F63").Value = "efter"
                Call writeSpmSvar("11.a_5", "", "fra 10 dage efter", "til 100 dage efter", 6)
               ShowFunc (formName)
               result = CStr(frm037.ComboBox2.Value)
            Case "ombobox4"
               'ThisWorkbook.Sheets(sheet).Range("I63").Value = "efter"
                Call writeSpmSvar("11.a_5", "", "fra 10 dage efter", "til 100 dage efter", 6)
               ShowFunc (formName)
               result = CStr(frm037.ComboBox4.Value)
        End Select
    Else
        Select Case parameters("testParameter")
           Case "textbox1"
               'ThisWorkbook.Sheets(sheet).Range("D63").Value = ""
               ShowFunc (formName)
               result = CStr(frm037.TextBox1.Value)
           Case "textbox2"
               'ThisWorkbook.Sheets(sheet).Range("G63").Value = ""
               result = CStr(frm037.TextBox2.Value)
           Case "combobox2"
               'ThisWorkbook.Sheets(sheet).Range("F63").Value = ""
               ShowFunc (formName)
               result = CStr(frm037.ComboBox2.Value)
            Case "combobox4"
               'ThisWorkbook.Sheets(sheet).Range("I63").Value = ""
               ShowFunc (formName)
               result = CStr(frm037.ComboBox4.Value)
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
            rulCells.Add "G21", "JA"
            rulCells.Add "J21", parameters("textbox1")
            rulCells.Add "M21", parameters("textbox2")
            
            Call addSpm("11.a_1", parameters("textbox1"), parameters("textbox2"))
        Case "config2"
            rulCells.Add "G21", "JA"
            rulCells.Add "J21", parameters("textbox1")
            rulCells.Add "M21", parameters("textbox2")
            
            Call addSpm("11.a_1", parameters("textbox1"), parameters("textbox2"))
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
'            rulCells = Array("G21", "J21", "M21")
'            groCells = Array()
'            spmCells = Array("D63", "F63", "G63", "I63", "C63")
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
    If IsLoaded("frm037") Then
        Unload frm037
    End If
    If IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
    If IsLoaded("frm036") Then
        Unload frm036
    End If
    If IsLoaded("frm038") Then
        Unload frm038
    End If
    If IsLoaded("frm047") Then
        Unload frm047
    End If
End Function
