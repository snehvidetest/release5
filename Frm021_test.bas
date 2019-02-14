Attribute VB_Name = "Frm021_test"
Private result As String
Private formID As Integer
Private formName As String
Private stopFormTest As Boolean
Private parameters As Scripting.Dictionary
Private parametersAndCols As Scripting.Dictionary
'Private spmCells As Scripting.Dictionary
Private popCells As Scripting.Dictionary
Private rulCells As Scripting.Dictionary
Private groCells As Scripting.Dictionary

Sub RunTests()
    On Error GoTo Error_handler
    formName = "frm021"
    formID = 21
    
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
    

    ThisWorkbook.Activate
    
    If parameters("run") = 0 Then
         Exit Function
    End If
        
    Select Case parameters("testSubject")
        Case "printsToSpmSheet"
            SetFields
            frm021.OKButton_Click 'Click on Videre button
            Select Case parameters("testParameter")
                Case "textbox1"
                    'CheckFields "SpmSvar", "D55"
                    result = findPreviousAns(findTopSpm("A"), "12", 1, 1)
                Case "checkbox1"
                    'CheckFields "SpmSvar", "D55"
                    result = findPreviousAns(findTopSpm("A"), "12", 1, 1)
            End Select
            
        Case "printsToRulSheet"
            SetFields
            frm021.OKButton_Click 'Click on Videre button
            If (parameters("testParameter") = "ruleActivation") Then
                Select Case parameters("rule")
                    Case "R0072"
                        CheckFields "Regler", "G73"
                    Case "R0073"
                        CheckFields "Regler", "G74"
                    Case "R0074"
                        CheckFields "Regler", "G76"
                    Case "R0103"
                        CheckFields "Regler", "G75"
                 End Select
            ElseIf (parameters("testParameter") = "amount") Then
                Select Case parameters("rule")
                    Case "R0072"
                        CheckFields "Regler", "H73"
                    Case "R0073"
                        CheckFields "Regler", "H74"
                 End Select
            End If
            
        Case "printsToGroSheet"
            SetFields
            frm021.OKButton_Click 'Click on Videre button
            Select Case parameters("group")
                Case "G0005"
                    CheckFields "Gruppering", "C6"
                Case "G0006"
                    CheckFields "Gruppering", "C7"
            End Select
            
        Case "errorMessage"
            SetFields
            frm021.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.errorMessage
            
        Case "nextStep"
            SetFields
            frm021.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "backButton"
            If (parameters("testParameter") = "frm037") Then
                recHis ("frm037")
                frm039.CheckBox4.Value = True
            Else
                recHis ("frm038")
                frm039.CheckBox4.Value = False
            End If
            frm021.Tilbage_Click
            result = Global_Test_Func.NextStep(parameters("expected"))
        Case "tidligereBesvarelse"
            DataIsSaved "SpmSvar"
            
        Case "noExtraPrints"
            Sheet1.recordChangingCells = True
            SetFields
            frm021.OKButton_Click 'Click on Videre button
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
   
    frm021.TextBox1.Value = parameters("textbox1")
    frm021.CheckBox1.Value = parameters("checkbox1")
    
End Function
Private Function CheckFields(sheet As String, cell As String)
    'Check results
    result = ThisWorkbook.Sheets(sheet).Range(cell).Text

End Function
Private Function DataIsSaved(sheet As String)
   
   If parameters("expected") = True Then
        Select Case parameters("testParameter")
           Case "textbox1"
               'ThisWorkbook.Sheets(sheet).Range("D55").Value = "10"
               Call writeSpmSvar("12", "", "10", "", 6)
               ShowFunc (formName)
               result = CStr(frm021.TextBox1.Value)
           Case "checkbox1"
                'ThisWorkbook.Sheets(sheet).Range("D55").Value = "Ved ikke"
                If parameters("checkbox1") Then
                    Call writeSpmSvar("12", "", "Ved ikke", "", 6)
                End If
                result = CStr(frm021.CheckBox1.Value)
        End Select
    Else
        Select Case parameters("testParameter")
           Case "textbox1"
               'ThisWorkbook.Sheets(sheet).Range("D55").Value = ""
               ShowFunc (formName)
               result = CStr(frm021.TextBox1.Value)
           Case "checkbox1"
               'ThisWorkbook.Sheets(sheet).Range("D55").Value = ""
               result = CStr(frm021.CheckBox1.Value)
               Debug.Print (result)
               Debug.Print ("hej")
               Debug.Print (result)
        End Select
    End If
    
End Function

Private Function CheckNoExtraPrints()
    'Check Which configuration to choose
    Select Case parameters("testParameter")
        'Test different cases were different cells should be changed
        Case "noChangeWhenError"
        Case "noChangeWhenBackButton"
            rulCells.Add "G73", "NEJ"
            rulCells.Add "G74", "NEJ"
            rulCells.Add "G75", "NEJ"
            rulCells.Add "G76", "NEJ"
            rulCells.Add "H74", ""
            rulCells.Add "H75", ""
            
            groCells.Add "C6", "NEJ"
            groCells.Add "C7", "NEJ"
            
        Case "config1"
            rulCells.Add "G73", "JA"
            rulCells.Add "G74", "JA"
            rulCells.Add "G75", "NEJ"
            rulCells.Add "G76", "NEJ"
            rulCells.Add "H73", parameters("textbox1")
            rulCells.Add "H74", parameters("textbox1")
            
            groCells.Add "C6", "JA"
            groCells.Add "C7", "NEJ"
            
            Call addSpm("12", parameters("textbox1"), "kr.")
            
        Case "config2"
            rulCells.Add "G73", "NEJ"
            rulCells.Add "G74", "NEJ"
            rulCells.Add "G75", "NEJ"
            rulCells.Add "G76", "NEJ"
            rulCells.Add "H73", parameters("textbox1")
            rulCells.Add "H74", parameters("textbox1")
            
            groCells.Add "C6", "NEJ"
            groCells.Add "C7", "NEJ"
            Call addSpm("12", "Ved ikke")
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
'            rulCells = Array("G73", "G74", "G75", "G76", "H73", "H74")
'            groCells = Array("C6", "C7")
'            spmCells = Array("D55", "C55")
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
'
'End Function
Private Function KillForms()
    'Closes forms
    ThisWorkbook.Activate
    If Global_Test_Func.IsLoaded("frm021") Then
        Unload frm021
    End If
    If Global_Test_Func.IsLoaded("frm022") Then
        Unload frm022
    End If
    If Global_Test_Func.IsLoaded("frm037") Then
        Unload frm037
    End If
    If Global_Test_Func.IsLoaded("frm038") Then
        Unload frm038
    End If
    If Global_Test_Func.IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
End Function

