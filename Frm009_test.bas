Attribute VB_Name = "Frm009_test"
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
    
    formName = "frm009"
    formID = 9
    
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
            frm009.OKButton_Click 'Click on Videre button
            CheckFields "Gruppering"
            
        Case "printsToPopSheet"
            SetFields
            frm009.OKButton_Click 'Click on Videre button
            CheckFields "Population"
            
        Case "printsToSpmSheet"
            SetFields
            frm009.OKButton_Click 'Click on Videre button
            CheckFields "SpmSvar"
            
        Case "errorMessage"
            SetFields
            frm009.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.errorMessage
            
        Case "nextStep"
            SetFields
            frm009.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "backButton"
            recHis ("frm008")
            frm009.Tilbage_Click
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "tidligereBesvarelse"
            DataIsSaved "SpmSvar", "D19"
            
        Case "noExtraPrints"
            Sheet1.recordChangingCells = True
            SetFields
            
            If parameters("testParameter") = "noChangeWhenBackButton" Then
                frm009.Tilbage_Click 'Click back button
            Else
                frm009.OKButton_Click 'Click on Videre button
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
    frm009.OptionButton1.Value = parameters("optionButton1")
    frm009.OptionButton2.Value = parameters("optionButton2")

End Function

Private Function CheckFields(sheet As String)

    'Check results
    If (sheet = "SpmSvar") Then
        'result = ThisWorkbook.Sheets(sheet).Range("D19").Text
        result = findPreviousAns(findTopSpm("A"), "9.a.2", 1, 1)
    ElseIf (sheet = "Population") Then
        Select Case parameters("testParameter")
            Case "trustRIM"
                result = ThisWorkbook.Sheets(sheet).Range("B16").Text
            Case "rimFOKO"
                result = ThisWorkbook.Sheets(sheet).Range("B17").Text
        End Select
        
    ElseIf (sheet = "Gruppering") Then
        Select Case parameters("group")
            Case "G0001"
                result = ThisWorkbook.Sheets(sheet).Range("C2").Text
            Case "G0002"
                result = ThisWorkbook.Sheets(sheet).Range("C3").Text
        End Select
    End If
End Function


Function DataIsSaved(sheet As String, cell As String)
    
    
    Select Case parameters("testParameter")
        Case "optionButton1"
            
            If parameters("optionButton1") = "True" Then
                'ThisWorkbook.Sheets(sheet).Range(cell).Value = "Ja"
                Call writeSpmSvar("9.a.2", "", "Ja", "", 6)
            ElseIf parameters("optionButton1") = "False" Then
                'ThisWorkbook.Sheets(sheet).Range(cell).Value = ""
            End If
            
            ShowFunc (formName)
            result = CStr(frm009.OptionButton1.Value)
            
        Case "optionButton2"
            If parameters("optionButton2") = "True" Then
                'ThisWorkbook.Sheets(sheet).Range(cell).Value = "Nej"
                Call writeSpmSvar("9.a.2", "", "Nej", "", 6)
            ElseIf parameters("optionButton2") = "False" Then
                'ThisWorkbook.Sheets(sheet).Range(cell).Value = ""
            End If
            
            ShowFunc (formName)
            result = CStr(frm009.OptionButton2.Value)
            
    End Select
            
            
End Function
Private Function CheckNoExtraPrints()
    'Check Which configuration to choose
    Select Case parameters("testParameter")
        'Test different cases were different cells should be changed
        Case "noChangeWhenError"
        Case "noChangeWhenBackButton"
            popCells.Add "B16", ""
            popCells.Add "B17", ""
            
            groCells.Add "C2", "NEJ"
            groCells.Add "C3", "NEJ"
        Case "config1"
            popCells.Add "B16", "JA"
            popCells.Add "B17", "NEJ"
            
            groCells.Add "C2", "NEJ"
            groCells.Add "C3", "JA"
            
            Call addSpm("9.a.2", "Ja")
        Case "config2"
            Call addSpm("9.a.2", "Nej")
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
'        Case "noChangeWhenBackButton"
'            popCells = Array()
'            rulCells = Array()
'            groCells = Array()
'            spmCells = Array()
'        Case "config1"
'            popCells = Array("B16", "B17")
'            rulCells = Array()
'            groCells = Array("C2", "C3")
'            spmCells = Array("D19", "C19")
'        Case "config2"
'            popCells = Array()
'            rulCells = Array()
'            groCells = Array()
'            spmCells = Array("D19", "C19")
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
    If IsLoaded("frm008") Then
        Unload frm008
    End If
    If IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
    If IsLoaded("frm007") Then
        Unload frm010
    End If
    If IsLoaded("frm009") Then
        Unload frm009
    End If
    If IsLoaded("frm010") Then
        Unload frm010
    End If
    If IsLoaded("frm039") Then
        Unload frm039
    End If
End Function


