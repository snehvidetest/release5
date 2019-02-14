Attribute VB_Name = "Frm003_test"
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
    formName = "frm003"
    formID = 3
    
    Set parametersAndCols = Global_Test_Func.getParamtersAndTheirCols(formID)
    
    Set spmCells = New Scripting.Dictionary   'Dictionaries to log changes in SpmSvar sheet
    Set popCells = New Scripting.Dictionary   'Dictionaries to log changes in Population sheet
    Set groCells = New Scripting.Dictionary   'Dictionaries to log changes in Gruppering sheet
    Set rulCells = New Scripting.Dictionary   'Dictionaries to log changes in Regler sheet

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
            frm003.OKButton_Click 'Click on Videre button
            'CheckFields "SpmSvar", "D6"
            result = findPreviousAns(findTopSpm("A"), "4.a", 1, 1)
            
        Case "checkCaption"
            ThisWorkbook.Activate
            Select Case parameters("testParameter")
                Case "optionButton1"
                    result = frm003.OptionButton1.caption
                Case "optionButton2"
                    result = frm003.OptionButton2.caption
                Case "optionButton3"
                    result = frm003.OptionButton3.caption
            End Select
            
        Case "errorMessage"
            SetFields
            frm003.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.errorMessage
            
        Case "nextStep"
            SetFields
            frm003.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "backButton"
            recHis ("frm002")
            frm003.Tilbage_Click
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "tidligereBesvarelse"
            DataIsSaved "SpmSvar", "D6" 'findPreviousAns(500, "4.a", 1, 1)
            
            
        Case "noExtraPrints"
            SFunc.ShowFunc ("frm002")
            Global_Test_Func.resetSheets ThisWorkbook
            
            Sheet1.recordChangingCells = True
            SetFields
            If parameters("testParameter") = "noChangeWhenBackButton" Then
                frm003.Tilbage_Click 'Click back button
            Else
                frm003.OKButton_Click 'Click on Videre button
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
   
    ThisWorkbook.Activate
   
    frm003.OptionButton1.Value = parameters("optionButton1")
    frm003.OptionButton2.Value = parameters("optionButton2")
    frm003.OptionButton3.Value = parameters("optionButton3")
    

End Function
Private Function CheckFields(sheet As String, cell As String)
   'Check results
    result = ThisWorkbook.Sheets(sheet).Range(cell).Text

End Function
Private Function DataIsSaved(sheet As String, cell As String)
   
   If parameters("expected") = True Then
        Select Case parameters("testParameter")
           Case "optionButton1"
               'ThisWorkbook.Sheets(sheet).Range(cell).Value = "At der enten foretages en tilpasning af den allerede afgrænsede modtagelsesperiode"
               Call writeSpmSvar("4.a", "", "At der foretages en tilpasning af den allerede afgrænsede modtagelsesperiode, eller", "", 6)
               ShowFunc (formName)
               result = CStr(frm003.OptionButton1.Value)
           Case "optionButton2"
               'ThisWorkbook.Sheets(sheet).Range(cell).Value = "At der foretages en periodemæssig afgrænsning af den allerede afgrænsede modtagelsesperiode via ét eller flere af stamdatafelterne"
               Call writeSpmSvar("4.a", "", "At der foretages en periodemæssig afgrænsning af den allerede afgrænsede modtagelsesperiode via ét eller flere af stamdatafelterne", "", 6)
               ShowFunc (formName)
               result = CStr(frm003.OptionButton2.Value)
           Case "optionButton3"
               'ThisWorkbook.Sheets(sheet).Range(cell).Value = "Nej/Ved ikke"
               Call writeSpmSvar("4.a", "", "Nej/Ved ikke", "", 6)
               ShowFunc (formName)
               result = CStr(frm003.OptionButton3.Value)
        End Select
    Else
        Select Case parameters("testParameter")
           Case "optionButton1"
               ThisWorkbook.Sheets(sheet).Range(cell).Value = ""
               ShowFunc (formName)
               result = CStr(frm003.OptionButton1.Value)
           Case "optionButton2"
               ThisWorkbook.Sheets(sheet).Range(cell).Value = ""
               ShowFunc (formName)
               result = CStr(frm003.OptionButton2.Value)
           Case "optionButton3"
               ThisWorkbook.Sheets(sheet).Range(cell).Value = ""
               ShowFunc (formName)
               result = CStr(frm003.OptionButton3.Value)
        End Select
    End If
    
End Function

Private Function CheckNoExtraPrints()
    'Check Which configuration to choose
    Select Case parameters("testParameter")
        'Test different cases were different cells should be changed
        Case "noChangeWhenError"
        Case "noChangeWhenBackButton"
            'spmCells.Add "D6", ""
            Call addSpm("4.a", "", "")
        Case "config1"
            'spmCells.Add "D6", ""
            Call addSpm("4.a", "At der foretages en tilpasning af den allerede afgrænsede modtagelsesperiode, eller")
        Case "config2"
            'spmCells.Add "D6", "At der foretages en periodemæssig afgrænsning af den allerede afgrænsede modtagelsesperiode via ét eller flere af stamdatafelterne"
            Call addSpm("4.a", "At der foretages en periodemæssig afgrænsning af den allerede afgrænsede modtagelsesperiode via ét eller flere af stamdatafelterne")
        Case "config3"
            'spmCells.Add "D6", "Nej/Ved ikke"
            Call addSpm("4.a", "Nej/Ved ikke")
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
     
    If Global_Test_Func.IsLoaded("frm002") Then
        Unload frm002
    End If
    If Global_Test_Func.IsLoaded("frm003") Then
        Unload frm003
    End If
    If Global_Test_Func.IsLoaded("frm004") Then
        Unload frm004
    End If
    If Global_Test_Func.IsLoaded("frm026") Then
        Unload frm026
    End If
    If Global_Test_Func.IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
    
    ThisWorkbook.Activate
End Function



