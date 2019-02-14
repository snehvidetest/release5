Attribute VB_Name = "Frm026_test"
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
    formName = "frm026"
    formID = 26

    
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
        
        Case "printsToPopSheet"
            SetFields
            frm026.OKButton_Click 'Click on Videre button
            CheckFields "Population"
            
        Case "printsToSpmSheet"
            SetFields
            frm026.OKButton_Click 'Click on Videre button
            CheckFields "SpmSvar"
            
        Case "errorMessage"
            SetFields
            frm026.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.errorMessage
        
            
        Case "nextStep"
            SetFields
            frm026.OKButton_Click 'Click on Videre button
            result = Global_Test_Func.NextStep(parameters("expected"))

            
        Case "backButton"
            recHis ("frm003")
            frm026.Tilbage_Click
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "tidligereBesvarelse"
            DataIsSaved "SpmSvar"
            
        Case "noExtraPrints"
            Sheet1.recordChangingCells = True
            SetFields
            If parameters("testParameter") = "noChangeWhenBackButton" Then
                frm026.Tilbage_Click 'Click back button
            Else
                frm026.OKButton_Click 'Click on Videre button
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
    frm026.Forfaldsdato.Value = parameters("forfaldsdato")
    frm026.txtFFStart.Value = parameters("forfaldsdatoFrom")
    frm026.txtFFSlut.Value = parameters("forfaldsdatoTo")
    
    frm026.SRB.Value = parameters("srb")
    frm026.txtSRBstart.Value = parameters("srbFrom")
    frm026.txtSRBslut.Value = parameters("srbTo")
    
    frm026.Stiftelsesdato.Value = parameters("stiftelsesdato")
    frm026.txtSTIstart.Value = parameters("stiftelsesdatoFrom")
    frm026.txtSTIslut.Value = parameters("stiftelsesdatoTo")
    
    frm026.PeriodeStartdato.Value = parameters("periodeStart")
    frm026.txtPSTstart.Value = parameters("periodeStartFrom")
    frm026.txtPSTslut.Value = parameters("periodeStartTo")
    
    frm026.PeriodeSlutdato.Value = parameters("periodeSlut")
    frm026.txtPSLstart.Value = parameters("periodeSlutFrom")
    frm026.txtPSLslut.Value = parameters("periodeSlutTo")
    
End Function

Private Function CheckFields(sheet As String)

    Select Case sheet
        Case "SpmSvar"
            Select Case parameters("testParameter")
                Case "forfaldsdato"
                    'result = ThisWorkbook.Sheets(sheet).Range("D8").Text
                    result = findPreviousAns(findTopSpm("A"), "4.a.2.1_1", 0, 1)
                Case "forfaldsdatoFrom"
                    'result = ThisWorkbook.Sheets(sheet).Range("E8").Text
                    result = findPreviousAns(findTopSpm("A"), "4.a.2.1_1", 1, 1)
                Case "forfaldsdatoTo"
                    'result = ThisWorkbook.Sheets(sheet).Range("F8").Text
                    result = findPreviousAns(findTopSpm("A"), "4.a.2.1_1", 2, 1)
                Case "srb"
                    'result = ThisWorkbook.Sheets(sheet).Range("D9").Text
                    result = findPreviousAns(findTopSpm("A"), "4.a.2.1_2", 0, 1)
                Case "srbFrom"
                    'result = ThisWorkbook.Sheets(sheet).Range("E9").Text
                    result = findPreviousAns(findTopSpm("A"), "4.a.2.1_2", 1, 1)
                Case "srbTo"
                    'result = ThisWorkbook.Sheets(sheet).Range("F9").Text
                    result = findPreviousAns(findTopSpm("A"), "4.a.2.1_2", 2, 1)
                Case "stiftelsesdato"
                    'result = ThisWorkbook.Sheets(sheet).Range("D10").Text
                    result = findPreviousAns(findTopSpm("A"), "4.a.2.1_3", 0, 1)
                Case "stiftelsesdatoFrom"
                    'result = ThisWorkbook.Sheets(sheet).Range("E10").Text
                    result = findPreviousAns(findTopSpm("A"), "4.a.2.1_3", 1, 1)
                Case "stiftelsesdatoTo"
                    'result = ThisWorkbook.Sheets(sheet).Range("F10").Text
                    result = findPreviousAns(findTopSpm("A"), "4.a.2.1_3", 2, 1)
                Case "periodeStart"
                    'result = ThisWorkbook.Sheets(sheet).Range("D11").Text
                    result = findPreviousAns(findTopSpm("A"), "4.a.2.1_4", 0, 1)
                Case "periodeStartFrom"
                    'result = ThisWorkbook.Sheets(sheet).Range("E11").Text
                    result = findPreviousAns(findTopSpm("A"), "4.a.2.1_4", 1, 1)
                Case "periodeStartTo"
                    'result = ThisWorkbook.Sheets(sheet).Range("F11").Text
                    result = findPreviousAns(findTopSpm("A"), "4.a.2.1_4", 2, 1)
                Case "periodeSlut"
                    'result = ThisWorkbook.Sheets(sheet).Range("D12").Text
                    result = findPreviousAns(findTopSpm("A"), "4.a.2.1_5", 0, 1)
                Case "periodeSlutFrom"
                    'result = ThisWorkbook.Sheets(sheet).Range("E12").Text
                    result = findPreviousAns(findTopSpm("A"), "4.a.2.1_5", 1, 1)
                Case "periodeSlutTo"
                    'result = ThisWorkbook.Sheets(sheet).Range("F12").Text
                    result = findPreviousAns(findTopSpm("A"), "4.a.2.1_5", 2, 1)
            End Select
            
        Case "Population"
            Select Case parameters("testParameter")
                Case "forfaldsdatoFrom"
                    result = ThisWorkbook.Sheets(sheet).Range("B6").Text
                Case "forfaldsdatoTo"
                    result = ThisWorkbook.Sheets(sheet).Range("B7").Text
                Case "srbFrom"
                    result = ThisWorkbook.Sheets(sheet).Range("B8").Text
                Case "srbTo"
                    result = ThisWorkbook.Sheets(sheet).Range("B9").Text
                Case "stiftelsesdatoFrom"
                    result = ThisWorkbook.Sheets(sheet).Range("B10").Text
                Case "stiftelsesdatoTo"
                    result = ThisWorkbook.Sheets(sheet).Range("B11").Text
                Case "periodeStartFrom"
                    result = ThisWorkbook.Sheets(sheet).Range("B12").Text
                Case "periodeStartTo"
                    result = ThisWorkbook.Sheets(sheet).Range("B13").Text
                Case "periodeSlutFrom"
                    result = ThisWorkbook.Sheets(sheet).Range("B14").Text
                Case "periodeSlutTo"
                    result = ThisWorkbook.Sheets(sheet).Range("B15").Text
            End Select
            
        End Select
        
End Function


Function DataIsSaved(sheet As String)

    If parameters("forfaldsdato") = True Then
'        ThisWorkbook.Sheets(sheet).Range("D8").Value = "Forfaldsdato"
'        ThisWorkbook.Sheets(sheet).Range("E8").Value = parameters("forfaldsdatoFrom")
'        ThisWorkbook.Sheets(sheet).Range("F8").Value = parameters("forfaldsdatoTo")
        Call writeSpmSvar("4.a.2.1_1", "Forfaldsdato", parameters("forfaldsdatoFrom"), parameters("forfaldsdatoTo"), 6)
    End If
        
    If parameters("srb") = True Then
'        ThisWorkbook.Sheets(sheet).Range("D9").Value = "SRB Dato"
'        ThisWorkbook.Sheets(sheet).Range("E9").Value = parameters("srbFrom")
'        ThisWorkbook.Sheets(sheet).Range("F9").Value = parameters("srbTo")
        Call writeSpmSvar("4.a.2.1_2", "SRB Dato", parameters("forfaldsdatoFrom"), parameters("forfaldsdatoTo"), 6)
    End If
    
    If parameters("stiftelsesdato") = True Then
'        ThisWorkbook.Sheets(sheet).Range("D10").Value = "Stiftelsesdato"
'        ThisWorkbook.Sheets(sheet).Range("E10").Value = parameters("stiftelsesdatoFrom")
'        ThisWorkbook.Sheets(sheet).Range("F10").Value = parameters("stiftelsesdatoTo")
        Call writeSpmSvar("4.a.2.1_3", "Stiftelsesdato", parameters("forfaldsdatoFrom"), parameters("forfaldsdatoTo"), 6)
    End If
    
    If parameters("periodeStart") = True Then
'        ThisWorkbook.Sheets(sheet).Range("D11").Value = "PeriodeStartdato"
'        ThisWorkbook.Sheets(sheet).Range("E11").Value = parameters("periodeStartFrom")
'        ThisWorkbook.Sheets(sheet).Range("F11").Value = parameters("periodeStartTo")
        Call writeSpmSvar("4.a.2.1_4", "Periode startdato", parameters("forfaldsdatoFrom"), parameters("forfaldsdatoTo"), 6)
    End If
    
    If parameters("periodeSlut") = True Then
'        ThisWorkbook.Sheets(sheet).Range("D12").Value = "PeriodeSlutdato"
'        ThisWorkbook.Sheets(sheet).Range("E12").Value = parameters("periodeSlutFrom")
'        ThisWorkbook.Sheets(sheet).Range("F12").Value = parameters("periodeSlutTo")
        Call writeSpmSvar("4.a.2.1_5", "Periode slutdato", parameters("forfaldsdatoFrom"), parameters("forfaldsdatoTo"), 6)
    End If
            
    ShowFunc (formName)
    
    Select Case parameters("testParameter")
        Case "forfaldsdato"
            result = CStr(frm026.Forfaldsdato.Value)
        Case "forfaldsdatoFrom"
            result = CStr(frm026.txtFFStart.Value)
        Case "forfaldsdatoTo"
            result = CStr(frm026.txtFFSlut.Value)
        Case "srb"
            result = CStr(frm026.SRB.Value)
        Case "srbFrom"
            result = CStr(frm026.txtSRBstart.Value)
        Case "srbTo"
            result = CStr(frm026.txtSRBslut.Value)
        Case "stiftelsesdato"
            result = CStr(frm026.Stiftelsesdato.Value)
        Case "stiftelsesdatoFrom"
            result = CStr(frm026.txtSTIstart.Value)
        Case "stiftelsesdatoTo"
            result = CStr(frm026.txtSTIslut.Value)
        Case "periodeStart"
            result = CStr(frm026.PeriodeStartdato.Value)
        Case "periodeStartFrom"
            result = CStr(frm026.txtPSTstart.Value)
        Case "periodeStartTo"
            result = CStr(frm026.txtPSTslut.Value)
        Case "periodeSlut"
            result = CStr(frm026.PeriodeSlutdato.Value)
        Case "periodeSlutFrom"
            result = CStr(frm026.txtPSLstart.Value)
        Case "periodeSlutTo"
            result = CStr(frm026.txtPSLslut.Value)
    End Select
                        
End Function
Private Function CheckNoExtraPrints()
    'Check Which configuration to choose
    Select Case parameters("testParameter")
        'Test different cases were different cells should be changed
        
        Case "noChangeWhenError"
        Case "noChangeWhenBackButton"
            popCells.Add "B6", ""
            popCells.Add "B7", ""
            popCells.Add "B8", ""
            popCells.Add "B9", ""
            popCells.Add "B10", ""
            popCells.Add "B11", ""
            popCells.Add "B12", ""
            popCells.Add "B13", ""
            popCells.Add "B14", ""
            popCells.Add "B15", ""
            
            
        Case "config1"
            popCells.Add "B6", parameters("forfaldsdatoFrom")
            popCells.Add "B7", parameters("forfaldsdatoTo")
            popCells.Add "B8", ""
            popCells.Add "B9", ""
            popCells.Add "B10", ""
            popCells.Add "B11", ""
            popCells.Add "B12", ""
            popCells.Add "B13", ""
            popCells.Add "B14", ""
            popCells.Add "B15", ""
            
            Call addSpm("4.a.2.1", "")
            Call addSpm("4.a.2.1_1", parameters("forfaldsdatoFrom"), parameters("forfaldsdatoTo"))
            
        Case "config2"
            popCells.Add "B6", ""
            popCells.Add "B7", ""
            popCells.Add "B8", parameters("srbFrom")
            popCells.Add "B9", parameters("srbTo")
            popCells.Add "B10", ""
            popCells.Add "B11", ""
            popCells.Add "B12", ""
            popCells.Add "B13", ""
            popCells.Add "B14", ""
            popCells.Add "B15", ""
            
            Call addSpm("4.a.2.1", "")
            Call addSpm("4.a.2.1_2", parameters("srbFrom"), parameters("srbTo"))
        Case "config3"
        
            popCells.Add "B6", ""
            popCells.Add "B7", ""
            popCells.Add "B8", ""
            popCells.Add "B9", ""
            popCells.Add "B10", parameters("stiftelsesdatoFrom")
            popCells.Add "B11", parameters("stiftelsesdatoTo")
            popCells.Add "B12", ""
            popCells.Add "B13", ""
            popCells.Add "B14", ""
            popCells.Add "B15", ""
            
            Call addSpm("4.a.2.1", "")
            Call addSpm("4.a.2.1_3", parameters("stiftelsesdatoFrom"), parameters("stiftelsesdatoTo"))
        Case "config4"
        
            popCells.Add "B6", ""
            popCells.Add "B7", ""
            popCells.Add "B8", ""
            popCells.Add "B9", ""
            popCells.Add "B10", ""
            popCells.Add "B11", ""
            popCells.Add "B12", parameters("periodeStartFrom")
            popCells.Add "B13", parameters("periodeStartTo")
            popCells.Add "B14", ""
            popCells.Add "B15", ""
            
            Call addSpm("4.a.2.1", "")
            Call addSpm("4.a.2.1_4", parameters("periodeStartFrom"), parameters("periodeStartTo"))
        Case "config5"
        
            popCells.Add "B6", ""
            popCells.Add "B7", ""
            popCells.Add "B8", ""
            popCells.Add "B9", ""
            popCells.Add "B10", ""
            popCells.Add "B11", ""
            popCells.Add "B12", ""
            popCells.Add "B13", ""
            popCells.Add "B14", parameters("periodeSlutFrom")
            popCells.Add "B15", parameters("periodeSlutTo")
            
            Call addSpm("4.a.2.1", "")
            Call addSpm("4.a.2.1_5", parameters("periodeSlutFrom"), parameters("periodeSlutTo"))
        
        
            
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
'            popCells = Array("B6", "B7", "B8", "B9", "B10", "B11", "B12", "B13", "B14", "B15")
'            rulCells = Array()
'            groCells = Array()
'            spmCells = Array("C7", "D8", "D9", "D10", "D11", "D12", "E8", "E9", "E10", "E11", "E12", "F8", "F9", "F10", "F11", "F12")
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
    If IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
    If IsLoaded("frm026") Then
        Unload frm026
    End If
    If IsLoaded("frm003") Then
        Unload frm003
    End If
    If IsLoaded("frm005") Then
        Unload frm005
    End If
End Function


    





