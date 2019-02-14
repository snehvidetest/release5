Attribute VB_Name = "Frm002_test"
'****Test script for frm002****

Private formID As Integer
Public parameters As Scripting.Dictionary
Private parametersAndCols As Scripting.Dictionary


Private popCells As Scripting.Dictionary
Private rulCells As Scripting.Dictionary
Private groCells As Scripting.Dictionary
Private result As String


Sub RunTests()
'On Error GoTo Error_handler
'****GUIDE****:

'The main testcase template for form 2 is called:
'frm002_testcase ( tc )
'
'Input: tc
'   An integer which identifies the test case number. This is not the tcid!
'
'The function read data from the testcases excel workbook, hereunder
'
'   - testSubject: This string defines what the testcase is testing. Possible values are:
'       - printsToPopSheet: Checks the form input is written correctly to the population sheet
'       - printsToSpmSheet: Checks the form input is written correctly to the SpmSvar sheet
'       - errorMessage: Check error message is correct
'       - tidligereBesvarelse: Checks that a form can correctly load a previous response to that form
'       - nextStep: Checks that next form(s) is(are) called correctly
'       - backButton: Checks the back button functions
'
'   - testParameter (where relevant): If the testcase relates to a certain parameter, this variable identifies it. Possible values:
'       - "fordringshaverID"
'       - "fordringType"
'       - "modtagelseStart"
'       - "modtagelseSlut"
'       - "forkertData"
'       - "korrektData"
'
'   - expected: The relevant value we expect to find
'
'   - The relevant test paramters. For this form they are:
'       - fordringshaverID (String): The fordringshaverID
'       - fordringType (String): The fordrings type
'       - modtagelseStart (String): The modtagelses start date
'       - modtagelseSlut (String): The modtagelses end date
'       - forkertData (Boolean): If we believe data is entered wrong by fordringhaver.
'       - korrektData (Boolean): If we believe data is entered korrekt by fordringhaver.
'


'Which form are we testing?
formID = 2

'Get parameters relevant for testcase including their respective columns
Set parametersAndCols = New Scripting.Dictionary
Set parametersAndCols = Global_Test_Func.getParamtersAndTheirCols(formID)

Set Main_Test_Form.spmCells = New Scripting.Dictionary   'Dictionaries to log changes in SpmSvar sheet
Set popCells = New Scripting.Dictionary   'Dictionaries to log changes in Population sheet
Set groCells = New Scripting.Dictionary   'Dictionaries to log changes in Gruppering sheet
Set rulCells = New Scripting.Dictionary   'Dictionaries to log changes in Regler sheet

'Get the total number of testcases associated with the form
Dim nrTC As Integer, i As Integer
nrTC = Application.WorksheetFunction.CountIf(testWS.Range("A:A"), formID)

'Run all testcases incl. printing of results to the testcase workbook

For i = 1 To nrTC
    Testcase i
Next i

    Exit Sub
Error_handler:
    Global_Test_Func.PrintTestResults CStr(formID) + "." + CStr(i), "crash", "False"
    Resume Next

End Sub


'The following code is the skeleton for form 2 testcases.
Private Function Testcase(tc As Integer)
    
    Dim review As Boolean, tcid As String
    
    'Reset spørgeskema workbook
    Global_Test_Func.resetSheets ThisWorkbook
    
    'Create the TCID
    tcid = Global_Test_Func.GetTCID(tc, formID)
    If logging Then
        Write #1, tcid
    End If
    
    'Get testcase data
    Set parameters = New Scripting.Dictionary 'Resets the testcase parameters
    Set parameters = Global_Test_Func.getData(tcid, parametersAndCols)
    
    'Check if testcase should be run
    If parameters("run") = 0 Then
        Exit Function
    End If
       
    'Get results
    Select Case parameters("testSubject")
        Case "printsToPopSheet"
        
            'Enter data into form
            SetFields
            
            'Reset spørgeskema workbook
            Global_Test_Func.resetSheets ThisWorkbook
    
            'Execute/Click button
            frm002.OKButton_Click
            
            
            Select Case parameters("testParameter")
            Case "fordringshaverID"
                result = ThisWorkbook.Sheets("Population").Range("B2").Text
            Case "fordringType"
                result = ThisWorkbook.Sheets("Population").Range("B3").Text
            Case "modtagelseStart"
                result = ThisWorkbook.Sheets("Population").Range("B4").Text
            Case "modtagelseSlut"
                result = ThisWorkbook.Sheets("Population").Range("B5").Text
            Case Else
                MsgBox "Error in 'testParameter' input: tcid " & tcid 'Msgbox to stop the code because you made a mistake in the inputs..
            End Select
        
        Case "printsToSpmSheet"
        
            'Enter data into form
            SetFields
            
            'Reset spørgeskema workbook
            Global_Test_Func.resetSheets ThisWorkbook
            
            'Execute/Click button
            frm002.OKButton_Click
            
            Select Case parameters("testParameter")
            Case "fordringshaverID"
                result = findPreviousAns(findTopSpm("A"), "1", 1, 1)
                'result = ThisWorkbook.Sheets("SpmSvar").Range("D2").Text
            Case "fordringType"
                result = findPreviousAns(findTopSpm("A"), "2", 1, 1)
                'result = ThisWorkbook.Sheets("SpmSvar").Range("D3").Text
            Case "modtagelseStart"
                result = findPreviousAns(findTopSpm("A"), "3", 1, 1)
                'result = ThisWorkbook.Sheets("SpmSvar").Range("D4").Text
            Case "modtagelseSlut"
                result = findPreviousAns(findTopSpm("A"), "3", 2, 1)
                'result = ThisWorkbook.Sheets("SpmSvar").Range("E4").Text
            Case "forkertData"
                result = findPreviousAns(findTopSpm("A"), "4", 1, 1)
                'result = ThisWorkbook.Sheets("SpmSvar").Range("D5").Text
            Case "korrektData"
                result = findPreviousAns(findTopSpm("A"), "4", 1, 1)
                'result = ThisWorkbook.Sheets("SpmSvar").Range("D5").Text
            Case Else
                MsgBox "Error in 'testParameter' input: tcid " & tcid
            End Select
            
            
        Case "errorMessage"
        
            'Enter data into form
            SetFields
            
            'Reset spørgeskema workbook
            Global_Test_Func.resetSheets ThisWorkbook
            
            'Execute/Click button
            frm002.OKButton_Click
            
            If parameters("testParameter") <> "before01092013" Then
                'Get the error message
                result = Global_Test_Func.errorMessage()
            Else
                If IsLoaded("frm043") Then
                    result = frm043.Label1.caption
                Else
                    result = "Message did not pop up"
                End If
            End If
        Case "nextStep"
        
            'Enter data into form
            SetFields
            
            'Reset spørgeskema workbook
            Global_Test_Func.resetSheets ThisWorkbook
            
            'Execute/Click button
            frm002.OKButton_Click
            
            'Check if the expected form opened
            If Global_Test_Func.IsLoaded(parameters("expected")) Then
                result = parameters("expected")
            Else
                result = "Incorrect"
            End If
            
        Case "backButton"
        
            'Enter data into form
            SetFields
            'frm002.Show 'Check it is correct
            
            'Reset spørgeskema workbook
            Global_Test_Func.resetSheets ThisWorkbook
            
            recHis ("frm001")
            
            'Execute/Click button
            frm002.Tilbage_Click
            
            'Check if the expected form opened
            If Global_Test_Func.IsLoaded(parameters("expected")) Then
                result = parameters("expected")
            Else
                result = "Incorrect"
            End If
            
        Case "tidligereBesvarelse"
            
            'Reset spørgeskema workbook
            Global_Test_Func.resetSheets ThisWorkbook
            
            'Pre-populate SpmSvar sheet
            prePopulateFields
            
            'Initialise form
            SFunc.ShowFunc ("frm002")
        
            Select Case parameters("testParameter")
                Case "fordringshaverID"
                    result = frm002.txtFordringsId.Text
                Case "fordringType"
                    result = frm002.cboFordringstype.Value
                Case "modtagelseStart"
                    result = frm002.txtModtStart.Value
                Case "modtagelseSlut"
                    result = frm002.txtModtSlut.Value
                Case "forkertData"
                    result = frm002.forkertData.Value
                Case "korrektData"
                    result = frm002.korrektData.Value
                Case Else
                    MsgBox "Error in 'testParameter' input: tcid " & tcid 'Msgbox to stop the code because you made a mistake in the inputs..
            End Select
            
        Case "noExtraPrints"
            
            'Reset spørgeskema workbook
            Global_Test_Func.resetSheets ThisWorkbook
            
            SetFields
            Sheet1.recordChangingCells = True
            
            
            
            If parameters("testParameter") = "noChangeWhenBackButton" Then
                frm002.Tilbage_Click 'Click back button
            Else
                frm002.OKButton_Click 'Click on Videre button
            End If
            CheckNoExtraPrints
            Sheet1.recordChangingCells = False
                    
        Case Else
            MsgBox "Error in 'testSubject' input: tcid " & tcid 'Msgbox to stop the code because you made a mistake in the inputs..
    
    End Select


    'Compare actual and expected
    If result = parameters("expected") Then
        review = True
    Else:
        review = False
    End If
    
    KillForms

    'Print results
    Global_Test_Func.PrintTestResults tcid, result, review

End Function
Private Function CheckNoExtraPrints()
    'Check Which configuration to choose
    Select Case parameters("testParameter")
        'Test different cases were different cells should be changed
        Case "noChangeWhenError"
        Case "noChangeWhenBackButton"
            popCells.Add "B2", ""
            popCells.Add "B3", ""
            popCells.Add "B4", ""
            popCells.Add "B5", ""
            
            
            Call addSpm("1", "")
            Call addSpm("2", "")
            Call addSpm("3", "")
            Call addSpm("4", "", "")
            Call addSpm("5", "")
            
'            Main_Test_Form.spmCells.Add "C2", ""
'            Main_Test_Form.spmCells.Add "C3", ""
'            Main_Test_Form.spmCells.Add "C4", ""
'            Main_Test_Form.spmCells.Add "C5", ""
'            Main_Test_Form.spmCells.Add "D4", ""
        
        Case "config1"
            popCells.Add "B2", parameters("fordringshaverID")
            popCells.Add "B3", parameters("fordringType")
            popCells.Add "B4", parameters("modtagelseStart")
            popCells.Add "B5", parameters("modtagelseSlut")
            
            Call addSpm("1", parameters("fordringshaverID"))
            Call addSpm("2", parameters("fordringType"))
            Call addSpm("3", parameters("modtagelseStart"), parameters("modtagelseSlut"))
            Call addSpm("4", "Ja")
            
        Case "config2"
            popCells.Add "B2", parameters("fordringshaverID")
            popCells.Add "B3", parameters("fordringType")
            popCells.Add "B4", parameters("modtagelseStart")
            popCells.Add "B5", parameters("modtagelseSlut")
            
            Call addSpm("1", parameters("fordringshaverID"))
            Call addSpm("2", parameters("fordringType"))
            Call addSpm("3", parameters("modtagelseStart"), parameters("modtagelseSlut"))
            Call addSpm("4", "Nej")
'            spmCells.Add "C2", parameters("fordringshaverID")
'            spmCells.Add "C3", parameters("fordringType")
'            spmCells.Add "C4", parameters("modtagelseStart")
'            spmCells.Add "C5", "Nej"
'            spmCells.Add "D4", parameters("modtagelseSlut")
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
    Main_Test_Form.spmCells.RemoveAll
End Function

Private Function SetFields()
    'The folowing code inserts the inputs into the actual form
    
    ThisWorkbook.Activate
    
    'Set values in form
    frm002.txtFordringsId.SetFocus
    frm002.txtFordringsId.Value = parameters("fordringshaverID")
    frm002.cboFordringstype.SetFocus
    frm002.cboFordringstype.Value = parameters("fordringType")
    If (parameters("modtagelseStart") <> "") Then
        frm002.txtModtStart.Value = parameters("modtagelseStart")
    End If
    frm002.txtModtSlut.Value = parameters("modtagelseSlut")
    frm002.forkertData.Value = parameters("forkertData")
    frm002.korrektData.Value = parameters("korrektData")
End Function

Private Function prePopulateFields()
    'The folowing code inserts the inputs spmSvar sheet
    
    Dim ws As Worksheet
    
    'Clear relevant fields
    ThisWorkbook.Activate
    Set ws = ThisWorkbook.Sheets("SpmSvar")
    
    'Set values in SpmSvar sheet
    Call writeSpmSvar("1", "", parameters("fordringshaverID"), "", 6)
    Call writeSpmSvar("2", "", parameters("fordringType"), "", 6)
    Call writeSpmSvar("3", "", parameters("modtagelseStart"), parameters("modtagelseSlut"), 6)
    'ws.Range("D2").Value = "'" & parameters("fordringshaverID")
    'ws.Range("D3").Value = parameters("fordringType")
    'ws.Range("D4").Value = "'" & parameters("modtagelseStart")
    'ws.Range("E4").Value = "'" & parameters("modtagelseSlut")
    If parameters("forkertData") = True Then
    '    ws.Range("D5").Value = "Ja"
        Call writeSpmSvar(4, "", "Ja", "", 6)
    ElseIf parameters("korrektData") = True Then
    '    ws.Range("D5").Value = "Nej"
        Call writeSpmSvar(4, "", "Nej", "", 6)
    End If
    
End Function



Private Function KillForms()
    'Closes forms
    ThisWorkbook.Activate
    If Global_Test_Func.IsLoaded("frm001") Then
        Unload frm001
    End If
    If Global_Test_Func.IsLoaded("frm002") Then
        Unload frm002
    End If
    If Global_Test_Func.IsLoaded("frm003") Then
        Unload frm003
    End If
    If Global_Test_Func.IsLoaded("frm005") Then
        Unload frm005
    End If
    If Global_Test_Func.IsLoaded("frm043") Then
        Unload frm043
    End If
    If Global_Test_Func.IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
End Function




