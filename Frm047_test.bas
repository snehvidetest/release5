Attribute VB_Name = "Frm047_test"
Private result As String
Private formID As Integer
Private formName As String
Private stopFormTest As Boolean
Private parameters As Scripting.Dictionary
Private parametersAndCols As Scripting.Dictionary
Private spmCells As Scripting.Dictionary
Private popCells As Scripting.Dictionary
Private rulCells As Scripting.Dictionary
Private groCells As Scripting.Dictionary


Sub RunTests()
    On Error GoTo Error_handler
    formName = "frm047"
    formID = 47
    
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
        Case "nextStep"
            frm047.CommandButton2_Click 'Click on Videre button
            result = Global_Test_Func.NextStep(parameters("expected"))
            
        Case "backButton"
            frm047.CommandButton1_Click
            result = Global_Test_Func.IsLoaded(formName)
            
        Case "noExtraPrints"
            Sheet1.recordChangingCells = True
            If (parameters("testParameter") = "buttonOne") Then
                frm047.CommandButton1_Click
            ElseIf (parameters("testParameter") = "buttonTwo") Then
                frm047.CommandButton2_Click
            End If
            CheckNoExtraPrints
            Sheet1.recordChangingCells = False
            
        Case "checkCaption"
            If (parameters("testParameter") = "buttonOne") Then
                result = frm047.CommandButton1.caption
            ElseIf (parameters("testParameter") = "buttonTwo") Then
                result = frm047.CommandButton2.caption
            ElseIf (parameters("testParameter") = "beskrivelse") Then
                result = frm047.Label1.caption
            End If
            
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

Private Function CheckNoExtraPrints()
    'Check Which configuration to choose
    Select Case parameters("testParameter")
        'Test different cases were different cells should be changed
        Case "buttonOne"
        Case "buttonTwo"
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
    If Global_Test_Func.IsLoaded("frm047") Then
        Unload frm047
    End If
    If Global_Test_Func.IsLoaded("frm021") Then
        Unload frm021
    End If
    If Global_Test_Func.IsLoaded("frmMsg") Then
        Unload frmMsg
    End If
End Function


