Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************
'Änderungen:
'-----------
'16.05.2011:    Beim Drucken des Reklamationsbericht wird jetzt vorher das Formular "Fehlererfassung" gespeichert,
'               damit alle Daten im Druck erfasst werden.

Option Compare Database


Private Sub Drucken_Click()
On Error GoTo Err_Befehl18_Click

    ilReklID = Me.ID    'Save the Rekl.-Id

    Me.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.PrintOut acSelection

    DoCmd.SetWarnings False
    DoCmd.RunSQL "Update Mastertab Set Mastertab.[QS-Lieferant] = True Where Mastertab.[ID]=" & ilReklID & ";"
    DoCmd.SetWarnings True

    'Me.[QS-Lieferant] = True
    DoCmd.Close acForm, "Fxx_Reklamationsbericht_03_Sept08"
    If CurrentProject.AllForms("Fehlererfassung - Register").IsLoaded Then
        Forms("Fehlererfassung - Register").[Lieferanten-Code].SetFocus
    Else
        DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[ID]=" & ilReklID
    End If

'DoCmd.RunCommand acCmdPrint


Exit_Befehl18_Click:
    Exit Sub

Err_Befehl18_Click:
    MsgBox Err.Description
    Resume Exit_Befehl18_Click

End Sub

Private Sub E_Mail_Click()
On Error GoTo Err_E_Mail_Click

    Dim stDocName As String

    stDocName = "Makro8"
    DoCmd.RunMacro stDocName

Exit_E_Mail_Click:
    Exit Sub

Err_E_Mail_Click:
    MsgBox Err.Description
    Resume Exit_E_Mail_Click

End Sub
Private Sub Befehl172_Click()
On Error GoTo Err_Befehl172_Click

    ilReklID = Me.ID    'Save the Rekl.-Id

    DoCmd.Close acForm, "Fxx_Reklamationsbericht_03_Sept08"
    If CurrentProject.AllForms("Fehlererfassung - Register").IsLoaded Then
        Forms("Fehlererfassung - Register").[Lieferanten-Code].SetFocus
    Else
        DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[ID]=" & ilReklID
    End If

Exit_Befehl172_Click:
    Exit Sub

Err_Befehl172_Click:
    MsgBox Err.Description
    Resume Exit_Befehl172_Click

End Sub
