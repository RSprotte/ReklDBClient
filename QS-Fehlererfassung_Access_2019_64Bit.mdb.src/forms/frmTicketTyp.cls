Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cboFilterTicketTyp_AfterUpdate()

    Me.Filter = "[TicketTypID] = " & Me.cboFilterTicketTyp & ""
    Me.FilterOn = True
    Me.Requery
    Me.Refresh

End Sub

Private Sub Form_Load()

    Dim objMouseOverForm As clsMouseoverForm
    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With

'    Dim varKoord As Variant
'    'Form in richter Größe und Position laden
'    varKoord = getFktKoordinatenRelHF(Me)
'    If varKoord(3) > 0 And varKoord(2) > 0 Then
'        'Fenster setzen relativ zum Hauptfenster TB:autostart
'        DoCmd.MoveSize varKoord(0), varKoord(1), varKoord(2), varKoord(3)
'    End If
'    DoCmd.Restore

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

    Me.cboFilterTicketTyp = 1
    Me.Filter = "[TicketTypID] = " & Me.cboFilterTicketTyp & ""
    Me.FilterOn = True
    Me.Requery
    Me.Refresh

End Sub
Private Sub cmbNeuerTicketTyp_Click()
On Error GoTo Err_cmbNeuerTicketTyp_Click


    DoCmd.GoToRecord , , acNewRec

Exit_cmbNeuerTicketTyp_Click:
    Exit Sub

Err_cmbNeuerTicketTyp_Click:
    MsgBox Err.Description
    Resume Exit_cmbNeuerTicketTyp_Click

End Sub
Private Sub cmbSpeichernTicketTyp_Click()
On Error GoTo Err_cmbSpeichernTicketTyp_Click


    DoCmd.RunCommand acCmdSaveRecord
    Me.Requery
    Me.Refresh

Exit_cmbSpeichernTicketTyp_Click:
    Exit Sub

Err_cmbSpeichernTicketTyp_Click:
    MsgBox Err.Description
    Resume Exit_cmbSpeichernTicketTyp_Click

End Sub
Private Sub cmbFormTickettypSchließen_Click()
On Error GoTo Err_cmbFormTickettypSchließen_Click

    FormKoorSave Me
    Me.cboFilterTicketTyp = 1
    If Me.Dirty Then Me.Dirty = False
    DoCmd.Close

Exit_cmbFormTickettypSchließen_Click:
    Exit Sub

Err_cmbFormTickettypSchließen_Click:
    MsgBox Err.Description
    Resume Exit_cmbFormTickettypSchließen_Click

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub
