Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmbSchließen_Click()
On Error GoTo Err_cmbSchließen_Click

    FormKoorSave Me
    If Me.Dirty Then Me.Dirty = False
    DoCmd.Close

Exit_cmbSchließen_Click:
    Exit Sub

Err_cmbSchließen_Click:
    MsgBox Err.Description
    Resume Exit_cmbSchließen_Click

End Sub

Private Sub Form_DblClick(Cancel As Integer)

    Dim lsForm As String

    p_ReklNr = Me.ID
    NewReklamationE "Bearbeiten", get_p_ReklNr()
    lsForm = "Fehlererfassung - Register"

End Sub


Private Sub Form_Load()

    Dim i As Integer

    i = DCount("*", "qryTerminStellungnahmeKundeLieferant")

    If i = 0 Then
        MsgBox getTC_FCFromCode("TC4240"), vbInformation, getTC_FCFromCode("TC4208")
        DoCmd.Close
    Else
        FormOpenKoor Me
        Call FormEinUndAusBlenden(Me)
    End If

End Sub
