Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()

    closeWinHandle Me

End Sub

Private Sub Form_DblClick(Cancel As Integer)

    Dim lsForm As String

    p_ReklNr = Me.ID

    NewReklamationE "Bearbeiten", get_p_ReklNr()
    lsForm = "Fehlererfassung - Register"

End Sub



Private Sub Form_Open(Cancel As Integer)


    Call FormEinUndAusBlenden(Me)
    FormOpenKoor Me

End Sub
