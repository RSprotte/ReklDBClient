Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmbVorgangEnt_Click()
'Vorgang entsperren Zustand auf Frei setzten

    Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")

    ldb.Execute ("Update [Mastertab-DSBerechtigung] set [Zustand] = 'Frei' where ID = " & Me.ID & ";")

    Me.Requery

End Sub

Private Sub cmbVorgangEnt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbVorgangOeffnen_Click()
'Vorgang öffnen

    Call NewReklamationE("Bearbeiten", Me.ID)
    Forms("Fehlererfassung - Register").SetFocus

End Sub

Private Sub cmbVorgangOeffnen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Form_Close()

    FormKoorSave Me

End Sub

Private Sub Form_Open(Cancel As Integer)

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub

Private Sub Form_Timer()
'Aktuallisiert sich alle 5min, Zeitintervall steht auf 300000

    Me.Requery

End Sub
