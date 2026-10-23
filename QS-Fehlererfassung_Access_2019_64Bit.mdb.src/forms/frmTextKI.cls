Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmbNeuerDatensatz_Click()

    Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")

    ldb.Execute ("Insert into tblKI ([Anfrage],[KZ]) values ('### Definition eintragen ###','" & get_pKuerzel() & "');")
    Me.Requery

End Sub

Private Sub cmbNeuerDatensatz_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Form_Close()

    FormKoorSave Me
    If IsLoaded("frmKI") Then
        DoCmd.Close acForm, "frmKI"
    End If
    DoCmd.OpenForm ("frmKI")

End Sub

Private Sub Form_Load()

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub
