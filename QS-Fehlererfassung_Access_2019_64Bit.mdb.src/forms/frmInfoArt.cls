Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmbFilterLoeschen_Click()
    Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")

    Me.FilterOn = False
    Me.txtFilterFA = ""
    ldb.Execute ("delete * from tblTempInfoArt")
    Me.Requery

End Sub


Private Sub cmbFilterLoeschen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbPruefprotokoll_Click()

    Call cmbPruefprotokoll_Oeffnen(Me.FA)

End Sub

Private Sub cmbPruefprotokoll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Form_Close()
    Dim ldb As DAO.Database

    FormKoorSave Me

    Set ldb = openAccDBConnection("openDB")

    ldb.Execute ("delete * from tblTempInfoArt")
    Me.Requery

End Sub

Private Sub Form_DblClick(Cancel As Integer)

    Dim ls_sql As String

    If IsLoaded("frmStueliFA") Then DoCmd.Close acForm, "frmStueliFA"

    If Me.FA = "" Then
        Exit Sub
    Else
        ls_sql = cFaStueli
        If IsLoaded("Fehlererfassung - Register") Then
            ls_sql = Replace(ls_sql, "###AVISTA_KD_ID###", AVISTA_ID_KDnr(Forms("Fehlererfassung - Register").Controls("Mastertab#ID")))
        Else
            ls_sql = Replace(ls_sql, "###AVISTA_KD_ID###", "243") '243 ist LQ Kunden ID
        End If
        ls_sql = Replace(ls_sql, "###FA###", Me.FA)
        Call Stueli_Auslesen(ls_sql, Me.FA, "FA")
    End If

End Sub

Private Sub Form_Open(Cancel As Integer)

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub

Private Sub txtFilterFA_AfterUpdate()

    Dim l_Anzahl_DS As Integer

    If Me.txtFilterFA = "" Or IsNull(Me.txtFilterFA) Then
        Exit Sub
    Else
        Me.Filter = "[FA] = '" & Me.txtFilterFA & " '"
        Me.FilterOn = True
        l_Anzahl_DS = Me.RecordsetClone.RecordCount

        If l_Anzahl_DS = 0 Then
            InfoProduktion , , Me.txtFilterFA
        End If
    End If

End Sub
