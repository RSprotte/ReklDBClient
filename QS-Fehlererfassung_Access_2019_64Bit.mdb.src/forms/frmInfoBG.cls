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
    Me.txtFilterLI = ""
    Me.txtFilterBstNr = ""
    ldb.Execute ("delete * from tblTempInfoBG")
    Me.Requery

End Sub

Private Sub cmbFilterLoeschen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbInfoFa_Click()

    InfoProduktion , , Me.FA

End Sub

Private Sub cmbInfoFa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Form_Click()

    If IsLoaded("frmInfoArt") Then
        InfoProduktion , , Me.FA
        Forms("frmInfoArt").Requery
        Forms("frmInfoArt").SetFocus
    End If

End Sub

Private Sub Form_Close()
    Dim ldb As DAO.Database

    FormKoorSave Me

    Set ldb = openAccDBConnection("openDB")

    ldb.Execute ("delete * from tblTempInfoBG")
    Me.Requery

End Sub

Private Sub Form_DblClick(Cancel As Integer)

    Dim ls_sql As String

    If IsLoaded("frmStueliFA") Then DoCmd.Close acForm, "frmStueliFA"

    If Me.FA = "" Then
        Exit Sub
    Else
        ls_sql = cFaStueli
        ls_sql = Replace(ls_sql, "###AVISTA_KD_ID###", AVISTA_ID_KDnr(Forms("Fehlererfassung - Register").Controls("Mastertab#ID")))
        ls_sql = Replace(ls_sql, "###FA###", Me.FA)
        Call Stueli_Auslesen(ls_sql, Me.FA, "FA")
    End If

End Sub

Private Sub Form_Open(Cancel As Integer)

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub

Private Sub txtFilterBstNr_AfterUpdate()
'89588986
    Dim l_Anzahl_DS As Integer

    If Me.txtFilterBstNr = "" Or IsNull(Me.txtFilterBstNr) Then
        Exit Sub
    Else
        Me.Filter = "[Bestellnummer] = '" & Me.txtFilterBstNr & " '"
        Me.FilterOn = True
        l_Anzahl_DS = Me.RecordsetClone.RecordCount

        If l_Anzahl_DS = 0 Then
            InfosZuBG "BstNr", , , , Me.txtFilterBstNr
        End If
    End If

End Sub

Private Sub txtFilterFA_AfterUpdate()
'FA-230502-687651
    Dim l_Anzahl_DS As Integer

    If Me.txtFilterFA = "" Or IsNull(Me.txtFilterFA) Then
        Exit Sub
    Else
        Me.Filter = "[FA] = '" & Me.txtFilterFA & " '"
        Me.FilterOn = True
        l_Anzahl_DS = Me.RecordsetClone.RecordCount

        If l_Anzahl_DS = 0 Then
            InfosZuBG "FA", , Me.txtFilterFA
        End If
    End If

End Sub
Private Sub txtFilterLI_AfterUpdate()
'LI2307-14212
    Dim l_Anzahl_DS As Integer

    If Me.txtFilterLI = "" Or IsNull(Me.txtFilterLI) Then
        Exit Sub
    Else
        Me.Filter = "[Lieferschein] = '" & Me.txtFilterLI & " '"
        Me.FilterOn = True
        l_Anzahl_DS = Me.RecordsetClone.RecordCount

        If l_Anzahl_DS = 0 Then
            InfosZuBG "LI", , , Me.txtFilterLI
        End If
    End If
End Sub
