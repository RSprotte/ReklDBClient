Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Befehl59_Click()

    Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")
    ldb.Execute ("insert into tblKürzelZuTicketTyp ([SacharbeiterID], [AuswahlTicketTypID]) values (" & Me.ListeSacharbeiter & ", " & Me.ListeTicketTyp & ");")
    Me.Requery
    Me.Refresh

End Sub

Private Sub Befehl60_Click()

    Dim ldb As DAO.Database


    Set ldb = openAccDBConnection("openDB")

    ldb.Execute ("Delete from tblKürzelZuTicketTyp where SacharbeiterID =" & Me.ListeSacharbeiter & " and AuswahlTicketTypID=" & Me.ListeAuswahlTicketTyp & ";")
    Me.Requery
    Me.Refresh

End Sub


Private Sub ListeSacharbeiter_AfterUpdate()

    Me.Requery
    Me.Refresh

End Sub

Private Sub ListeSacharbeiter_Click()

    Me.Requery
    Me.Refresh

End Sub

Private Sub Objekt_schließen___Formular_Click()

    DoCmd.Close

End Sub

Private Sub txtKürzelSuchen_AfterUpdate()
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim KZ  As String

    KZ = (UCase(Me.txtKürzelSuchen))

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from [Sachbearbeiter] where Kurzzeichen = '" & KZ & "';")
        If rs.RecordCount Then
            Me.ListeSacharbeiter = rs("SB-ID")
        Else
'            MsgBox ("Kürzel nicht vorhanden")
            MsgBox getTC_FCFromCode("TC3197"), vbInformation, getTC_FCFromCode("TC3198")
        End If
    Set rs = Nothing

    Me.Requery
    Me.Refresh
End Sub
