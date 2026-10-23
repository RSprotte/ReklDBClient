Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Befehl79_Click()

    Call txtWirkungGeprüftAm_AfterUpdate

End Sub

Private Sub ButZurdnungTaetigkeiten_Click()

DoCmd.RunCommand acCmdSaveRecord
DoCmd.OpenForm ("MaßnahmeZuordnungFehlerklasse_Dialog")

End Sub

Private Sub cmbAktivierungMWPr_Click()          '20180125-ROS

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim rs2 As DAO.Recordset
    Dim rs_Anhaenge As DAO.Recordset
    Dim MailBCC As String
    Dim MailVon As String
    Dim MailCC As String
    Dim MailTo As String
    Dim lsaColAnhang As Collection


    If IsNull(Me.cboVerantPrüfSB) Or IsNull(Me.txtWirkungPrüfenBis) Then
'        MsgBox ("Daten sind nicht Vollständig")
        MsgBox getTC_FCFromCode("TC3203"), vbInformation, getTC_FCFromCode("TC3204")
        Exit Sub
    Else
        pl_KuerzelMWPR = Me.cboVerantPrüfSB
    End If
    Me.chkAktivierungMWPr = True
    Me.chkAktivierungMWPr.Enabled = False

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset(" Select * from qryTicketMaßnahmenPE where [SB-ID] = " & pl_KuerzelMWPR & ";")

    If IsNull(rs("MailBCC")) Then
        MailBCC = ""
    Else
        MailBCC = rs("MailBCC")
    End If

    MailVon = rs("MailVon")
    MailCC = rs("MailCC")
    MailTo = rs("MailTo")

    If Me.MaßnahmenHerkunftID <> 5 Then
        MailVon = DLookup("MailTOLQEinheit", "LQEinheit", "LQEinheit = 'QM'")
        MailCC = DLookup("MailTOLQEinheit", "LQEinheit", "LQEinheit = 'QM'")
        MailBCC = DLookup("MailCCLQEinheit", "LQEinheit", "LQEinheit = 'QM'")
    End If

    Call TicketMWAnlegen(Me.MaßnahmeID, 12, Me.txtWirkungPrüfenBis, Me.txtWirkungPrüfenBis, Me.txtWirkungPrüfenBis, 2, Me.txtWirkungPrüfenBis, MailVon, MailTo, MailCC, MailBCC)

    Set rs2 = ldb.OpenRecordset(" Select * from tblTicketListe where [ReklID] = " & Me.MaßnahmeID & " and [TicketTypID] = 12 ;")

    Set lsaColAnhang = New Collection
    Set rs_Anhaenge = ldb.OpenRecordset("select * From tblTicketAnhaenge where [ReklID] = " & Me.MaßnahmeID & ";")
    While Not rs_Anhaenge.EOF
        'Container der Anhänge füllen
        sPfad = rs_Anhaenge("AnhangPfad")
        n = InStrRev(rs_Anhaenge("AnhangPfad"), "\")
        sFile = Right$(rs_Anhaenge("AnhangPfad"), Len(rs_Anhaenge("AnhangPfad")) - n)
        lsaColAnhang.Add sPfad, sFile
        rs_Anhaenge.MoveNext
    Wend

    sendMailAufgabeAnhang rs2("MailVon"), rs2("MailAn"), rs2("MailBetreff"), rs2("MailText"), lsaColAnhang, True, rs2("MailCC"), rs2("MailBCC")

    Set rs = Nothing
    Set rs2 = Nothing
    Set rs_Anhaenge = Nothing


End Sub

Private Sub cmbMahnwesenAktivieren_Click()          '20180125-ROS

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim rs2 As DAO.Recordset
    Dim rs_Anhaenge As DAO.Recordset
    Dim MailBCC As String
    Dim MailVon As String
    Dim MailCC As String
    Dim MailTo As String
    Dim lsaColAnhang As Collection


    If IsNull(Me.Kombinationsfeld32) Or IsNull(Me.Kombinationsfeld20) Or IsNull(Me.StartUmsetzung) Or IsNull(Me.EndeUmsetzung) Or IsNull(Me.cboPriorität) Or IsNull(Me.cboMaßnahmenHerkunft) Then
'        MsgBox ("Daten sind nicht Vollständig")
        MsgBox getTC_FCFromCode("TC3203"), vbInformation, getTC_FCFromCode("TC3206")
        Exit Sub
    Else
        pl_KuerzelMWSB = Me.Kombinationsfeld32
    End If
    Me.chkAktivierungMM = True
    Me.chkAktivierungMM.Enabled = False

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset(" Select * from qryTicketMaßnahmenPE where [SB-ID] = " & pl_KuerzelMWSB & ";")

    If IsNull(rs("MailBCC")) Then
        MailBCC = ""
    Else
        MailBCC = rs("MailBCC")
    End If

    MailVon = rs("MailVon")
    MailCC = rs("MailCC")
    MailTo = rs("MailTo")

    If Me.MaßnahmenHerkunftID <> 5 Then
        MailVon = DLookup("MailTOLQEinheit", "LQEinheit", "LQEinheit = 'QM'")
        MailCC = DLookup("MailTOLQEinheit", "LQEinheit", "LQEinheit = 'QM'")
        MailBCC = DLookup("MailCCLQEinheit", "LQEinheit", "LQEinheit = 'QM'")
    End If


    Call TicketMWAnlegen(Me.MaßnahmeID, 11, Me.StartUmsetzung, Me.EndeUmsetzung, Me.EndeUmsetzung, 2, Me.StartUmsetzung, MailVon, MailTo, MailCC, MailBCC)

    Set rs2 = ldb.OpenRecordset(" Select * from tblTicketListe where [ReklID] = " & Me.MaßnahmeID & " and [TicketTypID] = 11 ;")

    Set lsaColAnhang = New Collection
    Set rs_Anhaenge = ldb.OpenRecordset("select * From tblTicketAnhaenge where [ReklID] = " & Me.MaßnahmeID & ";")
    While Not rs_Anhaenge.EOF
        'Container der Anhänge füllen
        sPfad = rs_Anhaenge("AnhangPfad")
        n = InStrRev(rs_Anhaenge("AnhangPfad"), "\")
        sFile = Right$(rs_Anhaenge("AnhangPfad"), Len(rs_Anhaenge("AnhangPfad")) - n)
        lsaColAnhang.Add sPfad, sFile
        rs_Anhaenge.MoveNext
    Wend

    sendMailAufgabeAnhang rs2("MailVon"), rs2("MailAn"), rs2("MailBetreff"), rs2("MailText") & "<br>" & getTC_FCFromCode("TC3214") & ": " & rs("Kurzzeichen"), lsaColAnhang, True, rs2("MailCC"), rs2("MailBCC")

    Set rs = Nothing
    Set rs2 = Nothing
    Set rs_Anhaenge = Nothing

End Sub

Private Sub EndeUmsetzung_AfterUpdate()

    Dim ldb As DAO.Database
    Dim rs As Recordset
    Dim lsRetFunc As String

        If IsNull(Me.txtUmgesetztAm) Or Me.txtUmgesetztAm = "" Then
            lsRetFunc = DatumAenderungMW(Me.MaßnahmeID, 11, Me.StartUmsetzung, Me.EndeUmsetzung)
        Else
            lsRetFunc = DatumAenderungMW(Me.MaßnahmeID, 11, Me.StartUmsetzung, Me.EndeUmsetzung, Me.txtUmgesetztAm)
        End If
        If lsRetFunc < "31.12.9999" Then Me.EndeUmsetzung = lsRetFunc


    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset(" Select * from tblTicketListe where ([ReklID] = " & Me.MaßnahmeID & ") and ([TicketTypID] = 11);", dbOpenSnapshot)

        DoCmd.RunCommand acCmdSaveRecord

        If Me.EndeUmsetzung <= gVerlängertBis Then
            Exit Sub
        Else
            If rs.RecordCount Then fct_changedateinTicketlist (rs("TicketID"))     '20190415-ROS ausführen nur wenn Ticket vorhanden ist
        End If

        Me.Requery
        Me.Refresh

        Forms("Maßnahmen").Requery
        Forms("Maßnahmen").Refresh

        Me.SetFocus

    Set rs = Nothing

End Sub

Private Sub EndeUmsetzung_GotFocus()

    gVerlängertBis = Me.EndeUmsetzung

End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)

    Me.StartUmsetzung = Date
    Me.EndeUmsetzung = Date
    Me.txtWirkungPrüfenBis = Date


End Sub

Private Sub Form_Load()

    Dim objMouseOverForm As clsMouseoverForm
    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With

'    Dim varKoord As Variant
'    varKoord = getFktKoordinatenRelHF(Me)
'    If varKoord(3) > 0 And varKoord(2) > 0 Then
'        'Fenster setzen relativ zum Hauptfenster TB:autostart
'        DoCmd.MoveSize varKoord(0), varKoord(1), varKoord(2), varKoord(3)
'    End If
'    DoCmd.Restore

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub


Private Sub ListeErMID_DblClick(Cancel As Integer)
On Error GoTo Err_DatensatzBearbeiten_Click

    Dim ErweiterungVonID As Long

    ErweiterungVonID = Me.ListeErMID

    DoCmd.Close acForm, "MaßnahmenDialog", acSaveYes
    DoCmd.OpenForm "MaßnahmenDialog", acNormal, , , acFormAddMID
    Forms("MaßnahmenDialog").Filter = "[Maßnahmen].[MaßnahmeID] = " & ErweiterungVonID
    Forms("MaßnahmenDialog").FilterOn = True
    Forms("MaßnahmenDialog").Requery
    Forms("MaßnahmenDialog").Refresh
    Forms("MaßnahmenDialog").AllowAdditions = False





Exit_DatensatzBearbeiten_Click:
    Exit Sub

Err_DatensatzBearbeiten_Click:
    MsgBox Err.Description
    Resume Exit_DatensatzBearbeiten_Click

End Sub

Private Sub ListeErweiterung_DblClick(Cancel As Integer)
On Error GoTo Err_DatensatzBearbeiten_Click

    Dim ErweiterungID As Long

    ErweiterungID = Me.ListeErweiterung

    DoCmd.Close acForm, "MaßnahmenDialog", acSaveYes
    DoCmd.OpenForm "MaßnahmenDialog", acNormal, , , acFormAddMID
    Forms("MaßnahmenDialog").Filter = "[Maßnahmen].[MaßnahmeID] = " & ErweiterungID
    Forms("MaßnahmenDialog").FilterOn = True
    Forms("MaßnahmenDialog").Requery
    Forms("MaßnahmenDialog").Refresh
    Forms("MaßnahmenDialog").AllowAdditions = False





Exit_DatensatzBearbeiten_Click:
    Exit Sub

Err_DatensatzBearbeiten_Click:
    MsgBox Err.Description
    Resume Exit_DatensatzBearbeiten_Click
End Sub

Private Sub ListeReklamation_DblClick(Cancel As Integer)

    p_ReklNr = Me.ListeReklamation

'    DoCmd.Close acForm, "MaßnahmenDialog", acSaveYes
'    DoCmd.Close acForm, "Maßnahmen", acSaveYes
    Call NewReklamationE("Bearbeiten", get_p_ReklNr())

End Sub

Private Sub ObjektSchließen_Click()

    FormKoorSave Me
    DoCmd.Save acForm, "MaßnahmenDialog"
    DoCmd.Close

End Sub

Private Sub StartUmsetzung_AfterUpdate()

    Dim lsRetFunc As String

    If IsNull(Me.txtUmgesetztAm) Or Me.txtUmgesetztAm = "" Then
        lsRetFunc = DatumAenderungMW(Me.MaßnahmeID, 11, Me.StartUmsetzung, Me.EndeUmsetzung)
    Else
        lsRetFunc = DatumAenderungMW(Me.MaßnahmeID, 11, Me.StartUmsetzung, Me.EndeUmsetzung, Me.txtUmgesetztAm)
    End If
    If lsRetFunc < "31.12.9999" Then Me.StartUmsetzung = lsRetFunc

End Sub

Private Sub txtUmgesetztAm_AfterUpdate()
On Error Resume Next
    Dim lsRetFunc As String
    Dim lKuerzel As String

    lKuerzel = DLookup("[Kurzzeichen]", "Sachbearbeiter", "[SB-ID] = " & Me.Kombinationsfeld32 & " ")

    If IsNull(Me.txtUmgesetztAm) Or Me.txtUmgesetztAm = "" Then
        lsRetFunc = DatumAenderungMW(Me.MaßnahmeID, 11, Me.StartUmsetzung, Me.EndeUmsetzung)
    Else
        lsRetFunc = DatumAenderungMW(Me.MaßnahmeID, 11, Me.StartUmsetzung, Me.EndeUmsetzung, Me.txtUmgesetztAm, lKuerzel)
    End If
    If lsRetFunc < "31.12.9999" Then Me.txtUmgesetztAm = lsRetFunc

    Me.Requery

End Sub

Private Sub txtUmgesetztAm_LostFocus()

'    Call txtUmgesetztAm_AfterUpdate

End Sub

Private Sub txtWirkungGeprüftAm_AfterUpdate()
On Error Resume Next
    Dim lsRetFunc As String
    Dim lKuerzel As String

    lKuerzel = DLookup("[Kurzzeichen]", "Sachbearbeiter", "[SB-ID] = " & Me.cboVerantPrüfSB & " ")

    If IsNull(Me.txtWirkungGeprüftAm) Or Me.txtWirkungGeprüftAm = "" Then
        lsRetFunc = DatumAenderungMW(Me.MaßnahmeID, 12, Me.txtWirkungPrüfenBis, Me.txtWirkungPrüfenBis)
    Else
        lsRetFunc = DatumAenderungMW(Me.MaßnahmeID, 12, Me.txtWirkungPrüfenBis, Me.txtWirkungPrüfenBis, Me.txtWirkungGeprüftAm, lKuerzel)
    End If
    If lsRetFunc < "31.12.9999" Then Me.txtWirkungGeprüftAm = lsRetFunc

    Me.Requery

End Sub

Private Sub txtWirkungGeprüftAm_LostFocus()

'    Call txtWirkungGeprüftAm_AfterUpdate

End Sub

Private Sub txtWirkungPrüfenBis_AfterUpdate()
On Error Resume Next

    Dim ldb As DAO.Database
    Dim rs As Recordset
    Dim lsRetFunc As String

    If IsNull(Me.txtWirkungGeprüftAm) Or Me.txtWirkungGeprüftAm = "" Then
        lsRetFunc = DatumAenderungMW(Me.MaßnahmeID, 12, Me.txtWirkungPrüfenBis, Me.txtWirkungPrüfenBis)
    Else
        lsRetFunc = DatumAenderungMW(Me.MaßnahmeID, 12, Me.txtWirkungPrüfenBis, Me.txtWirkungPrüfenBis, Me.txtWirkungPrüfenBis)
    End If
    If lsRetFunc < "31.12.9999" Then Me.txtWirkungPrüfenBis = lsRetFunc

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset(" Select * from tblTicketListe where ([ReklID] = " & Me.MaßnahmeID & ") and ([TicketTypID] = 12);", dbOpenSnapshot)

        DoCmd.RunCommand acCmdSaveRecord

        If Me.WirkungPrüfenBis <= gVerlängertBis Then
            Exit Sub
        Else
            fct_changedateinTicketlist (rs("TicketID"))
        End If

        Me.Requery
        Me.Refresh

        Forms("Maßnahmen").Requery
        Forms("Maßnahmen").Refresh

        Me.SetFocus
    Set rs = Nothing

End Sub

Private Sub txtWirkungPrüfenBis_GotFocus()

'     gVerlängertBis = Me.WirkungPrüfenBis

End Sub

Private Sub txtWirkungPrüfenBis_LostFocus()

    gVerlängertBis = Me.WirkungPrüfenBis

End Sub

Private Sub ZuordBearbOrt_Click()
On Error GoTo err_ZuordBearbOrt_Click
DoCmd.RunCommand acCmdSaveRecord
DoCmd.OpenForm ("MaßnahmeZuordnungBearbeitungsort_Dialog")
'With Forms("MaßnahmeZuordnungBearbeitungsort_Dialog")
'    .Filter = "MaßnahmeID = " & Forms("MaßnahmenDialog").Controls("MaßnahmeID").Value
'    .FilterOn = True
'    .AllowAdditions = False
'    .Requery
'    .Refresh
'End With
Exit Sub
err_ZuordBearbOrt_Click:
    Resume Next

End Sub

Private Sub ZurdnungFehlerklasse_Click()

DoCmd.RunCommand acCmdSaveRecord
DoCmd.OpenForm ("MaßnahmeZuordnungFehlerklasse_Dialog")
'Forms("MaßnahmeZuordnungFehlerklasse_Dialog").Controls("MID") = Me.MaßnahmeID
'Forms("MaßnahmeZuordnungFehlerklasse_Dialog").Controls("IDBeschreibung") = Me.Beschreibung


End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub

Private Sub CmdMaßnahmeErweitern_Click()
On Error GoTo Err_CmdMaßnahmeErweitern_Click

    Dim ldb As DAO.Database
    Dim rs As Recordset
    Dim MID As Integer
    Dim Maßnahme As String

    If Me.ListeErweiterung.ListCount = 0 Then
        MID = Me.MaßnahmeID
        TextMaßnahme = Me.Maßnahme

        Set ldb = openAccDBConnection("openDB")
        Set rs = ldb.OpenRecordset("Maßnahmen")
            rs.AddNew
            rs("Maßnahme") = "Erweiterung von ID " & MID
            rs("Priorität") = 1
            rs("MaßnahmenHerkunftID") = 5
            rs.Update
            rs.MoveLast
            Me.Requery
            Me.Refresh

        DoCmd.Close acForm, "MaßnahmenDialog"
        DoCmd.OpenForm "MaßnahmenDialog", acNormal, , , acFormAdd
        Forms("MaßnahmenDialog").Filter = "[Maßnahmen].[MaßnahmeID] = " & rs![MaßnahmeID]
        Forms("MaßnahmenDialog").FilterOn = True
        Forms("MaßnahmenDialog").Requery
        Forms("MaßnahmenDialog").Refresh
        Forms("MaßnahmenDialog").AllowAdditions = False

        ldb.Execute ("Insert into ErMIDZuMID ([MID],[ErMID]) values (" & MID & ", " & rs![MaßnahmeID] & ");")
        Forms("MaßnahmenDialog").Requery
        Forms("MaßnahmenDialog").Refresh

        Me.SetFocus

        Set rs = Nothing
    Else
'        MsgBox ("Es existiert schon eine Erweiterung zu dieser Maßnahme")
        MsgBox getTC_FCFromCode("TC3207"), vbInformation, getTC_FCFromCode("TC3208")
        Exit Sub
    End If

Exit_CmdMaßnahmeErweitern_Click:
    Exit Sub

Err_CmdMaßnahmeErweitern_Click:
    MsgBox Err.Description
    Resume Exit_CmdMaßnahmeErweitern_Click

End Sub
