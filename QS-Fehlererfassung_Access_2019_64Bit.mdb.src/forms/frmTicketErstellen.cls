Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

#If VBA7 Or Win64 Then
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
#End If


Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_DLGMODALFRAME As Long = &H1
Private mRefMailSenden As Form_MailSendenAnNeu

'EventHandler Subscriber
Private WithEvents MessageForm As Form_MailSendenAnNeu
Attribute MessageForm.VB_VarHelpID = -1
Dim plTicketID As Long

Private Sub cboFilterForm_Click()

'    Dim ldb As DAO.Database
'    Dim Date1 As String
'    Dim Schaltjahr As String
'    Dim strFormFilter As String
'    Dim rsFormFilter As DAO.Recordset
'    Dim ld_date As String
'    Dim idKuerzel As Integer
'
'    idKuerzel = DLookup("[SB-ID]", "Sachbearbeiter", "[Kurzzeichen] = '" & get_pKuerzel & "'")
'
'    Set ldb = openAccDBConnection("openDB")
'    Set rsFormFilter = ldb.OpenRecordset("Select * from tblFilterForm where [FilterBezeichnung] = '" & Me.cboFilterForm & "' And [FormName]= '" & Me.Name & "' And [SachbearbeiterID] = " & idKuerzel & " ;")
'
'    strFormFilter = rsFormFilter("FilterParameter")
'
'    DoCmd.OpenForm "frmTicketListe", acNormal
'    Forms("frmTicketListe").Filter = strFormFilter ' & sAuftrag
'    Forms("frmTicketListe").FilterOn = True
'    Me.Requery
'    Me.Refresh
'
'
'    'DoCmd.Maximize
'
'    Set rsFormFilter = Nothing


End Sub

Private Sub cboFilterLQEinheit_AfterUpdate()

    setFilterTicketliste

End Sub

Private Sub cboPersFilterSpeichern_Click()
'    Dim ldb As DAO.Database
'    Dim strFilterBezeichnung As String
'    Dim strFilterparameter As String
'
'    strFilterparameter = Me.Filter
'    strFilterparameter = Replace(strFilterparameter, "'", "''")
'
'    Set ldb = openAccDBConnection("openDB")
''    strFilterBezeichnung = InputBox("Bitte Filterbezeichnung eingeben")
''    strFilterBezeichnung = InputBox(getTC_FCFromCode("IC3845"), getTC_FCFromCode("IC3846"))
'    strFilterBezeichnung = getValueFromInputBox(getTC_FCFromCode("IC3845"), getTC_FCFromCode("IC3846"))
'    If strFilterBezeichnung = "" Then Exit Sub
'
'
'    ldb.Execute ("Insert into tblFilterForm ([SachbearbeiterID], [FormName], [FilterParameter], [FilterBezeichnung], [Aktiv], [Dynamisch], [DatumVon], [DatumBis])" & _
'                "values(" & get_SachbearbeiterID() & ", '" & Me.Name & "' , '" & Replace(Me.Filter, "'", "''") & "', '" & strFilterBezeichnung & "' , True, True, '','');")
'
'
'    Me.Requery
'    Me.Refresh
End Sub

Private Sub cboSachbearbeiter_AfterUpdate()

    setFilterTicketliste

End Sub

Private Sub cboSachbearbeiter2_AfterUpdate()        '20191009-ROS

    setFilterTicketliste

End Sub

Private Sub cboTicketStatusFilter_AfterUpdate()

    setFilterTicketliste

End Sub

Private Sub cboTicketTyp_AfterUpdate()

    If IsNull(Me.cboTicketTyp) Then
            Me.txtBearbeitungBis = ""
        Else
'             Me.txtBearbeitungBis = getBearbDateTicketTyp(Me.cboTicketTyp.Column(0, 0))
             Me.txtBearbeitungBis = getBearbDateTicketTyp(Me.cboTicketTyp)
    End If

'    Me.Requery
'    Me.Refresh

End Sub

Private Sub cboTicketTypFilter_AfterUpdate()

     setFilterTicketliste

End Sub

Private Sub cboVorgangStatus_AfterUpdate()

    setFilterTicketliste

End Sub

Private Sub chkAlleOffenenTicket_AfterUpdate()

    setFilterTicketliste

End Sub

Private Sub chkAlleOffenenTicket_Click()

'    setFilterTicketliste

End Sub



Private Sub cmbOdrnerOeffnen_Click()

    Dim ReklnrKnd As String

    If Me.TicketTypID = 11 Or Me.TicketTypID = 12 Then
        MsgBox getTC_FCFromCode("TC4255"), vbInformation, getTC_FCFromCode("TC4256")
    Else
        ReklnrKnd = Nz(DLookup("[Meldungsnummer Kunde]", "Mastertab", "[ID] = " & Me.ReklID & ""))
        Call VorgangsOrdner(Me.ReklID, ReklnrKnd, Me.MailDatum, True)   '20250729-ROS
    End If


End Sub

Private Sub cmbTicketLöschen_Click()
'    On Error GoTo Err

    TicketLoeschen (Me.TicketID)
'
'    Dim TicketID As Long
'    Dim ldb As DAO.Database
'
'        Set ldb = openAccDBConnection("openDB")
'
'        'Tickets können nur gelöscht werden, wenn Ticketstatus = Erstellt
'        If DLookup("[TicketStatusID]", "tblTicketListe", "[TicketID] = " & Me.TicketID) = 1 Then   ' 1 = Erstellt
'            TicketID = Me.TicketID
'        Else
'            Exit Sub
'        End If
'
'        'Sicherheitsabfrage ob Ticket wirklich gelöscht werden soll
'        If MsgBox(getTC_FCFromCode("TC4213"), vbYesNo, getTC_FCFromCode("TC3210")) = vbYes Then
'            DoCmd.SetWarnings False
'            ldb.Execute ("Delete * from tblTicketListe where [TicketID] = " & TicketID & ";")
'            DoCmd.SetWarnings True
'        Else
'            Exit Sub
'        End If

        Me.Requery
        Me.Refresh

        If IsLoaded("Fehlererfassung - Register") Then
            Forms("Fehlererfassung - Register").Requery
            Forms("Fehlererfassung - Register").Refresh
        End If


'Err:
'    If Err.Number = 2427 Then
''        MsgBox ("Kein Datensatz ausgewählt")
'        MsgBox getTC_FCFromCode("TC3240"), vbExclamation
'        Exit Sub
'    End If


End Sub

Private Sub cmbVorgangOeffnen_Click()

    If Me.TicketTypID = 11 Or Me.TicketTypID = 12 Then
        DoCmd.OpenForm "MaßnahmenDialog", acNormal, , , acFormAdd
        Forms("MaßnahmenDialog").Filter = "[Maßnahmen].[MaßnahmeID] = " & Me.ReklID
        Forms("MaßnahmenDialog").FilterOn = True
        Forms("MaßnahmenDialog").Requery
        Forms("MaßnahmenDialog").Refresh
        Forms("MaßnahmenDialog").AllowAdditions = False
    Else
        p_ReklNr = Me.ReklID
        Call NewReklamationE("Bearbeiten", get_p_ReklNr())
        Forms("Fehlererfassung - Register").SetFocus
    End If

End Sub

Private Sub cmbTicketListeSchließen_Click()
On Error GoTo Err_cmbTicketListeSchließen_Click


    FormKoorSave Me

    If Me.Dirty Then Me.Dirty = False
    DoCmd.Close

Exit_cmbTicketListeSchließen_Click:
    Exit Sub

Err_cmbTicketListeSchließen_Click:
    MsgBox Err.Description
    Resume Exit_cmbTicketListeSchließen_Click

End Sub

Private Sub cmbFilterLöschen_Click()
'On Error GoTo Err_cmbFilterLöschen_Click
'
'    Me.cboFilterLQEinheit = ""
'    Me.cboTicketTypFilter = ""
'    Me.cboTicketStatusFilter = ""
'    Me.txtReklID = ""
'    Me.cboSachbearbeiter = ""
'    Me.cboVorgangStatus = ""
'    Me.FilterOn = False
'    Me.chkAlleOffenenTicket = 0
'    Me.Requery
'    Me.Refresh
'    Me.cboFilterForm = ""
'
'Exit_cmbFilterLöschen_Click:
'    Exit Sub
'
'Err_cmbFilterLöschen_Click:
'    MsgBox Err.Description
'    Resume Exit_cmbFilterLöschen_Click

End Sub
Private Sub cmbNeuesTicket_Click2(fNew As Boolean, flTicketID As Long) 'Optional ByRef frsTicket As DAO.Recordset = Null
'###########################
' Erstellt das Ticket oder ermittelt das selektierte Ticket
' Parameter: FNew - True: Neues Ticket erstellen
'                   False: Bestehendes Ticket anzeigen
'
'############################
    On Error Resume Next

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
'    Dim lTicketID As Long
    Dim lngCurrentID As Long
    Dim lAbsender As String

    Set ldb = openAccDBConnection("openDB")

    'Neues oder bestehendes Ticket bearbeiten oder anlegen
    If fNew = False Then
        'bestehendes Ticket
        Set rs = ldb.OpenRecordset("Select * from tblTicketListe where [TicketID] = " & flTicketID & " And [TicketStatusID] = 3 ;")
        'Ist Ticket vorhanden und schon geschlossen, dann mit Fehlerausgabe - Ticket geschlossen - Vorgang abbrechen. Vorgang wird nicht geöffnet
        If rs.RecordCount Then
'            MsgBox ("Ticket kann nicht geöffnet werden, ist schon Abgeschlossen. Bitte neues Ticket erstellen")
            MsgBox getTC_FCFromCode("TC3241"), vbExclamation, getTC_FCFromCode("TC3210")
            Set rs = Nothing
            Exit Sub
        Else
            'Neues berechnetes Bearbeitungsdatum anhand TicketTyp im Mailtext, wenn vorhanden ersetzen!
            If DLookup("[TicketStatusID]", "tblTicketListe", "[TicketID] = " & flTicketID) = 1 Then
                'Aktuelles Ticket ermitteln und bearbeiten
                Set rs = ldb.OpenRecordset("Select * from tblTicketListe where [TicketID] = " & flTicketID & ";")
                rs.Edit
                rs("MailText") = Replace(rs("MailText"), format(rs("BearbeitungBis"), "dd.mm.yyyy"), getBearbDateTicketTyp(rs("TicketTypID")))
                rs("BearbeitungBis") = getBearbDateTicketTyp(rs("TicketTypID"))
                rs.Update
    '            lngCurrentID = Me.TicketID   'Me.TicketID wird Zwischengespeichert da nach der Datensatzdem Update seinen Fokus verliert.
            End If
            lngCurrentID = flTicketID   'Me.TicketID wird Zwischengespeichert da nach der Datensatz dem Update seinen Fokus verliert.
            DoCmd.Save acForm, Me.Name
    '        Me.Requery
            Me.RecordsetClone.FindFirst "[TicketID] =" & lngCurrentID
            Me.Bookmark = Me.RecordsetClone.Bookmark
            Set rs = ldb.OpenRecordset(("Select * from qryTicketListe where [TicketID] = " & flTicketID & ";"))
        End If
    Else
        If IsNull(Me.cboTicketTyp) Then
'            MsgBox ("Es ist kein Ticket-Typ ausgewählt")
            MsgBox getTC_FCFromCode("TC3242"), vbInformation, getTC_FCFromCode("TC3210")
            Exit Sub
        End If

        plTicketID = getNewTicketID(Me.txtReklID, Me.cboTicketTyp)

        lAbsender = Nz(DLookup("[MailVon]", "tblTicketListe", "[TicketID] = " & plTicketID & " "))

        If lAbsender = "" Then Exit Sub

        Set rs = ldb.OpenRecordset("Select * from qryTicketListe where [TicketID] = " & plTicketID & ";")

    End If

    'Instanz des MailDialog erzeugen und festhalten
    Set mRefMailSenden = New Form_MailSendenAnNeu
    mRefMailSenden.Visible = True
    Set MessageForm = mRefMailSenden
    'Übergabe der eigenen Instanz an den MailDialog als Referenz
    mRefMailSenden.parentForm = Me
    'Property setzen ob es sich hier um ein neues oder bestehendes Ticket handelt
    mRefMailSenden.LetTicketNew = fNew
    mRefMailSenden.loadForm rs ' Hier werden die Daten via Recordset übergeben

    Set rs = Nothing


End Sub
Private Sub cmbNeuesTicket_Click() 'Optional ByRef frsTicket As DAO.Recordset = Null

    Dim rs As DAO.Recordset
    Dim ldb As DAO.Database
    Dim MehrfachTicket As Boolean

    If Me.txtBearbeitungBis < Date Then     '20190821-ROS
        MsgBox getTC_FCFromCode("TC4036"), vbExclamation, getTC_FCFromCode("TC4037")
        Exit Sub
    End If

    If IsNull(Me.cboTicketTyp) Then
'        MsgBox ("Bitte Ticket-Typ auswählen")
        MsgBox getTC_FCFromCode("TC3242"), vbExclamation, getTC_FCFromCode("TC3210")
        Exit Sub
    End If

    Set ldb = openAccDBConnection("openDB")
    'Abfrage ob es zu der Reklamation schon ein offenes Ticket mit dem eingestellten Ticket-Typ gibt 'RSP_8_2_5
    Set rs = ldb.OpenRecordset("SELECT tblTicketListe.* From tblTicketListe WHERE (((tblTicketListe.ReklID) = " & Me.txtReklID & ") And " & _
                                "((tblTicketListe.TicketTypID) = " & Me.cboTicketTyp & ") And ((tblTicketListe.TicketStatusID) <> 3)) ORDER BY tblTicketListe.TicketTypID;", dbOpenSnapshot)

    'Einige Tickets können mehrfach erstellt werden wie z.b. Lagerprüfung oder sonstige Klärung
    'Prüfung ob Haken gesetzt ist
    MehrfachTicket = DLookup("[MehrfachTicket]", "tblTicketTyp", "[TicketTypID] = " & Me.cboTicketTyp & " ")    '20240822-ROS
    If MehrfachTicket = True Then
        Call cmbNeuesTicket_Click2(True, 0)
        Exit Sub
    Else
        'Ticket-Typ noch nicht vorhanden wir ein neues Ticket erstellt
        If rs.RecordCount = 0 Then
            Call cmbNeuesTicket_Click2(True, 0)
        Else
        'wenn Ticket schon vorhanden dann Ausgabe Fehlermeldung
'            MsgBox "VorgangsTyp schon angelegt und nicht abgeschlossen! Bestehenden Vorgang erst abschließen!"
            MsgBox getTC_FCFromCode("TC4232"), vbExclamation, getTC_FCFromCode("TC4233")
        End If
    End If

End Sub


Private Sub Form_Click()

   If Me.Recordset.RecordCount > 0 Then
        Me.txtBearbeitungBis = Me.BearbeitungBis
        Me.cboTicketTyp = Me.TicketTyp
        Me.txtBearbeitungBis.Requery
        Me.cboTicketTyp.Requery

        Me.[frmUFMailAnhaenge].Form.Filter = " TicketID = " & Me.TicketID
        Me.[frmUFMailAnhaenge].Form.FilterOn = True
        Me.[frmUFMailAnhaenge].Requery
    End If


End Sub

Private Sub Form_Current()

    Me.[frmUFMailAnhaenge].Form.Filter = " TicketID = " & Me.TicketID
    'me.[frmUFMailAnhaenge].Form.Filter = " TicketID = " & "73937"
    Me.[frmUFMailAnhaenge].Form.FilterOn = True
    Me.[frmUFMailAnhaenge].Requery

End Sub

Private Sub Form_DblClick(Cancel As Integer)

    plTicketID = Me.TicketID
    Call cmbNeuesTicket_Click2(False, Me.TicketID)

End Sub


Private Sub Form_Load()

    Dim LQQS As String
    Dim lForm As [Form_Fehlererfassung - Register]

    Dim objMouseOverForm As clsMouseoverForm
    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With



    If IsFormDialog(Me) = True Then
'        Me.Rechteck71.Visible = False
'        Me.chkAlleOffenenTicket.Visible = False
'        Me.txtAlleOffenenTicket.Visible = False
'        Me.cboSachbearbeiter.Visible = False
'        Me.BezSachbearbeiter.Visible = False
'        Me.txtReklID.Visible = False
'        Me.cboTicketTypFilter.Visible = False
'        Me.cboTicketStatusFilter.Visible = False
'        Me.cboFilterLQEinheit.Visible = False
'        Me.cmbFilterLöschen.Visible = False
        Me.cmbNeuesTicket.Visible = True
        Me.Rechteck39.Visible = True
        Me.BezNeuesTicketErstellen.Visible = True
        Me.BezTicketTyp.Visible = True
        Me.cboTicketTyp.Visible = True
        Me.BezBearbeitungBis.Visible = True
        Me.txtBearbeitungBis.Visible = True
        Me.VerlängertBis.Locked = False
        Me.ErledigtAm.Locked = False
'        Me.cmbVorgangOeffnen.Visible = False
'        Me.cboVorgangStatus.Visible = False
'        Me.cboFilterForm.Visible = False
'        Me.BezFilterForm.Visible = False
'        Me.cboPersFilterSpeichern.Visible = False
'        Me.cboSachbearbeiter2.Visible = False
'        Me.Bezeichnungsfeld77.Visible = False
'        Me.ErledigtVon.Visible = False
'        Me.cmbOdrnerOeffnen.Visible = False
    Else
        Me.cmbNeuesTicket.Visible = False
'        Me.chkAlleOffenenTicket = 0
    End If


'    Set lForm = Forms("Fehlererfassung - Register")
'    Me.txtReklID = lForm.Controls("Mastertab#ID")

    'If CurrentProject.AllForms("Fehlererfassung - Register").IsLoaded = True Then
    If Me.OpenArgs = "Liste selektiert" Then

        Set lForm = Forms("Fehlererfassung - Register")
        Me.txtReklID = lForm.Controls("Mastertab#ID")

        'Me.Filter = "(ErledigtAM Is Null) And (ReklID = " & Me.Controls("Mastertab#ID") & " AND TicketTypID <> 11 AND TicketTypID <> 12)"
        Me.RecordSource = "SELECT tblTicketListe.*, tblTicketStatus.TicketStatusID, tblTicketTyp.Anzeige, tblTicketStatus.Anzeige, tblTicketListe.VerlängertBis, tblTicketStatus.Anzeige, " & _
                          "[Kategorien-Vorgangstatus].Anzeige " & _
                          "FROM (((tblTicketListe LEFT JOIN tblTicketStatus ON tblTicketListe.TicketStatusID = tblTicketStatus.TicketStatusID) " & _
                          "LEFT JOIN tblTicketTyp ON tblTicketListe.TicketTypID = tblTicketTyp.TicketTypID) LEFT JOIN Mastertab ON tblTicketListe.ReklID = Mastertab.ID) " & _
                          "LEFT JOIN [Kategorien-Vorgangstatus] ON Mastertab.[Kategorie-Status] = [Kategorien-Vorgangstatus].[Kategorie-Vorgangstatus] " & _
                          "WHERE [ReklID] = " & lForm.Controls("Mastertab#ID") & " AND [tblTicketTyp].TicketTypID <> 11 AND [tblTicketTyp].TicketTypID <> 12 " & _
                          "ORDER BY tblTicketListe.VerlängertBis;"

    ElseIf Me.OpenArgs = "Neu erstellen" Then
        Set lForm = Forms("Fehlererfassung - Register")
        Me.txtReklID = lForm.Controls("Mastertab#ID")
        Me.RecordSource = "SELECT tblTicketListe.*, tblTicketStatus.TicketStatusID, tblTicketTyp.Anzeige, tblTicketStatus.Anzeige, tblTicketListe.VerlängertBis, tblTicketStatus.Anzeige, " & _
                  "[Kategorien-Vorgangstatus].Anzeige " & _
                  "FROM (((tblTicketListe LEFT JOIN tblTicketStatus ON tblTicketListe.TicketStatusID = tblTicketStatus.TicketStatusID) " & _
                  "LEFT JOIN tblTicketTyp ON tblTicketListe.TicketTypID = tblTicketTyp.TicketTypID) LEFT JOIN Mastertab ON tblTicketListe.ReklID = Mastertab.ID) " & _
                  "LEFT JOIN [Kategorien-Vorgangstatus] ON Mastertab.[Kategorie-Status] = [Kategorien-Vorgangstatus].[Kategorie-Vorgangstatus] " & _
                  "WHERE [ErledigtAM] is Null And [ReklID] = " & lForm.Controls("Mastertab#ID") & " AND [tblTicketTyp].TicketTypID <> 11 AND [tblTicketTyp].TicketTypID <> 12 " & _
                  "ORDER BY tblTicketListe.VerlängertBis;"

    Else
        Me.RecordSource = "SELECT tblTicketListe.*, tblTicketStatus.TicketStatusID, tblTicketTyp.Anzeige, tblTicketStatus.Anzeige, LQEinheit.LQEinheit, tblTicketListe.VerlängertBis, " & _
                          "qryTicketSacharbeiterKz.Kurzzeichen, tblTicketStatus.Anzeige, [Kategorien-Vorgangstatus].Anzeige " & _
                          "FROM (((Mastertab LEFT JOIN LQEinheit ON Mastertab.LQEinheitID = LQEinheit.ID) " & _
                          "RIGHT JOIN ((tblTicketListe LEFT JOIN tblTicketStatus ON tblTicketListe.TicketStatusID = tblTicketStatus.TicketStatusID) " & _
                          "LEFT JOIN tblTicketTyp ON tblTicketListe.TicketTypID = tblTicketTyp.TicketTypID) ON Mastertab.ID = tblTicketListe.ReklID) " & _
                          "LEFT JOIN qryTicketSacharbeiterKz ON tblTicketListe.ReklID = qryTicketSacharbeiterKz.VorgangID) " & _
                          "LEFT JOIN [Kategorien-Vorgangstatus] ON Mastertab.[Kategorie-Status] = [Kategorien-Vorgangstatus].[Kategorie-Vorgangstatus] " & _
                          "ORDER BY tblTicketListe.VerlängertBis;"
'        Me.chkAlleOffenenTicket = -1
        setFilterTicketliste
    End If

    If get_ServiceZuordnung = False Then
'        LQQS = "select * from tblTicketTyp WHERE AnsichtSteuerung = True;"
        LQQS = "SELECT [TicketTypID],[Anzeige],[AnsichtSteuerung] From tblTicketTyp WHERE [AnsichtSteuerung]=True;"     '20191016-ROS
        Me.cboTicketTyp.RowSource = LQQS
    End If

    Set lForm = Nothing

    Call FormEinUndAusBlenden(Me)
    FormOpenKoor Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub
Private Sub setFilterTicketliste()

'    Dim ls_Filter As String
'
'    ls_Filter = ""
'
'
''    If Me.chkAlleOffenenTicket = 0 Then
''        ls_Filter = ls_Filter & "(tblTicketListe.TicketStatusID = 3) AND "
''    End If
'
'    If Me.chkAlleOffenenTicket = -1 Then
'        ls_Filter = ls_Filter & "(tblTicketListe.TicketStatusID <> 3) AND "
'    End If
'
'    If Me.txtReklID <> "" Then
'        ls_Filter = ls_Filter & "(ReklID = " & Me.txtReklID & ") AND "
'    End If
'    If Me.cboTicketTypFilter <> "" Then
'        ls_Filter = ls_Filter & "(tblTicketTyp.Anzeige = '" & Me.cboTicketTypFilter & "') AND "
'    End If
'    If Me.cboTicketStatusFilter <> "" Then
'        ls_Filter = ls_Filter & "(tblTicketStatus.Anzeige = '" & Me.cboTicketStatusFilter & "') AND "
'    End If
'    If Me.cboSachbearbeiter <> "" Then
'        ls_Filter = ls_Filter & "(Kurzzeichen = '" & Me.cboSachbearbeiter & "') AND "
'    End If
'
'    If Me.cboSachbearbeiter2 <> "" Then                                                 '20191009-ROS
'        ls_Filter = ls_Filter & "(ErledigtVon = '" & Me.cboSachbearbeiter2 & "') AND "
'    End If
'
'    If Me.cboVorgangStatus <> "" Then
'        ls_Filter = ls_Filter & "([Kategorien-Vorgangstatus].[Anzeige] = '" & Me.cboVorgangStatus & "') AND "
'    End If
'
'    If Me.cboFilterLQEinheit <> "" Then
'        ls_Filter = ls_Filter & "(LQEinheit = '" & Me.cboFilterLQEinheit & "')"
'    End If
'
'    If Right(ls_Filter, 5) = " AND " Then
'        Me.Filter = Left(ls_Filter, Len(ls_Filter) - 5)
'    Else
'        Me.Filter = ls_Filter
'    End If
'
'
'    If Me.Recordset.RecordCount = 0 Then
'        MsgBox getTC_FCFromCode("TC4207"), vbInformation, getTC_FCFromCode("TC4208")
'        Me.cboFilterLQEinheit = ""
'        Me.cboTicketTypFilter = ""
'        Me.cboTicketStatusFilter = ""
'        Me.txtReklID = ""
'        Me.cboSachbearbeiter = ""
'        Me.cboVorgangStatus = ""
'        Me.FilterOn = False
'        Me.chkAlleOffenenTicket = 0
'        Me.cboFilterForm = ""
'        Me.Refresh
'    Else
'        Me.FilterOn = True
'    End If


End Sub

Private Sub Form_Open(Cancel As Integer)

    If get_AdminZuordnung = True Then
        Me.cmbTicketLöschen.Visible = True
    Else
        Me.cmbTicketLöschen.Visible = False
    End If


End Sub

Private Sub txtBearbeitungBis_AfterUpdate()

    If Me.txtBearbeitungBis < Date Then     '20190821-ROS
        MsgBox getTC_FCFromCode("TC4036"), vbExclamation, getTC_FCFromCode("TC4037")
        Exit Sub
    End If

End Sub



Private Sub txtReklID_AfterUpdate()

        setFilterTicketliste

End Sub


Private Function getNewTicketID(ReklID As Long, TicketTypID As Long) As Long
    On Error Resume Next


    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim rs2 As DAO.Recordset
    Dim rs3 As DAO.Recordset
    Dim rs5 As DAO.Recordset
    Dim MailAn As String
    Dim MailCC As String
    Dim MailBCC As String
    Dim s As String
    Dim MailVon As String
    Dim Location As String



    '20181022-ROS
    '########################################################################################################################
    Dim Report As String
    Dim FirmaID As Integer
    Dim LieferantID As Integer
    Dim JaNein As Boolean

    Select Case TicketTypID
        Case 1
'            FirmaID = DLookup("[Firma_ID]", "tblFirma", "[Firma_Name] = '" & Forms("Fehlererfassung - Register").Firma_Name & "'")
            FirmaID = Forms("Fehlererfassung - Register").Controls("Mastertab#Kunden-ID#1")
            JaNein = DLookup("[8DReport]", "tblReklBearbZeit", "[KundenID] = " & FirmaID & " ")
        Case 2
            LieferantID = DLookup("[Lieferant-ID]", "Lieferanten", "[Lieferant-ID] = " & Forms("Fehlererfassung - Register").Controls("Mastertab#Lieferanten-Code 1") & "")
            JaNein = DLookup("[8DReport]", "tblReklBearbZeit", "[LiefID] = " & LieferantID & " ")
    End Select

'    JaNein = DLookup("[8DReport]", "tblReklBearbZeit", "[KundenID] = " & FirmaID & " or [LiefID] = " & LieferantID & " ")

    If JaNein = True Then
         Report = getTC_FCFromCode("TC0484")    'Ja
    Else
         Report = getTC_FCFromCode("TC0485")    'Nein
    End If
    '########################################################################################################################

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("tblTicketListe")
    Set rs2 = ldb.OpenRecordset("select * from qryTicketQSLQ where[ID] = " & Me.txtReklID & "; ") 'Abfrage zur Ermittlung der LQEinheit und Sachbearbeiter
    Set rs3 = ldb.OpenRecordset("select * from tblTicketTyp where[TicketTypID] = " & TicketTypID & "; ")
    '13 = Vertrieb ersetzten durch Konstante LQEinheit
    Set rs5 = ldb.OpenRecordset("select * from qryTicketLQEinheit where[ID] = 13 ; ")

    'Bei Lagerprüfung muss explizit die Einheit über einen Dialog angegeben werden, bei der die Lagerprüfung stattfinden soll
'    If TicketTypID = 8 Then 'Lager-Prüfung 20181217-ROS
'        DoCmd.OpenForm "frmLagerPrüfung", , , , , acDialog
'        'Im Dialog werden die public Parameter beschrieben
'        MailAn = ps_MailAn
'        MailCC = ps_MailCC
'        MailBCC = ps_MailBCC
'    Else
        s = get_EMailAdressen(MailAn, MailCC, MailBCC)      '20181213-ROS
'    End If

    Location = DLookup("Value", "KonfigurationReklDBLocal", "Konfiguration = 'Location'")

    'Versendete Mails durch Servicemitarbeiter werdden auf die Service@lQ-group.com umgehoben
    If get_ServiceZuordnung = True And Location = "German" Then
        MailVon = DLookup("EmailName", "Sachbearbeiter", "[Name] = '" & fOSUserName & "'")
    ElseIf IsNull(rs2("ANE")) Then
        MailVon = rs2("CCSACHB")
    Else
        MailVon = rs2("ANE")
    End If

    rs.AddNew
    'Rückgabewert Neue TicketID setzen
    getNewTicketID = rs("TicketID")
    rs("TicketTypID") = TicketTypID
    rs("ReklID") = txtReklID
    rs("BearbeitungBis") = Me.txtBearbeitungBis
    rs("TicketStatusID") = 1 'Erstellt
    rs("MailErstellt") = Date
    rs("MailVon") = MailVon
    rs("MailAn") = MailAn
    rs("MailCC") = MailCC
    rs("MailBCC") = MailBCC


    Select Case TicketTypID

        Case 1 'QS Bearbeitung LQ
'            rs("MailVon") = MailVon
'            rs("MailAn") = MailAn
'            rs("MailCC") = MailCC
'            rs("MailBCC") = MailBCC
            rs("LinkZuVorgangsOrdner") = getLinkVorgangsordner
            'Betreff erstellen
            rs("MailBetreff") = rs3![MailBetreff]
            rs("MailBetreff") = Replace(rs("MailBetreff"), "<ReklNr>", Me.txtReklID)
            rs("MailBetreff") = Replace(rs("MailBetreff"), "<FirmaName>", rs2![Firma_Name]) & " / Masch.Nr.: " & rs2("Maschinennummer")
            'Body erstellen
            rs("MailText") = rs3![MailText]
            rs("MailText") = Replace(rs3![MailText], "<JaNein>", Report)    '20181022-ROS
            rs("MailText") = Replace(rs("MailText"), "<ReklNr>", Me.txtReklID)
            rs("MailText") = Replace(rs("MailText"), "<Bis>", rs![BearbeitungBis])
            rs("MailText") = Replace(rs("MailText"), "<Fehlerbeschreibung>", rs2![Fehlerbeschreibung])

        Case 2 'QS Bearbeitung Lieferant
    '                If IsNull(rs2("ANE")) Then
    '                    rs("MailVon") = rs2("CCSACHB")
    '                Else
    '                    rs("MailVon") = rs2![ANE]
    '                End If
            'rs("MailVon") = MailVon
            rs("MailAn") = rs2![ANLIEF]
            rs("MailCC") = rs2![CCSACHB]
            'Ermitteln BCC anhand TicketTyp Feld EskalationsAdresse
            rs("MailBCC") = rs3("EskalationsAdresse").Value
            'Betreff erstellen
            rs("MailBetreff") = rs3![MailBetreff]
            rs("MailBetreff") = Replace(rs("MailBetreff"), "<ReklNr>", txtReklID)
            'Body erstellen
            rs("MailText") = rs3![MailText]
            rs("MailText") = Replace(rs3![MailText], "<JaNein>", Report)    '20181022-ROS
            rs("MailText") = Replace(rs("MailText"), "<Bis>", rs![BearbeitungBis])
            rs("MailText") = Replace(rs("MailText"), "<ReklNr>", Me.txtReklID)
            rs("MailText") = Replace(rs("MailText"), "<E-Mail_Absender>", rs("MailVon"))
            rs("MailText") = Replace(rs("MailText"), "<Telefon>", rs2![Telefon])

        Case 3 'Kalkulation
'            rs("MailVon") = MailVon
'            rs("MailAn") = MailAn
'            rs("MailCC") = MailCC
'            rs("MailBCC") = MailBCC
            rs("LinkZuVorgangsOrdner") = getLinkVorgangsordner
            'Betreff erstellen
            rs("MailBetreff") = rs3![MailBetreff]
            rs("MailBetreff") = Replace(rs("MailBetreff"), "<ReklNr>", Me.txtReklID)
            'Body erstellen
            rs("MailText") = rs3![MailText]
            rs("MailText") = Replace(rs("MailText"), "<Bis>", rs![BearbeitungBis])
            rs("MailText") = Replace(rs("MailText"), "<ReklNr>", ReklID)

        Case 4 'Rechnungskorrektur Vertrieb
'            rs("MailVon") = MailVon
'            rs("MailAn") = MailAn
'            rs("MailCC") = MailCC
'            rs("MailBCC") = MailBCC
            rs("LinkZuVorgangsOrdner") = getLinkVorgangsordner
            'Betreff erstellen
            rs("MailBetreff") = rs3![MailBetreff]
            rs("MailBetreff") = Replace(rs("MailBetreff"), "<ReklNr>", Me.txtReklID)
            'Body erstellen
            rs("MailText") = rs3![MailText]
            rs("MailText") = Replace(rs("MailText"), "<Bis>", rs![BearbeitungBis])
            rs("MailText") = Replace(rs("MailText"), "<ReklNr>", txtReklID)

        Case 5 'Reparatur-Abschätzung
'            rs("MailVon") = MailVon
'            rs("MailAn") = MailAn
'            rs("MailCC") = MailCC
'            rs("MailBCC") = MailBCC
            rs("LinkZuVorgangsOrdner") = getLinkVorgangsordner
            'Betreff erstellen
            rs("MailBetreff") = rs3![MailBetreff]
            rs("MailBetreff") = Replace(rs("MailBetreff"), "<ReklNr>", Me.txtReklID)
            'Body erstellen
            rs("MailText") = rs3![MailText]
            rs("MailText") = Replace(rs("MailText"), "<Bis>", rs![BearbeitungBis])
            rs("MailText") = Replace(rs("MailText"), "<ReklNr>", txtReklID)

        Case 6 'Reparatur

'            rs("MailVon") = MailVon
'            rs("MailAn") = MailAn
'            rs("MailCC") = MailCC
'            rs("MailBCC") = MailBCC
            rs("LinkZuVorgangsOrdner") = getLinkVorgangsordner
            'Betreff erstellen
            rs("MailBetreff") = rs3![MailBetreff]
            rs("MailBetreff") = Replace(rs("MailBetreff"), "<ReklNr>", Me.txtReklID)
            'Body erstellen
            rs("MailText") = rs3![MailText]
            rs("MailText") = Replace(rs("MailText"), "<Bis>", rs![BearbeitungBis])
            rs("MailText") = Replace(rs("MailText"), "<ReklNr>", txtReklID)

        Case 7 'sonstige Klärung
            If get_ServiceZuordnung = True And getCurrLanguage = "German" Then
                rs("MailVon") = getTC_FCFromCode("TC0510")        'service@de.lq-group.com
            Else
                If IsNull(rs2![CCSACHB]) Or rs2![CCSACHB] = "" Then
                    MsgBox getTC_FCFromCode("TC4243"), vbExclamation, getTC_FCFromCode("TC3214")
                    Exit Function
                Else
                    rs("MailVon") = rs2![CCSACHB]
                End If
            End If
            'Betreff erstellen
            rs("MailBetreff") = rs3![MailBetreff]
            rs("MailBetreff") = Replace(rs("MailBetreff"), "<ReklNr>", Me.txtReklID)
            rs("MailBetreff") = Replace(rs("MailBetreff"), "<FirmaName>", rs2![Firma_Name]) & " / Masch.Nr.: " & rs2("Maschinennummer")

        Case 8 'Lager-Prüfung
            'Absender ermitteln, da es mehrer Adressaten geben kann die eine Lagerprüfung durchführen müssen
            DoCmd.OpenForm "frmLagerPrüfung", , , , , acDialog
            'Im Dialog werden die public Parameter beschrieben
'            MailAn = ps_MailAn
'            MailCC = ps_MailCC
'            MailBCC = ps_MailBCC

            'rs("MailVon") = MailVon
            rs("MailAn") = ps_MailAn
            rs("MailCC") = ps_MailCC
            rs("MailBCC") = ps_MailBCC
            rs("LinkZuVorgangsOrdner") = getLinkVorgangsordner
            'Betreff erstellen
            rs("MailBetreff") = rs3![MailBetreff]
            rs("MailBetreff") = Replace(rs("MailBetreff"), "<ReklNr>", Me.txtReklID)
            'Body erstellen
            rs("MailText") = rs3![MailText]
            rs("MailText") = Replace(rs("MailText"), "<Bis>", rs![BearbeitungBis])
            rs("MailText") = Replace(rs("MailText"), "<ReklNr>", txtReklID)

        Case 9 'Rechnungskorrektur PE
'            rs("MailVon") = MailVon
'            rs("MailAn") = MailAn
'            rs("MailCC") = MailCC
'            rs("MailBCC") = MailBCC
            rs("LinkZuVorgangsOrdner") = getLinkVorgangsordner
            'Betreff erstellen
            rs("MailBetreff") = rs3![MailBetreff]
            rs("MailBetreff") = Replace(rs("MailBetreff"), "<ReklNr>", Me.txtReklID)
            'Body erstellen
            rs("MailText") = rs3![MailText]
            rs("MailText") = Replace(rs("MailText"), "<Bis>", rs![BearbeitungBis])
            rs("MailText") = Replace(rs("MailText"), "<ReklNr>", txtReklID)

        Case 10 'Ersatzteil Bestellung
    '                If IsNull(rs2("ANE")) Then
    '                    rs("MailVon") = rs2("CCSACHB")
    '                Else
    '                    rs("MailVon") = rs2![ANE]
    '                End If
            'rs("MailVon") = MailVon
            rs("MailAn") = rs2![ANLIEF]
            rs("MailCC") = rs2![CCSACHB]
            'Ermitteln BCC anhand TicketTyp Feld EskalationsAdresse
            rs("MailBCC") = rs3("EskalationsAdresse").Value
            'Betreff erstellen
            rs("MailBetreff") = rs3![MailBetreff]
            rs("MailBetreff") = Replace(rs("MailBetreff"), "<ReklNr>", txtReklID)
            'Body erstellen
            rs("MailText") = rs3![MailText]
            rs("MailText") = Replace(rs("MailText"), "<Bis>", rs![BearbeitungBis])
            rs("MailText") = Replace(rs("MailText"), "<ReklNr>", Me.txtReklID)
            rs("MailText") = Replace(rs("MailText"), "<E-Mail_Absender>", rs("MailVon"))
            rs("MailText") = Replace(rs("MailText"), "<Telefon>", rs2![Telefon])

        Case 13 'Anfrage Lieferschein
            Dim lstrBestellNrtemp As String
            Dim lstrArtNrIDTemp As String
            Dim lstrArtNrTemp As String
            Dim lstrArtNrVarTemp As String

'            txtReklID = 500322

            lstrBestellNrtemp = Nz(DLookup("WEBestellnummer", "ReklMemo", "ID = " & txtReklID))
            lstrArtNrIDTemp = Nz(DLookup("[Artikel-ID 1]", "Mastertab", "ID = " & txtReklID))
            lstrArtNrTemp = Nz(DLookup("[LQ Art-Nr]", "[Artikel-all_avista_out]", "[Artikel-ID] = " & lstrArtNrIDTemp))
            lstrArtNrVarTemp = Nz(DLookup("[Variante]", "[Artikel-all_avista_out]", "[Artikel-ID] = " & lstrArtNrIDTemp))

            rs("MailVon") = MailVon
            rs("MailAn") = rs2![ANLIEF]
            rs("MailCC") = rs2![CCSACHB]
            'Betreff erstellen
            rs("MailBetreff") = rs3![MailBetreff]
            rs("MailBetreff") = Replace(rs("MailBetreff"), "<ReklNr>", txtReklID)
            'Body erstellen
            rs("MailText") = rs3![MailText]
            rs("MailText") = Replace(rs("MailText"), "<Bis>", rs![BearbeitungBis])
            rs("MailText") = Replace(rs("MailText"), "<ReklNr>", Me.txtReklID)
            rs("MailText") = Replace(rs("MailText"), "<E-Mail_Absender>", rs("MailVon"))
            rs("MailText") = Replace(rs("MailText"), "<Telefon>", rs2![Telefon])
            rs("MailText") = Replace(rs("MailText"), "<art>", lstrArtNrTemp)
            rs("MailText") = Replace(rs("MailText"), "<var>", lstrArtNrVarTemp)
            rs("MailText") = Replace(rs("MailText"), "<bestellung>", lstrBestellNrtemp)

        Case 19 'PE-Systech_AV
            rs("MailVon") = rs2![CCSACHB]
            rs("MailCC") = rs2![CCSACHB]
            'Betreff erstellen
            rs("MailBetreff") = rs3![MailBetreff]
            rs("MailBetreff") = Replace(rs("MailBetreff"), "<ReklNr>", Me.txtReklID)
            'Body erstellen
            rs("MailText") = rs3![MailText]
            rs("MailText") = Replace(rs("MailText"), "<ReklNr>", txtReklID)

        Case 20 'Rückweisung Stellungnahme
            'Betreff erstellen
            rs("MailBetreff") = rs3![MailBetreff]
            rs("MailBetreff") = Replace(rs("MailBetreff"), "<ReklNr>", Me.txtReklID)
            rs("MailBetreff") = Replace(rs("MailBetreff"), "<FirmaName>", rs2![Firma_Name]) & " / Masch.Nr.: " & rs2("Maschinennummer")
            'Body erstellen
            rs("MailText") = rs3![MailText]
            rs("MailText") = Replace(rs("MailText"), "<ReklNr>", Me.txtReklID)
            rs("MailText") = Replace(rs("MailText"), "<Bis>", rs![BearbeitungBis])

    End Select

    rs.Update

    'RSP_8.2.3 Requery und Refresh eleminiert
    'Me.Requery
    'Me.Refresh

    Set rs = Nothing
    Set rs2 = Nothing
    Set rs3 = Nothing
    Set rs5 = Nothing


End Function
Private Function getLinkVorgangsordner() As String

Dim sPfadDatei As String


        With Forms("Fehlererfassung - Register")
            If IsNull(.Controls("Mastertab#Meldungsnummer Kunde#2")) Or .Controls("Mastertab#Meldungsnummer Kunde#2") = "" Then
                sPfadDatei = getReklVerzeichnis(.Controls("Mastertab#ID"), format(.Controls("Mastertab#Auftragsdatum#3"), "yyyy"), "")
            Else
                sPfadDatei = getReklVerzeichnis(.Controls("Mastertab#ID"), format(.Controls("Mastertab#Auftragsdatum#3"), "yyyy"), .Controls("Mastertab#Meldungsnummer Kunde#2"))
            End If
        End With

        getLinkVorgangsordner = "<a href=""" & sPfadDatei & """>" & getTC_FCFromCode("TC0482") & "</a></body></html>"

End Function



Public Property Let closeMailSendenAnNeu(fs As String)

    mRefMailSenden.Visible = False
    Set mRefMailSenden = Nothing

End Property
Public Property Let MailSendenAnNeuClosed(fs As String)

    Set mRefMailSenden = Nothing

End Property


Private Sub MessageForm_MessageChanged(Control As String, Message As String)
    On Error GoTo err_MessageForm_MessageChanged

    Dim ls_getAenderung As String
    Dim ls_FieldName As String
    Dim ldb As DAO.Database
    Dim li_TicketID As Long
    Dim rs As Recordset
    Dim DatumBis As Date
    Dim ls_TextHtml As String
    Dim llTicketIDSave As Long
    Dim lrs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    If CurrentProject.AllForms("frmTicketErstellen").IsLoaded = True Then
        DatumBis = Me.txtBearbeitungBis
    Else
        DatumBis = getBearbDateTicketTyp(Me.TicketTypID)
    End If

    Select Case Control
        Case "txtVon"
            'ls_getAenderung = mRefMailSenden.GettxtVON
            ls_getAenderung = Message
            'Me.MailVon = Message
            ls_FieldName = "MailVon"
        Case "txtHTMLText"
            'ls_getAenderung = mRefMailSenden.GettxtVON
            ls_getAenderung = Message
            ls_FieldName = "MailTextHTML"
        Case "txtAn"
            'ls_getAenderung = mRefMailSenden.GettxtAN
            ls_getAenderung = Message
            'Me.MailAn = Message
            ls_FieldName = "MailAn"
        Case "txtCC"
            ls_getAenderung = Message
            'Me.MailCC = Message
            ls_FieldName = "MailCC"
        Case "txtBCC"
            ls_getAenderung = Message
            ls_FieldName = "MailBCC"
        Case "txtBetreff"
            ls_getAenderung = Message
            'Me.MailBetreff = Message
            ls_FieldName = "MailBetreff"
            ChangeSonderzeichenSQLText ls_getAenderung, ls_getAenderung
        Case "txtBody"
            ls_getAenderung = Message
            'Me.MailText = Message
            ls_FieldName = "MailText"
            HtmlTextCreate ls_getAenderung, ls_TextHtml
            ChangeSonderzeichenSQLText ls_getAenderung, ls_getAenderung
        Case "AnhangNeu"
            ls_getAenderung = mRefMailSenden.GetNewAnhang
            ldb.Execute ("Insert into tblTicketAnhaenge ([TicketID], [ReklID], [AnhangPfad]) values (" & plTicketID & ", " & Me.ReklID & ", '" & ls_getAenderung & "');")
            Exit Sub
        Case "DeleteAnhang"
            ls_getAenderung = ""
            '!. Lösche alle Anhänge in Tabelle und hole alle noch bstehenden Anhänge vom MailClient
            ldb.Execute ("Delete * from tblTicketAnhaenge where [TicketID] = " & plTicketID & ";")
            Dim i As Integer
            Dim ls_Anhang As String
            i = 1
            Do
                ls_Anhang = mRefMailSenden.GetAnhangByItem(i)
                If ls_Anhang <> "" Then
                    ldb.Execute ("Insert into tblTicketAnhaenge ([TicketID], [ReklID], [AnhangPfad]) values (" & plTicketID & ", " & Me.ReklID & ", '" & ls_Anhang & "');")
                End If
                i = i + 1
            Loop Until ls_Anhang = ""
            Exit Sub
        Case "SendXDRanLiefMail"
            'Alle Daten zum Mahnwesen eintragen
            ls_getAenderung = 2   ' Versendet
            ls_FieldName = "TicketStatusID"
            'RSP_8.2.3 'Filter aus TicketID aus MailClient via Membervariable des Dialog ersetzt
            Set rs = ldb.OpenRecordset("select * from tblTicketListe where [TicketID] = " & mRefMailSenden.Controls("txtTicketID") & "; ")
            rs.Edit
                rs("MailDatum") = Date
                rs("BearbeitungBis") = DatumBis
                rs("VerlängertBis") = DatumBis
            rs.Update
            Set rs = Nothing
            'RSP_8.2.3 'Filter aus TicketID aus MailClient via Membervariable des Dialog ersetzt
            ldb.Execute ("Update tblTicketListe set [" & ls_FieldName & "] = '" & ls_getAenderung & "' where [TicketID] = " & mRefMailSenden.Controls("txtTicketID") & ";")
            'Mahnwaesen aktivieren
            'RSP_8.2.3 'Finaler Requery nach dem Versenden der Mail, um das Ticket in der Liste anzuzeigen
            If mRefMailSenden.GetTicketNew Then
                Me.Requery
                Me.Recordset.FindFirst "TicketID = " & li_TicketID
            End If
        Case Else
            Exit Sub
    End Select

    If plTicketID = 0 Then
        li_TicketID = Me.TicketID
    Else
        li_TicketID = plTicketID
    End If

    'RSP_8_2_5
    'If ls_getAenderung <> "" Then
    If ls_FieldName = "MailText" Then

        ldb.Execute ("Update tblTicketListe set [" & ls_FieldName & "] = '" & ls_getAenderung & _
                                              "', [MailTextHTML] = '" & ls_TextHtml & "' " & _
                                              "where [TicketID] = " & li_TicketID & ";")
    Else
        ldb.Execute ("Update tblTicketListe set [" & ls_FieldName & "] = '" & ls_getAenderung & "' where [TicketID] = " & li_TicketID & ";")
        '!!!! Requery sehr zeitaufwändig, weil die Abfrage hinter dem Formular sehr komplex ist
        'Forms("frmTicketListe").Requery
    End If
    If Not mRefMailSenden.GetTicketNew Then
        llTicketIDSave = Me.TicketID
        Me.Painting = False
        Me.Requery
'        Set lrs = Me.RecordsetClone
'        lrs.FindFirst "TicketID = " & llTicketIDSave
'        Me.Bookmark = lrs.Bookmark
        Call setBookmark(Me, llTicketIDSave, "TicketID")
        Me.Painting = True
        Set lrs = Nothing
    End If

'    Else
'            ldb.Execute ("Update tblTicketListe set [" & ls_FieldName & "] = '' where [TicketID] = " & li_TicketID & ";")
'    End If
'
'    RSP_8.2.3 'Requery eleminiert
'    Me.Requery
    'Me.Recordset.FindFirst "TicketID = " & li_TicketID
    Exit Sub

err_MessageForm_MessageChanged:
    Resume Next
End Sub

Private Function getBearbDateTicketTyp(fTicketTyp As Long) As Date  '20181022-ROS

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim FirmaID As Integer
    Dim LieferantID As Integer

'    FirmaID = DLookup("[Firma_ID]", "tblFirma", "[Firma_Name] = '" & Forms("Fehlererfassung - Register").Firma_Name & "'")
'    LieferantID = DLookup("[Lieferant-ID]", "Lieferanten", "[Lieferant-ID] = " & Forms("Fehlererfassung - Register").[Lieferanten-Code 1] & "")

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from tblTicketTyp where TicketTypID = " & fTicketTyp & ";")

    Select Case fTicketTyp
        Case 1
'            FirmaID = DLookup("[Firma_ID]", "tblFirma", "[Firma_Name] = '" & Forms("Fehlererfassung - Register").Firma_Name & "'")
            FirmaID = Forms("Fehlererfassung - Register").Controls("Mastertab#Kunden-ID#1") Or Forms("Fehlererfassung - Register").Controls("Mastertab#Kunden-ID#4")
            getBearbDateTicketTyp = getDateAddTage(Date, Nz(DLookup("[ATQS]", "tblReklBearbZeit", "[KundenID] = " & FirmaID & " ")))
            If getBearbDateTicketTyp = Date Then getBearbDateTicketTyp = getDateAddTage(Date, rs("Arbeitstage"))
        Case 2
            LieferantID = DLookup("[Lieferant-ID]", "Lieferanten", "[Lieferant-ID] = " & Forms("Fehlererfassung - Register").Controls("Mastertab#Lieferanten-Code 1") & "")
            getBearbDateTicketTyp = getDateAddTage(Date, Nz(DLookup("[ATQS]", "tblReklBearbZeit", "[LiefID] = " & LieferantID & " ")))
            If getBearbDateTicketTyp = Date Then getBearbDateTicketTyp = getDateAddTage(Date, rs("Arbeitstage"))
        Case Else
            getBearbDateTicketTyp = getDateAddTage(Date, rs("Arbeitstage"))
    End Select

    Set rs = Nothing

 End Function
