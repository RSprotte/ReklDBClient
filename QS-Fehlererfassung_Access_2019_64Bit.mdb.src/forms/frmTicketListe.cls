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

    Dim ldb As DAO.Database
    Dim Date1 As String
    Dim Schaltjahr As String
    Dim strFormFilter As String
    Dim rsFormFilter As DAO.Recordset
    Dim ld_date As String
    Dim idKuerzel As Integer

    idKuerzel = DLookup("[SB-ID]", "Sachbearbeiter", "[Kurzzeichen] = '" & get_pKuerzel & "'")

    Set ldb = openAccDBConnection("openDB")
    Set rsFormFilter = ldb.OpenRecordset("Select * from tblFilterForm where [FilterBezeichnung] = '" & Me.cboFilterForm & "' And [FormName]= '" & Me.Name & "' And [SachbearbeiterID] = " & idKuerzel & " ;")

    strFormFilter = rsFormFilter("FilterParameter")

    DoCmd.OpenForm "frmTicketListe", acNormal
    Forms("frmTicketListe").Filter = strFormFilter ' & sAuftrag
    Forms("frmTicketListe").FilterOn = True
    Me.Requery
    Me.Refresh


    'DoCmd.Maximize

    Set rsFormFilter = Nothing


End Sub

Private Sub cboFilterLQEinheit_AfterUpdate()

    setFilterTicketliste

End Sub

Private Sub cboPersFilterSpeichern_Click()
    Dim ldb As DAO.Database
    Dim strFilterBezeichnung As String
    Dim strFilterparameter As String

    strFilterparameter = Me.Filter
    strFilterparameter = Replace(strFilterparameter, "'", "''")

    Set ldb = openAccDBConnection("openDB")
    strFilterBezeichnung = getValueFromInputBox(getTC_FCFromCode("IC3845"), getTC_FCFromCode("IC3846"))
    If strFilterBezeichnung = "" Then Exit Sub


    ldb.Execute ("Insert into tblFilterForm ([SachbearbeiterID], [FormName], [FilterParameter], [FilterBezeichnung], [Aktiv], [Dynamisch], [DatumVon], [DatumBis])" & _
                "values(" & get_SachbearbeiterID() & ", '" & Me.Name & "' , '" & Replace(Me.Filter, "'", "''") & "', '" & strFilterBezeichnung & "' , True, True, '','');")


    Me.Requery
    Me.Refresh
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

Private Sub cboTicketTypFilter_AfterUpdate()

     setFilterTicketliste

End Sub

Private Sub cboVorgangStatus_AfterUpdate()

    setFilterTicketliste

End Sub

Private Sub chkAlleOffenenTicket_AfterUpdate()

    setFilterTicketliste

End Sub

Private Sub chkLzAusblenden_AfterUpdate()

    setFilterTicketliste

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

    TicketLoeschen (Me.TicketID)

    Me.Requery
    Me.Refresh

    If IsLoaded("Fehlererfassung - Register") Then
        Forms("Fehlererfassung - Register").Requery
        Forms("Fehlererfassung - Register").Refresh
    End If

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
On Error GoTo Err_cmbFilterLöschen_Click

    Me.cboFilterLQEinheit = ""
    Me.cboTicketTypFilter = ""
    Me.cboTicketStatusFilter = ""
    Me.txtReklID = ""
    Me.cboSachbearbeiter = ""
    Me.cboVorgangStatus = ""
    Me.FilterOn = False
    Me.chkAlleOffenenTicket = 0
    Me.Requery
    Me.Refresh
    Me.cboFilterForm = ""

Exit_cmbFilterLöschen_Click:
    Exit Sub

Err_cmbFilterLöschen_Click:
    MsgBox Err.Description
    Resume Exit_cmbFilterLöschen_Click

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


Private Sub Form_Click()

   If Me.Recordset.RecordCount > 0 Then
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

    Dim objMouseOverForm As clsMouseoverForm
    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With

'    Me.RecordSource = "SELECT tblTicketListe.*, tblTicketStatus.TicketStatusID, tblTicketTyp.Anzeige, tblTicketStatus.Anzeige, LQEinheit.LQEinheit, tblTicketListe.VerlängertBis, " & _
'                      "qryTicketSacharbeiterKz.Kurzzeichen, tblTicketStatus.Anzeige, [Kategorien-Vorgangstatus].Anzeige " & _
'                      "FROM (((Mastertab LEFT JOIN LQEinheit ON Mastertab.LQEinheitID = LQEinheit.ID) " & _
'                      "RIGHT JOIN ((tblTicketListe LEFT JOIN tblTicketStatus ON tblTicketListe.TicketStatusID = tblTicketStatus.TicketStatusID) " & _
'                      "LEFT JOIN tblTicketTyp ON tblTicketListe.TicketTypID = tblTicketTyp.TicketTypID) ON Mastertab.ID = tblTicketListe.ReklID) " & _
'                      "LEFT JOIN qryTicketSacharbeiterKz ON tblTicketListe.ReklID = qryTicketSacharbeiterKz.VorgangID) " & _
'                      "LEFT JOIN [Kategorien-Vorgangstatus] ON Mastertab.[Kategorie-Status] = [Kategorien-Vorgangstatus].[Kategorie-Vorgangstatus] " & _
'                      "ORDER BY tblTicketListe.VerlängertBis;"
    Me.chkAlleOffenenTicket = -1
    Me.chkLzAusblenden = 0
    setFilterTicketliste

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

    Dim ls_Filter As String

    ls_Filter = ""

    If Me.chkAlleOffenenTicket = -1 Then
        ls_Filter = ls_Filter & "(tblTicketListe.TicketStatusID <> 3) AND "
    End If

    If Me.chkLzAusblenden = -1 Then
        ls_Filter = ls_Filter & "(tblTicketListe.MailAn <> 'Service') AND "
    End If

    If Me.txtReklID <> "" Then
        ls_Filter = ls_Filter & "(ReklID = " & Me.txtReklID & ") AND "
    End If

    If Me.cboTicketTypFilter <> "" Then
        ls_Filter = ls_Filter & "(tblTicketTyp.Anzeige = '" & Me.cboTicketTypFilter & "') AND "
'    Else
'        ls_Filter = ls_Filter & " ((tblTicketListe.TicketStatusID <> 3) AND  ((tblTicketListe.TicketTypID)<>14) OR ((tblTicketListe.TicketTypID)=15) OR ((tblTicketListe.TicketTypID)=16) OR ((tblTicketListe.TicketTypID)=17) OR ((tblTicketListe.TicketTypID)=18)) AND "
    End If

    If Me.cboTicketStatusFilter <> "" Then
        ls_Filter = ls_Filter & "(tblTicketStatus.Anzeige = '" & Me.cboTicketStatusFilter & "') AND "
    End If

    If Me.cboSachbearbeiter <> "" Then
        ls_Filter = ls_Filter & "(Kurzzeichen = '" & Me.cboSachbearbeiter & "') AND "
    End If

    If Me.cboSachbearbeiter2 <> "" Then                                                 '20191009-ROS
        ls_Filter = ls_Filter & "(ErledigtVon = '" & Me.cboSachbearbeiter2 & "') AND "
    End If

    If Me.cboVorgangStatus <> "" Then
        ls_Filter = ls_Filter & "([Kategorien-Vorgangstatus].[Anzeige] = '" & Me.cboVorgangStatus & "') AND "
    End If

    If Me.cboFilterLQEinheit <> "" Then
        ls_Filter = ls_Filter & "(LQEinheit = '" & Me.cboFilterLQEinheit & "')"
    End If

    If Right(ls_Filter, 5) = " AND " Then
        Me.Filter = Left(ls_Filter, Len(ls_Filter) - 5)
    Else
        Me.Filter = ls_Filter
    End If


    If Me.Recordset.RecordCount = 0 Then
        MsgBox getTC_FCFromCode("TC4207"), vbInformation, getTC_FCFromCode("TC4208")
        Me.cboFilterLQEinheit = ""
        Me.cboTicketTypFilter = ""
        Me.cboTicketStatusFilter = ""
        Me.txtReklID = ""
        Me.cboSachbearbeiter = ""
        Me.cboVorgangStatus = ""
        Me.FilterOn = False
        Me.chkAlleOffenenTicket = 0
        Me.cboFilterForm = ""
        Me.Refresh
    Else
        Me.FilterOn = True
    End If


End Sub

Private Sub Form_Open(Cancel As Integer)

    If get_AdminZuordnung = True Then
        Me.cmbTicketLöschen.Visible = True
    Else
        Me.cmbTicketLöschen.Visible = False
    End If
End Sub

Private Sub txtReklID_AfterUpdate()

        setFilterTicketliste

End Sub


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
