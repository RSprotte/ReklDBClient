Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DatumVorVerlängerung As Date

Option Compare Database

Private Sub cmbErledigt_Click()

    Dim lclsMailProv As clsMailProviderOutlook
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim MailText As String
    Dim HTMLMailText As String
    Dim MailBetreff As String
    Dim lo_Coll As Collection
    Dim eMailVon As String


        Set lo_Coll = New Collection
        Set ldb = openAccDBConnection("openDB")
        Set rs = ldb.OpenRecordset("Select * from tblTicketListe where TicketID = " & Me.TicketID & ";", dbOpenSnapshot)

        Select Case Me.TicketTypID      '20181109-ROS
            Case 2, 10, 13
                eMailVon = rs("MailVon")
'            Case 10
'                eMailVon = rs("MailVon")
            Case Else
                eMailVon = rs("MailAn")
        End Select

        If Me.TicketStatusID > 1 Then
            Me.ErledigtAm = Date
            Me.TicketStatusID = 3
            Me.ErledigtVon = get_pKuerzel   '20191009-ROS
            MailBetreff = Me.tblTicketTyp_Anzeige & " " & getTC_FCFromCode("TC0503") & " " & rs("ReklID")
            MailText = MailBetreff & " " & getTC_FCFromCode("TC0504") & _
                    "<br>" & "<br>" & getTC_FCFromCode("TC0505") & " " & get_pKuerzel
        '        sendMailAufgabeAnhang rs("MailAn"), rs("MailVon"), MailBetreff, MailText, lo_Coll, True
            'Instanz auf einen Outlook MailProvider
            Set lclsMailProv = New clsMailProviderOutlook
            lclsMailProv.setBetreff = MailBetreff
            'Mailinhalte = Body füllen
            lclsMailProv.setBody = MailText
            'Mailobjekt füllen, versenden und abspeichern
            lclsMailProv.setAn = rs("MailVon")
        '        lclsMailProv.setAn = "service@de.lq-group.com"
            lclsMailProv.setVon = eMailVon
            lclsMailProv.fillMailItem
            lclsMailProv.SaveAsMailItem get_DateiPfad & "E-Mail\", get_pKuerzel
        '        lclsMailProv.displayMailItem
            lclsMailProv.sendMailItem
            Set lclsMailProv = Nothing

            'sendMailAufgabeAnhang eMailVon, rs("MailVon"), MailBetreff, MailText, lo_Coll, True
        Else
        '        MsgBox ("Ticket ist nicht aktiv!")
            MsgBox getTC_FCFromCode("TC3201"), vbInformation, getTC_FCFromCode("TC3210")
        End If

        Me.Requery
        Me.Refresh

        Set lo_Coll = Nothing
        Set rs = Nothing

End Sub

Private Sub cmbErledigt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Form_DblClick(Cancel As Integer)
    Dim Kriterium As String

    Kriterium = "(TicketID = " & Me.TicketID & ") And (ReklID = " & Me.ReklID & ")"
'    Kriterium = "(TicketID = " & Me.TicketID & ")"
'    Kriterium = "(ReklID = " & Me.ReklID & ")"

'    DoCmd.OpenForm "frmTicketListe", , , Kriterium, , acDialog
'    DoCmd.OpenForm "frmTicketListe", , , Kriterium, , acDialog, "Liste selektiert"
    DoCmd.OpenForm "frmTicketErstellen", , , Kriterium, , acDialog, "Liste selektiert" '20251216-ROS


End Sub

Private Sub VerlängertBis_AfterUpdate()

    Dim strKontakt As String
    Dim strBetreff As String
    Dim strBody As String
    Dim sPfadDatei As String
    Dim strTextHtml As String
    Dim strVon As String
    Dim lDatum As Date          '20190822-ROS_1 wird benötigt um Wert zwischen zu speichern
    Dim lSachbea As String      'Kürzel Sachbearbeiter

    lSachbea = Nz(DLookup("[Sachbearbeiter]", "Mastertab", "[ID] = " & Me.ReklID & ""))

    lDatum = gVerlängertBis

    If Me.VerlängertBis <= lDatum Then
        MsgBox getTC_FCFromCode("TC4036"), vbExclamation, getTC_FCFromCode("TC4037")
        Exit Sub
    Else
        Me.Refresh      '20190904-ROS geändert von Requery auf Refresh
        fct_changedateinTicketlist (Me.TicketID)
    End If

    '20181211-ROS
    If get_ServiceZuordnung = True And getCurrLanguage = "German" Then
        sPfadDatei = get_DateiPfad & "E-Mail\"
        strKontakt = DLookup("MailTOLQEinheit", "LQEinheit", "[ID] = " & Forms("Fehlererfassung - Register").Controls("Mastertab#LQEinheitID#1") & "")

        strBetreff = getTC_FCFromCode("TC0506") & " -" & Me.[tblTicketTyp.Anzeige] & "- " & _
                     getTC_FCFromCode("TC0503") & " " & Forms("Fehlererfassung - Register").Controls("Mastertab#ID")

        strBody = getTC_FCFromCode("TC0424") & "," & _
                  vbNewLine & vbNewLine & _
                  getTC_FCFromCode("TC0507") & "  -" & Me.[tblTicketTyp.Anzeige] & "-  " & getTC_FCFromCode("TC0508") & " " & Me.VerlängertBis & " " & getTC_FCFromCode("TC0509") & _
                  vbNewLine & vbNewLine & _
                  getTC_FCFromCode("TC3214") & " " & getTC_FCFromCode("TC0503") & ": " & lSachbea

        HtmlTextCreate strBody, strTextHtml
        strVon = getTC_FCFromCode("TC0510")        'service@de.lq-group.com
        Call EMailSendenUndSpeichern(strVon, strKontakt, strBetreff, strBody, sPfadDatei)
    End If

End Sub

Private Sub VerlängertBis_GotFocus()

    gVerlängertBis = Me.VerlängertBis

End Sub
