Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Dim lclsMailProv As clsMailProviderOutlook

Private Sub Befehl7_Click()

    Dim objOL As Outlook.Application
    Dim msg As Outlook.MailItem

    'Instanz auf einen Outlook MailProvider
    Set lclsMailProv = New clsMailProviderOutlook
    lclsMailProv.setfrmParent = Me
    lclsMailProv.OpenItemFromFile ("F:\Qualitätsmanagement\Fehlerberichte_internerGebrauch\2023\77564\E-Mail\20230215_1630-MAF_AW Reklamationsbericht der Firma LQ - Vorgangsnummer 77564.msg")

End Sub


Private Sub cboNameEMail_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
      '** Makro aufrufen
      Call cmbDatenUebernahme_Click
    End If

End Sub

Private Sub cmbCloseFrm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbDatenUebernahme_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Form_Open(Cancel As Integer)


    Dim XPos As Long
    Dim YPos As Long
    Dim lgPosition As Long

    Set lclsMailProv = Nothing

    lgPosition = getPointerPos(XPos, YPos)
    FormOpenKoorMousPointer Me, XPos, YPos, 380, 150

End Sub
Private Sub cmbDatenUebernahme_Click()
On Error GoTo Err_cmbClose_Click

    Dim AntwortEigeneMsgBox As Variant
    Dim Ansprache As String
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim rs_2 As DAO.Recordset
    Dim strABNr As String
    Dim strArt As String
    Dim olApp As Object
    Dim strBody As String
    Dim strBetreff As String
    Dim strEmailAdresse As String
    Dim strCC As String
    Dim sql As String
    Dim sql_2 As String
    Dim strKontaktperson As String
    Dim lReklNr As Long
    Dim lEMailLief As String
    Dim strDatum As String
    Dim strEmailSach As String
    Dim lAnsprache As Boolean
    Dim lAvistaKdNr As String   'ist die AVISTA Kundennummer zB. 01052

    lAnsprache = False

    If CurrentProject.AllForms("Fehlererfassung - Register").IsLoaded = False Then
        lReklNr = InputBox(getTC_FCFromCode("IC4218"), getTC_FCFromCode("IC4219"))
    Else
        lReklNr = Forms("Fehlererfassung - Register").Controls("Mastertab#ID")
    End If

    sql_2 = "SELECT Sachbearbeiter.EmailName " & _
            "FROM Mastertab LEFT JOIN (Lieferanten LEFT JOIN Sachbearbeiter ON Lieferanten.VerantwortlichEinkauf = Sachbearbeiter.Kurzzeichen) " & _
            "ON Mastertab.[Lieferanten-Code 1] = Lieferanten.[Lieferant-ID]" & _
            "WHERE (((Mastertab.ID)= " & lReklNr & ") AND ((Mastertab.[Lieferanten-Code 1]) Is Not Null));"

    sql = "SELECT Mastertab.ID, Mastertab.[Meldungsnummer Kunde], Mastertab.Auftragsnummer, Kunden.EmailAdresse, Mastertab.Auftragsdatum," & _
        "[Baugruppen-avista_out].[LQ Art-Nr], [Artikel-all_avista_out].[LQ Art-Nr]," & _
        "[Artikel-all_avista_out].Variante AS Var, Kunden.Kontaktperson, Mastertab.[Zuständiger Service-MA], tblFirma.Firma_Name " & _
        "FROM (((Mastertab LEFT JOIN [Artikel-all_avista_out] ON Mastertab.[Artikel-ID 1] = [Artikel-all_avista_out].[Artikel-ID]) " & _
        "LEFT JOIN [Baugruppen-avista_out] ON Mastertab.[Baugruppen-ID] = [Baugruppen-avista_out].[Baugruppen-ID]) LEFT JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]) " & _
        "LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID " & _
        "GROUP BY Mastertab.ID, Mastertab.[Meldungsnummer Kunde], Mastertab.Auftragsnummer, Kunden.EmailAdresse, Mastertab.Auftragsdatum, " & _
        "[Baugruppen-avista_out].[LQ Art-Nr] , [Artikel-all_avista_out].[LQ Art-Nr], [Artikel-all_avista_out].Variante, Kunden.Kontaktperson, Mastertab.[Zuständiger Service-MA], tblFirma.Firma_Name " & _
        "HAVING ((([Mastertab.ID])=" & lReklNr & "));"

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset(sql, dbOpenSnapshot)
    Set rs_2 = ldb.OpenRecordset(sql_2, dbOpenSnapshot)

    HtmlTextCreate Nz(DLookup("[Text]", "[tblEmailService]", "[ID] = " & Me.cboNameEMail & " ")), strBody
    strEmailAdresse = Nz(DLookup("[An]", "[tblEmailService]", "[ID] = " & Me.cboNameEMail & ""))
    strBetreff = Nz(DLookup("[Betreff]", "[tblEmailService]", "[ID] = " & Me.cboNameEMail & ""))
    strKontaktperson = Left(rs("Kontaktperson"), InStr(rs("Kontaktperson"), ",") - 1)
    lAvistaKdNr = Nz(DLookup("[KndNr_Avista]", "tblFirma", "[Firma_Name] = '" & rs("Firma_Name") & "'"))

    Select Case Me.cboNameEMail
        Case 15
            'AntwortEigeneMsgBox = EigeneMsgBox("Ansprache auswählen", vbYesNoCancel, "Ansprache")
            'Über die MsgBox wird die Anrede Herr/Frau festgelegt
            AntwortEigeneMsgBox = EigeneMsgBox(getTC_FCFromCode("TC3832"), vbYesNoCancel, getTC_FCFromCode("TC3833"))
            If AntwortEigeneMsgBox = 0 Then Exit Sub
            If AntwortEigeneMsgBox = vbYes Then
                Ansprache = getTC_FCFromCode("TC0498")
            End If
            If AntwortEigeneMsgBox = vbNo Then
                Ansprache = getTC_FCFromCode("TC0499")
            End If
            If AntwortEigeneMsgBox = vbCancel Then
                Ansprache = getTC_FCFromCode("TC0500")
                lAnsprache = True
            End If
            lEMailLief = Nz(DLookup("[MailAn]", "[tblTicketListe]", "[ReklID] = " & rs("ID") & " and [TicketTypID]= 2"))
            strEmailAdresse = Replace(strEmailAdresse, "###MailLief###", lEMailLief)
            strBetreff = Replace(strBetreff, "ReklaNr", rs("ID"))
            strKontaktperson = Left(Forms("Fehlererfassung - Register").Form("frmUFLiefernatDaten").Controls("qryLieferantenDaten#Kontaktperson"), InStr(Forms("Fehlererfassung - Register").Form("frmUFLiefernatDaten").Controls("qryLieferantenDaten#Kontaktperson"), ",") - 1)
            If lAnsprache = True Then
                strBody = Replace(strBody, "Ansprache", Ansprache)
            Else
                strBody = Replace(strBody, "Ansprache", Ansprache & " " & strKontaktperson)
            End If
            strEmailSach = Nz(DLookup("[EmailName]", "qryReklamationSachbearbeiter", "[ID] = " & lReklNr & ""))
            If rs_2.RecordCount Then
                strCC = rs_2("EmailName") & "; " & strEmailSach
            Else
                strCC = strEmailSach
            End If

        Case 6, 7, 8, 9, 10, 11
            'AntwortEigeneMsgBox = EigeneMsgBox("Ansprache auswählen", vbYesNoCancel, "Ansprache")
            'Über die MsgBox wird die Anrede Herr/Frau festgelegt
            AntwortEigeneMsgBox = EigeneMsgBox(getTC_FCFromCode("TC3832"), vbYesNoCancel, getTC_FCFromCode("TC3833"))
            If AntwortEigeneMsgBox = 0 Then Exit Sub
            If AntwortEigeneMsgBox = vbYes Then
                Ansprache = getTC_FCFromCode("TC0498")
            End If
            If AntwortEigeneMsgBox = vbNo Then
                Ansprache = getTC_FCFromCode("TC0499")
            End If
            If AntwortEigeneMsgBox = vbCancel Then
                Ansprache = getTC_FCFromCode("TC0500")
                lAnsprache = True
'                Exit Sub
            End If
            If Not IsNull(rs("Meldungsnummer Kunde")) Then
                strBetreff = Replace(strBetreff, "KDReklaNr", rs("Meldungsnummer Kunde"))
                strBody = Replace(strBody, "KDReklaNr", rs("Meldungsnummer Kunde"))
            Else
                strBetreff = Replace(strBetreff, "KDReklaNr", "vom " & rs("Auftragsdatum"))
            End If
            strBetreff = Replace(strBetreff, "ReklaNr", rs("ID"))
            If lAnsprache = True Then
                strBody = Replace(strBody, "Ansprache", Ansprache)
            Else
                strBody = Replace(strBody, "Ansprache", Ansprache & " " & strKontaktperson)
            End If
'            strBody = Replace(strBody, "xxxx", Forms("Fehlererfassung - Register")![Mastertab#Zuständiger Service-MA])
            strBody = Replace(strBody, "ReklaNr", rs("ID"))
            strBody = Replace(strBody, "YYYY", rs("Auftragsdatum"))
            strDatum = DateAdd("ww", 6, Date)
            strBody = Replace(strBody, "###Datum###", strDatum)
            strEmailAdresse = rs("EmailAdresse")
        Case Else
            If IsNull(rs("Meldungsnummer Kunde")) Then
                strBetreff = Replace(strBetreff, "KDReklaNr", "vom " & rs("Auftragsdatum"))
            Else
                strBetreff = Replace(strBetreff, "KDReklaNr", rs("Meldungsnummer Kunde"))
            End If

            strBetreff = Replace(strBetreff, "ReklaNr", rs("ID"))
            strBetreff = Replace(strBetreff, "###KD###", rs("Firma_Name"))

            If IsNull(rs("Artikel-all_avista_out.LQ Art-Nr")) Then
                strArt = rs("Baugruppen-avista_out.LQ Art-Nr")
            Else
                strArt = rs("Artikel-all_avista_out.LQ Art-Nr") & "-" & rs("Var")
            End If

            If IsNull(rs("Auftragsnummer")) Then
                strABNr = ""
            Else
                strABNr = rs("Auftragsnummer")
'                strBetreff = strBetreff & " / " & strABNr & " / "
                strBetreff = Replace(strBetreff, "#ABNr#", strABNr)
                strBetreff = Replace(strBetreff, "#BestNr#", BestNr_Avista(strABNr))
            End If

            strBody = Replace(strBody, "ArtNr", strArt)
            If IsNull(rs("Meldungsnummer Kunde")) Then
                strBody = Replace(strBody, "KDReklaNr", "vom " & rs("Auftragsdatum"))
            Else
                strBody = Replace(strBody, "KDReklaNr", rs("Meldungsnummer Kunde"))
            End If
            strBody = Replace(strBody, "ReklaNr", rs("ID"))
            strBody = Replace(strBody, "ABXXXX", strABNr)
            strBody = Replace(strBody, "###AVISTA_KD_NR###", lAvistaKdNr)

            If Me.cboNameEMail = 5 Then
                If EigeneMsgBox("Label", vbYesNo, "Label anbringen Ja oder Nein") = vbNo Then
                     strEmailAdresse = Replace(strEmailAdresse, getTC_FCFromCode("TC0512"), "")
                     strBody = Replace(strBody, "###LABEL###", "")
                Else
                    strBody = Replace(strBody, "###LABEL###", getTC_FCFromCode("TC0511"))
                End If
            End If
    End Select

    If IsNull(rs("Zuständiger Service-MA")) Then
        strBody = Replace(strBody, "xxxx", Nz(DLookup("[Name]", "Bearbeiter", "[Kurzzeichen] = '" & get_pKuerzel & "'"), ""))
    Else
        strBody = Replace(strBody, "xxxx", rs("Zuständiger Service-MA"))
    End If


    'Instanz auf einen Outlook MailProvider
    Set lclsMailProv = New clsMailProviderOutlook
    'Set myItem = lclsMailProv.getolMailItem1

    'Mailobjekt füllen und Mail über Outlook öffnen und anzeigen
    lclsMailProv.setVon = getTC_FCFromCode("TC0510")        'service@de.lq-group.com
    lclsMailProv.setAn = strEmailAdresse
    lclsMailProv.setBetreff = strBetreff
    lclsMailProv.setBody = strBody
    lclsMailProv.setCC = strCC
    lclsMailProv.fillMailItem
'    lclsMailProv.displayMailItem

    'Mail wird vor senden angezeigt
    If CurrentProject.AllForms("Fehlererfassung - Register").IsLoaded = False Then
        lclsMailProv.setSaveAs = get_DateiPfad(rs("ID"), rs("Meldungsnummer Kunde"), rs("Auftragsdatum")) & "E-Mail\"
    Else
        lclsMailProv.setSaveAs = get_DateiPfad & "E-Mail\"
    End If
    lclsMailProv.setKuerzel = get_pKuerzel
    lclsMailProv.setfrmParent = Me
    lclsMailProv.displayMailItem

    Set rs = Nothing
    Set rs_2 = Nothing


Exit_cmbClose_Click:
    Exit Sub

Err_cmbClose_Click:
    MsgBox Err.Description
    Resume Exit_cmbClose_Click

End Sub


Private Sub cmbCloseFrm_Click()
On Error GoTo Err_cmbCloseFrm_Click


    If Me.Dirty Then Me.Dirty = False
    DoCmd.Close

Exit_cmbCloseFrm_Click:
    Exit Sub

Err_cmbCloseFrm_Click:
    MsgBox Err.Description
    Resume Exit_cmbCloseFrm_Click

End Sub


Public Sub callBackOutlookProv()
'Call Back Funktion aus dem Outlook Provider
    Set lclsMailProv = Nothing
    DoCmd.Close acForm, Me.Name
End Sub
Public Sub callBackOutlookProvOpenItem()
'Call Back Funktion aus dem Outlook Provider
    Set lclsMailProv = Nothing
    DoCmd.Close acForm, Me.Name
End Sub
