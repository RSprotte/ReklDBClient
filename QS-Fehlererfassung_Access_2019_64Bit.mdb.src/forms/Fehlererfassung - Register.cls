Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Database
'***************************************************
'Änderungen:
'-----------
'16.05.2011:    Beim allgemeinen Drucken wird jetzt vorher das Formular gespeichert, damit alle Daten im Druck erfasst werden.
#If VBA7 Or Win64 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
      (ByVal hwnd As Long, ByVal lpOperation As String, _
      ByVal lpFile As String, ByVal lpParameters As String, _
      ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
      (ByVal hwnd As Long, ByVal lpOperation As String, _
      ByVal lpFile As String, ByVal lpParameters As String, _
      ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If


Private Field_Fehlerklasse As String
Private Field_FehlerartGrob As String
Private FieldBefundLiefVerl As Date
Private FieldQSBearbeitungVerl As Date
Public bArtikelNotInList As Boolean
Private l_Form As Form_DialogZuordnungTagZuRekl

Private gsCurrState As String

Private frm_MailSenden As Form_MailSendenAnNeu

Private ls_RecordSource As String
Private li_ActivePage As Integer
Private ls_AktiverStatus As String

Private Function setRecordSource(ByVal fsRecordSource As String) As Boolean

    setRecordSource = True
    ls_RecordSource = fsRecordSource

End Function
Public Function getRecordSource() As String

    getRecordSource = ls_RecordSource

End Function
Private Function setActivePage(ByVal fiActivePage As Integer) As Boolean

    setActivePage = True
    li_ActivePage = fiActivePage

End Function
Public Function getActivePage() As Integer

    getActivePage = li_ActivePage

End Function

Private Function set_FehlererfassungEnviroment() As String

    Dim lVorgangArtIntern As String
    Dim lInternExtern As Boolean
    Dim i As Integer
    Dim lFaNr As String

        Call RegisterStr246_Change

        lVorgangArtIntern = Nz(DLookup("[VorgangArtIntern]", "[ReklMemo]", "[ID] = " & Me.Controls("Mastertab#ID") & ""))
        lInternExtern = Nz(DLookup("[intern]", "[Mastertab]", "[ID] = " & Me.Controls("Mastertab#ID") & ""))
        lFaNr = Nz(DLookup("[Fertigungsauftrag-FA]", "[Mastertab]", "[ID] = " & Me.Controls("Mastertab#ID") & ""))

        '20190822-ROS_2
        '###########################################################################################################################
        'Flack "VorgangArtIntern" wir hier gesetzt da er momentan in der LQPP-Fehlererfassung nicht gesetzt wird bzw. übergeben wird
        'momentan ist LQPP-Fehlererfassung nur an WorkStations möglich
        '20250205-ROS: geändert da falscher Wert ausgegeben wurde.
        'If Me.Controls("Mastertab#Fertigungsauftrag-FA") <> "" Then lVorgangArtIntern = "FA"
        If lFaNr Like "FA*" Then lVorgangArtIntern = "FA"
        '###########################################################################################################################


        If lInternExtern = True Then
            set_FehlererfassungEnviroment = "Melder-Daten-Intern"
            For i = 0 To Me.Controls("RegisterStr246").Pages.Count - 1
                Me.Controls("RegisterStr246").Pages(i).Visible = False
            Next i
            Me.Controls("RegisterStr246").Pages(0).Visible = True
            Me.Controls("RegisterStr246").Pages(5).Visible = True
            Me.Controls("RegisterStr246").Pages(6).Visible = True
            Me.Controls("RegisterStr246").Pages(8).Visible = True
            Me.Controls("RegisterStr246").Pages(9).Visible = True
            Me.lblRegisterTitel.Caption = Me.Bezeichnungsfeld6407.Caption       '20191007-ROS_1 Eingefügt da Registerbeschriftung falsch war
            Me.lblRegisterTitel.BackColor = Me.Bezeichnungsfeld6407.BackColor   '20191007-ROS_1 Eingefügt da Registerbeschriftung falsch war
            Me.Controls("RegisterStr246").Pages(0).SetFocus
            setEnvirVorgangart lVorgangArtIntern
        Else
            set_FehlererfassungEnviroment = "Melder-Daten 1"
            For i = 0 To Me.Controls("RegisterStr246").Pages.Count - 1
                Me.Controls("RegisterStr246").Pages(i).Visible = True
            Next i
            Me.Controls("RegisterStr246").Pages(1).SetFocus
            Me.Controls("RegisterStr246").Pages(0).Visible = False
        End If

'    Set rs = Nothing

End Function



Private Sub AnzFehlerPos_Click()
Dim la_zahlenpos() As String
Dim Meldungsnummer As String

    '1. Inhalt Dictionary sichern und anschließend leeren
    UpdatetFormTab Me, "[ID]=" & Me.Controls("Mastertab#ID")

    If IsNull(Me.Controls("Mastertab#Referenznummer")) And Me.Controls("Mastertab#Fehlerpositionen") > 1 Then
'            If IsNull(Me.Controls("Mastertab#Meldungsnummer Kunde#2")) Then
'                Meldungsnummer = "LQ" & Me.Controls("Mastertab#ID")
'            Else
'                Meldungsnummer = Me.Controls("Mastertab#Meldungsnummer Kunde#2")
'            End If
'            sReklNrReferenzPos = Meldungsnummer
            sReklNrReferenzPos = Me.Controls("Mastertab#ID")
'            DoCmd.OpenForm "Referenzmeldungspositionen", , , , , acDialog
            DoCmd.OpenForm "Referenzmeldungspositionen"
        ElseIf IsNull(Me.Controls("Mastertab#Referenznummer")) Then
'            MsgBox (" Es sind keine Unterpositionen angelegt!")
            MsgBox getTC_FCFromCode("TC3221"), vbInformation, getTC_FCFromCode("TC3222")
            Exit Sub
        Else
'            sReklNrReferenzPos = Me.Controls("Mastertab#Meldungsnummer Kunde#2")
            sReklNrReferenzPos = Me.Controls("Mastertab#Referenznummer")
            DoCmd.OpenForm "Referenzmeldungspositionen"
    End If

'    DoCmd.RunCommand acCmdSaveRecord

End Sub


Private Sub Befehl6455_Click()

    Set l_Form = New Form_DialogZuordnungTagZuRekl
    l_Form.parentDialog = Me
    l_Form.Visible = True

End Sub

Private Sub Befehl6456_Click()

    ReklasZuArtikel Me.Controls("Mastertab#Artikel-ID 1#2").Column(1)

End Sub

Private Sub Befehl6553_Click()

    If get_ServiceZuordnung = True Then Call TruTicket_Click

End Sub

Private Sub Befehl6559_Click()

    If get_ServiceZuordnung = True Then Call cmb3DReport_Click

End Sub


Private Sub cmbBGUebersicht_Click()

    DoCmd.OpenForm ("frmBGUebersichtBMK")

End Sub

Private Sub cmbDatenKopieren_Click()

    Me.Controls("Mastertab#Bemerkungen - Maßnahmen") = Me.Controls("Mastertab#Welcher Fehler tritt auf")
    Call Mastertab_Bemerkungen___Maßnahmen_AfterUpdate

End Sub

Private Sub cmbDatenUebernehmen_Click()

    Me.Controls("Mastertab#Bemerkungen - Maßnahmen#1") = Me.Controls("Mastertab#Welcher Fehler tritt auf#1")
    Call Mastertab_Bemerkungen___Maßnahmen_1_AfterUpdate

End Sub

Private Sub cmbEmailSendenDispo_Click()

    DoCmd.OpenForm ("frmEmailService")

End Sub

Private Sub LagerübersichtArtikel(ByVal fArtNr As String, ByVal fArtNrVar As String)
'Es wird der Lagerbestand des Artikels bzw Baugruppe abgefragt bund wieder gegeben

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim rsAvistaOut As ADODB.Recordset
    Dim sql As String
    Dim sql_2 As String

    Set adbprov = New clsAvistaProvider
    Set rsAvistaOut = adbprov.getLQMatNrBezeichnung(fArtNr, fArtNrVar)


    sql = "SELECT tt_bestand_sum.besum_ist_menge as Menge ,  vm_lager.lager_bez as Lager, " & _
        "(besum_ist_menge-besum_qs_sperr-besum_we_sperr-besum_res_menge-besum_aint_resmenge-besum_leih_menge) AS Frei, tt_bestand_sum.besum_preis_avg as Lagerpreis, tt_bestand_sum.besum_lagerplatz as Lagerplatz " & _
        "FROM (DBA.tt_bestand_sum INNER JOIN DBA.tm_artikel ON tt_bestand_sum.ast_id = tm_artikel.ast_id) " & _
        "INNER JOIN DBA.ta_preiseinheit ON tm_artikel.pe_id = ta_preiseinheit.pe_id LEFT OUTER JOIN DBA.xm_artikel ON  tt_bestand_sum.ast_id = xm_artikel.ast_id  " & _
        "INNER JOIN DBA.vm_lager ON vm_lager.lager_id = tt_bestand_sum.lager_id " & _
        "LEFT OUTER JOIN DBA.ut_bestand_sum ON tt_bestand_sum.ast_id = ut_bestand_sum.ast_id AND tt_bestand_sum.avar_id = ut_bestand_sum.avar_id AND tt_bestand_sum.lager_id = ut_bestand_sum.lager_id  " & _
        "WHERE tt_bestand_sum.ast_id = '###AstID###' AND tt_bestand_sum.avar_id = '###AstAvar###'And  tt_bestand_sum.besum_ist_menge > 0 ;"

    sql = Replace(sql, "###AstID###", rsAvistaOut("ast_id"))
    sql = Replace(sql, "###AstAvar###", rsAvistaOut("avar_id"))
    sql_2 = Replace(sql, "And  tt_bestand_sum.besum_ist_menge > 0", "")

    Set rsAvista = adbprov.getRecordsetAvista(sql)
    'Daten werden in eine Temp-Tabelle eingetragen und bei öffnen des Formular frmLager_Temp wiedergegeben
    If rsAvista.RecordCount = 0 Then
        Set rsAvista = adbprov.getRecordsetAvista(sql_2)
        MsgBox getTC_FCFromCode("TC4239"), vbInformation, getTC_FCFromCode("TC3216")
    End If
    CurrentDb.Execute ("delete * from tblLagerTemp;")
    rsAvista.MoveFirst
    While Not rsAvista.EOF
        CurrentDb.Execute ("Insert into tblLagerTemp ([Menge],[Lager], [Frei],[Lagerpreis],[Lagerplatz]) values ('" & rsAvista("Menge") & "', '" & rsAvista("Lager") & "', '" & rsAvista("Frei") & "', '" & Round(rsAvista("Lagerpreis"), 2) & "','" & rsAvista("Lagerplatz") & "');")
        rsAvista.MoveNext
    Wend
    DoCmd.OpenForm "frmLager_Temp"

    Set rsAvista = Nothing
    Set rsAvistaOut = Nothing
    Set adbprov = Nothing

End Sub


Private Sub InfoZuFA(ByVal fFA_NR As String)
'Abrufen Informationen, für die Fertigung, zum Fertigungsauftrag

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim ls_sql As String
    Dim str_FAnr As String

    DoCmd.Hourglass True
    str_FAnr = fFA_NR

    ls_sql = "SELECT tt_auftrag_intern_text.aint_memo2 " & _
            "FROM (DBA.tt_auftrag_intern WITH(NOLOCK)" & _
            "LEFT OUTER JOIN DBA.tt_auftrag_intern_text ON tt_auftrag_intern.aint_id = tt_auftrag_intern_text.aint_id)" & _
            "WHERE tt_auftrag_intern.aint_nr = '###FA_Nr###'"

    ls_sql = Replace(ls_sql, "###FA_Nr###", str_FAnr)

    Set adbprov = New clsAvistaProvider
    Set rsAvista = New ADODB.Recordset
    Set rsAvista = adbprov.getRecordsetAvista(ls_sql)

    If Not IsNull(rsAvista("aint_memo2")) Then
        DoCmd.Hourglass False
        MsgBox rsAvista("aint_memo2"), vbInformation, getTC_FCFromCode("TC4204")
    Else
        DoCmd.Hourglass False
        MsgBox getTC_FCFromCode("TC4203"), vbInformation, getTC_FCFromCode("TC4204")
    End If

    Set rsAvista = Nothing
    Set adbprov = Nothing

End Sub

Private Sub cmbGD1_Daten_Abrufen_Click()

    GD1_Art_Daten_Abrufen (Me.Mastertab_Artikel_ID_1_5.Column(1))

End Sub

Private Sub cmbInfoArt_Click()

    If IsNull(Me.Mastertab_Artikel_ID_1_5) Or (Me.Mastertab_Artikel_ID_1_5 = 0) Then
        MsgBox getTC_FCFromCode("TC4186"), vbInformation, getTC_FCFromCode("TC4187")
    Else
        If IsNull(Me.Mastertab_KW_JJ_1) Or (Me.Mastertab_KW_JJ_1 = "") Then

            InfoProduktion Me.Mastertab_Artikel_ID_1_5.Column(1), Me.Artikel_all_avista_out_Variante
        Else
            InfoProduktion , , Left(Right(Me.Mastertab_KW_JJ_1, Len(Me.Mastertab_KW_JJ_1) - InStrRev(Me.Mastertab_KW_JJ_1, "FA") + 1), 16)
        End If
    End If


End Sub

Private Sub cmbInfoBG_Click()
'Ermittlung der Gewichte von ausgelieferten Baugruppen zur Kontrolle auf Vollständigkeit

    Dim ldb As DAO.Database
    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim ls_sql As String
    Dim str_Artikel As String
    Dim ls_Datum As Date


    If IsNull(Me.Mastertab_Baugruppen_ID_3) Then
        DoCmd.OpenForm ("frmInfoBG")
        Exit Sub
    Else
        str_Artikel = Nz(DLookup("[LQ Art-Nr]", "Baugruppen-avista_out", "[Baugruppen-ID]=" & Me.Mastertab_Baugruppen_ID_3))
        If str_Artikel = "" Then Exit Sub
        InfosZuBG "ArtNr", str_Artikel
    End If

End Sub

Private Sub cmbInfoFa_Click()

    Dim strFaNr As String

    strFaNr = InputBox(getTC_FCFromCode("IC4205"), getTC_FCFromCode("IC4206"))
    If strFaNr = "" Then
        Exit Sub
    Else
        InfoProduktion , , strFaNr
    End If

End Sub

Private Sub cmbInfoFaBG_Click()

    If IsNull(Me.Mastertab_Baugruppen_ID_3) Or (Me.Mastertab_Baugruppen_ID_3 = 0) Then
        MsgBox getTC_FCFromCode("TC4188"), vbInformation, getTC_FCFromCode("TC4189")
    Else
        If IsNull(Me.Mastertab_Herstellungsdatum___BG_1) Or (Me.Mastertab_Herstellungsdatum___BG_1 = "") Then
            InfoProduktion Me.Mastertab_Baugruppen_ID_3.Column(1), Me.Baugruppen_avista_out_Var_Nr
        Else
            InfoProduktion , , Left(Right(Me.Mastertab_Herstellungsdatum___BG_1, Len(Me.Mastertab_Herstellungsdatum___BG_1) - InStrRev(Me.Mastertab_Herstellungsdatum___BG_1, "FA") + 1), 16)
        End If
    End If

End Sub

Private Sub cmbKI_Click()

    DoCmd.OpenForm ("frmKI")

End Sub

Private Sub cmbKI_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbKorkenzieher_Click()

    Me.Mastertab_Bemerkungen___Maßnahmen_1 = getTC_FCFromCode("TC0513")

End Sub

Private Sub cmbLagerArt_Click()

    If IsNull(Me.Mastertab_Artikel_ID_1_5) Or (Me.Mastertab_Artikel_ID_1_5 = 0) Then
        MsgBox getTC_FCFromCode("TC4186"), vbInformation, getTC_FCFromCode("TC4187")
    Else
        Call LagerübersichtArtikel(Nz(DLookup("[LQ Art-Nr]", "Artikel-all_avista_out", "[Artikel-ID]=" & Me.Mastertab_Artikel_ID_1_5)), Me.Artikel_all_avista_out_Variante)
    End If

End Sub

Private Sub cmbLagerBG_Click()

    If IsNull(Me.Mastertab_Baugruppen_ID_3) Or (Me.Mastertab_Baugruppen_ID_3 = 0) Then
        MsgBox getTC_FCFromCode("TC4188"), vbInformation, getTC_FCFromCode("TC4189")
    Else
        Call LagerübersichtArtikel(Nz(DLookup("[LQ Art-Nr]", "Baugruppen-avista_out", "[Baugruppen-ID]=" & Me.Mastertab_Baugruppen_ID_3)), Me.Baugruppen_avista_out_Var_Nr)
    End If

End Sub

Private Sub cmbNext_VT_Click()

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim ls_sql As String
    Dim lKDN As String

    ls_sql = "SELECT MIN(tt_abtbpos_termin.ab_tpos_term_versand) as Next_VT " & _
            "FROM ((((DBA.tt_abtbkopf INNER JOIN DBA.tt_auftrag ON tt_abtbkopf.af_id = tt_auftrag.af_id) " & _
            "INNER JOIN DBA.tt_abtbpos ON tt_abtbkopf.ab_id = tt_abtbpos.ab_id) " & _
            "INNER JOIN DBA.tt_abtbpos_termin ON tt_abtbpos.ab_id = tt_abtbpos_termin.ab_id AND tt_abtbpos.ab_pos_id = tt_abtbpos_termin.ab_pos_id) " & _
            "LEFT OUTER JOIN DBA.tt_bestand_sum ON tt_abtbpos.ast_id = tt_bestand_sum.ast_id AND tt_abtbpos.avar_id = tt_bestand_sum.avar_id " & _
            "AND  tt_abtbpos.lager_id = tt_bestand_sum.lager_id) " & _
            "WHERE (ab_tpos_menge - ab_tpos_storno_menge - ab_tpos_wv_menge) > 0 And ab_status < 500 And tt_abtbkopf.ma_id = 2 " & _
            "AND tt_abtbkopf.adk_nummer = '###KDN###' AND tt_abtbpos.ast_nummer <> 'MNK';"

    lKDN = DLookup("[KndNr_Avista]", "qryFirmenName", "[Kunden-ID] = " & Me.Mastertab_Kunden_ID_1)

    ls_sql = Replace(ls_sql, "###KDN###", lKDN)

    Set adbprov = New clsAvistaProvider
    Set rsAvista = New ADODB.Recordset
    Set rsAvista = adbprov.getRecordsetAvista(ls_sql)

    If rsAvista.RecordCount Then
        MsgBox (rsAvista("Next_VT"))
    Else
        MsgBox ("Kein VT")
    End If


    Set rsAvista = Nothing
    Set adbprov = Nothing

End Sub

Private Sub cmbP4T_Click()

    Dim ldb As DAO.Database
    Dim rs, rsKuerzel As DAO.Recordset
    Dim sLastPos As String
    Dim i As Integer
    Dim iPos As Integer
    Dim Vorgang As Integer
    Dim Zeit As String
    Dim lStdSatz As Integer  'Standardstundensatz
    Dim lStdSatzId As Integer   'Standartstundensatz ID

    lStdSatz = Nz(DLookup("Betrag", "tblAufwandStdSaetze", "[Standard] = True"))
    lStdSatzId = Nz(DLookup("StdSatzID", "tblAufwandStdSaetze", "[Standard] = True"))

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from AufwandVorgangPositionen where MID=" & Me.Controls("Mastertab#ID") & " ORDER BY VorgangID;")

    iPos = InputBox(getTC_FCFromCode("IC4246"), getTC_FCFromCode("IC4247"))

    Vorgang = 132   'Übernahme P4T
    If iPos = 1 Then
        Zeit = 0.25
    Else
        Zeit = 0
    End If

    i = 1

    For i = 0 To iPos
        rs.MoveLast
        sLastPos = rs("Pos")
        rs.AddNew
        rs("MID") = Me.Controls("Mastertab#ID")
        rs("Datum") = Date
        rs("Pos") = CStr(CInt(sLastPos) + 10) 'Pos um 10 erhöhen
        rs("Kürzel") = get_pKuerzel
        rs("Zeit") = Zeit
        rs("Vorgang") = Vorgang
        rs("StdSatz") = lStdSatz
        rs("StdSatzID") = lStdSatzId
        rs("DatumStdSatz") = Date
        rs("Aufwand") = Zeit * lStdSatz
        rs.Update
        Vorgang = 80    'Vorgang Abschließen
        Zeit = 0.25
        i = i + 1
    Next i

    Me.UF_Fehlererfassung.SetFocus
    Me.UF_Fehlererfassung.Requery

    Call Forms("Fehlererfassung - Register").Form("UF_Fehlererfassung").Form.Zeit_AfterUpdate

    Set rs = Nothing

End Sub

Private Sub cmbPruefprotokoll_Click()

    Dim lFaNr As String

    If Me.Mastertab_KW_JJ_1 = "" Or IsNull(Me.Mastertab_KW_JJ_1) Then
        lFaNr = InputBox(getTC_FCFromCode("IC4248"), getTC_FCFromCode("IC4249"))
        If lFaNr = "" Then
            Exit Sub
        Else
            Call cmbPruefprotokoll_Oeffnen(lFaNr)
        End If
    Else
        Call cmbPruefprotokoll_Oeffnen(Me.Mastertab_KW_JJ_1)
    End If

End Sub

Private Sub cmbReklZuFA_Click()

    ReklasZuArtikel , Me.Controls("Mastertab#Fertigungsauftrag-FA")

End Sub

Private Sub cmbStueli_Click()

    Dim ldb As DAO.Database
    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim ls_sql As String
    Dim AntwortEigeneMsgBox As Variant
    Dim ls_Datum As String
    Dim lArtNr As String

    ls_sql = cStueli
    ls_Datum = format(Date, "yyyy-mm-dd")
    ls_sql = Replace(ls_sql, "###Datum###", ls_Datum)

    AntwortEigeneMsgBox = EigeneMsgBox(getTC_FCFromCode("TC4222"), vbYesNoCancel, getTC_FCFromCode("TC4221"))

    Select Case AntwortEigeneMsgBox
        Case 6
            If IsNull(Me.Mastertab_Baugruppen_ID_3.Column(1) Or Me.Mastertab_Baugruppen_ID_3.Column(1) = 0) Then
                MsgBox getTC_FCFromCode("TC4188"), vbInformation, getTC_FCFromCode("TC4189")
                Set rsAvista = Nothing
                Set adbprov = Nothing
                Exit Sub
            Else
                ls_sql = Replace(ls_sql, "###AVISTA_KD_ID###", AVISTA_ID_KDnr(Me.Mastertab_ID))
                ls_sql = Replace(ls_sql, "###Art###", Me.Mastertab_Baugruppen_ID_3.Column(1))
                lArtNr = Me.Mastertab_Baugruppen_ID_3.Column(1)
            End If
        Case 7
            If IsNull(Me.Mastertab_Artikel_ID_1_5.Column(1) Or Me.Mastertab_Artikel_ID_1_5.Column(1) = 0) Then
                MsgBox getTC_FCFromCode("TC4186"), vbInformation, getTC_FCFromCode("TC4187")
                Set rsAvista = Nothing
                Set adbprov = Nothing
                Exit Sub
            Else
                ls_sql = Replace(ls_sql, "###AVISTA_KD_ID###", AVISTA_ID_KDnr(Me.Mastertab_ID))
                ls_sql = Replace(ls_sql, "###Art###", Me.Mastertab_Artikel_ID_1_5.Column(1))
                lArtNr = Me.Mastertab_Artikel_ID_1_5.Column(1)
            End If
        Case 2
            Set rsAvista = Nothing
            Set adbprov = Nothing
            Exit Sub
    End Select

    Call Stueli_Auslesen(ls_sql, lArtNr)

End Sub

Private Sub cmbTVNArt_Click()

    If IsNull(Me.Artikel_all_avista_out_Variante) Then
        Call TVN_Abrufen(Me.Mastertab_Artikel_ID_1_5.Column(1))
    Else
        Call TVN_Abrufen(Me.Mastertab_Artikel_ID_1_5.Column(1), Me.Artikel_all_avista_out_Variante)
    End If

End Sub



Private Sub Form_Unload(Cancel As Integer)

    Dim i As Integer

    If Me.Controls("Mastertab#Zustand") = False Then     '20191028-ROS
        If getChangeCountCurrForm(Me) > 0 Then
            i = MsgBox(getTC_FCFromCode("TC3947"), vbYesNoCancel + vbDefaultButton1, getTC_FCFromCode("TC3948"))
            Select Case i
                Case vbYes
                    UpdatetFormTab Me, "[ID]=" & Me.Controls("Mastertab#ID")
                Case vbNo
                Case vbCancel
                    Cancel = True
                    Exit Sub
            End Select
        End If
        FormCloseDictionary Me
        privFormClose
    Else
        UpdatetFormTab Me, "[ID]=" & Me.Controls("Mastertab#ID")
        FormCloseDictionary Me
        privFormClose
    End If


    If IsLoaded("frmFormularROS") Then
        With Forms("frmFormularROS")
            .Form("frmÜbersichtNotizen").Form.Filter = "[Erledigt] = False"
            .Form("frmÜbersichtNotizen").Form.FilterOn = True
            .Refresh
        End With
    End If

    If IsLoaded("frmÜbersichtNotizen") Then
        With Forms("frmÜbersichtNotizen")
            .Filter = "[Erledigt] = False"
            .FilterOn = True
            .Refresh
        End With
    End If

End Sub

Private Sub fraVorgangart_AfterUpdate()
    Dim lrs As DAO.Recordset

    setEnvirVorgangart Me.fraVorgangart

    Set lrs = getCtrlDataset(Me, Me.Controls("ReklMemo#VorgangArtIntern"))
    Set lrs = Nothing

End Sub
Private Sub setEnvirVorgangart(fVorgang As String)
'Prozedur aktiviert in Abhängigkeit der Vorgangart bei der internen Fehlererfassung die Steuerelemente für
'FA = Fertigungsauftragsbezogen
'WE = Wareneingangsbezogen
'WA = Warenausgangsbezogen

    Me.Controls("Mastertab#Auftragsdatum").SetFocus
    Me.Controls("Mastertab#Fertigungsauftrag-FA").Visible = False
    Me.Controls("Mastertab#FALosgroesse").Visible = False
    Me.Controls("ReklMemo#WAWE-AB").Visible = False
    Me.Controls("ReklMemo#WEBestellnummer").Visible = False
    Me.Controls("ReklMemo#WABestellnummer").Visible = False
    Me.Controls("ReklMemo#WELieferschein").Visible = False
    Me.Controls("ReklMemo#WALieferschein").Visible = False
    Me.lblFA.Visible = False
    Me.lblFALosgroesse.Visible = False
    Me.lblWEBestellnummer.Visible = False
    Me.lblWABestellnummer.Visible = False
    Me.lblWELieferschein.Visible = False
    Me.lblWALieferschein.Visible = False

    Select Case fVorgang
        Case 1, "FA"
            Me.Controls("Mastertab#Fertigungsauftrag-FA").Visible = True
            Me.Controls("Mastertab#FALosgroesse").Visible = True
            Me.lblFA.Visible = True
            Me.lblFALosgroesse.Visible = True
            Me.fraVorgangart = 1
            Me.Controls("ReklMemo#VorgangArtIntern") = "FA"

        Case 2, "WE"
            Me.Controls("ReklMemo#WEBestellnummer").Visible = True
            Me.Controls("ReklMemo#WELieferschein").Visible = True
            Me.lblWEBestellnummer.Visible = True
            Me.lblWELieferschein.Visible = True
            Me.fraVorgangart = 2
            Me.Controls("ReklMemo#VorgangArtIntern") = "WE"

        Case 3, "WA"
            Me.Controls("ReklMemo#WABestellnummer").Visible = True
            Me.Controls("ReklMemo#WALieferschein").Visible = True
            Me.Controls("ReklMemo#WAWE-AB").Visible = True
            Me.lblWABestellnummer.Visible = True
            Me.lblWALieferschein.Visible = True
            Me.lblWAWE_AB.Visible = True
            Me.fraVorgangart = 3
            Me.Controls("ReklMemo#VorgangArtIntern") = "WA"
    End Select


End Sub


Private Sub Mastertab_Ablage_Rücklieferschein_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Ablage Rücklieferschein"))

    Set lrs = Nothing

End Sub


Private Sub Mastertab_Artikel_ID_1_2_AfterUpdate()

    Dim lrs As DAO.Recordset

    Me.Controls("Artikel-all_avista_out#Hersteller-Teilenummer").Requery            '20191023-ROS
    Me.Controls("Artikel-all_avista_out#Variante#1").Requery                        '20191023-ROS
    Me.Controls("Mastertab#Artikel-ID 1#2").Requery                                 '20191023-ROS

    Me.Controls("Artikel-all_avista_out#Variante#1") = Me.Controls("Mastertab#Artikel-ID 1#2").Column(2)
    Me.Controls("Artikel-all_avista_out#Hersteller-Teilenummer") = Me.Controls("Mastertab#Artikel-ID 1#2").Column(3)

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Artikel-ID 1#2"))

    Set lrs = Nothing

End Sub

Private Sub Mastertab_Artikel_ID_1_2_DblClick(Cancel As Integer)            '20191023-ROS
    If Not bArtikelNotInList Then
        If Not IsNull(Me.Controls("Mastertab#Artikel-ID 1#2").Column(1)) Then
            sLQArtikelTemp = Me.Controls("Mastertab#Artikel-ID 1#2").Column(1)
            sCurrArtikelSource = "Artikel"
            DoCmd.OpenForm "Artikel-avista_outTemp", , , , , acDialog
            Me.Controls("Mastertab#Artikel-ID 1#2") = lLQArtikelTemp
            Call Mastertab_Artikel_ID_1_2_AfterUpdate
        End If
    End If

    bArtikelNotInList = False

ende:
    Exit Sub
Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0259
    Resume ende
End Sub

Private Sub Mastertab_Artikel_ID_1_2_NotInList(NewData As String, Response As Integer)            '20191023-ROS

'    If Not bArtikelNotInList Then
'        If Not IsNull(Me.Controls("Mastertab#Artikel-ID 1#2").Column(1)) Then
'            sLQArtikelTemp = Me.Controls("Mastertab#Artikel-ID 1#2").Column(1)
'            sCurrArtikelSource = "Artikel"
'            DoCmd.OpenForm "Artikel-avista_outTemp", , , , , acDialog
'            Me.Controls("Mastertab#Artikel-ID 1#2") = lLQArtikelTemp
'            Call Mastertab_Artikel_ID_1_2_AfterUpdate
'        End If
'
'    End If
'
'    bArtikelNotInList = False
'ende:
'    Exit Sub
'Fehler:
'    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0259
'    Resume ende

    '20160510-RSP   '20191113-ROS
    sLQArtikelTemp = NewData
    'BG Zurodnung
    sCurrArtikelSource = "Artikel"
    DoCmd.OpenForm "Artikel-avista_outTemp", , , , , acDialog
    Me.Controls("Mastertab#Artikel-ID 1#2") = lLQArtikelTemp
    Response = acDataErrAdded
    bArtikelNotInList = True

ende:
    Exit Sub

Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0259
    Resume ende


End Sub

Private Sub Mastertab_Artikel_ID_1_4_AfterUpdate()

    Me.Controls("Mastertab#Artikel-ID 1#5") = Me.Controls("Mastertab#Artikel-ID 1#4")
    Call Mastertab_Artikel_ID_1_5_AfterUpdate

'    Dim lBG_ID As Long
'    Dim Response As Integer
'
'    If Not IsNull(Me.Controls("Mastertab#Artikel-ID 1#4")) Then
'        lBG_ID = getTRArtNrAvista(Me.Controls("Mastertab#Artikel-ID 1#4").Column(1))
'        If lBG_ID = 0 Then
'    '        MsgBox ("Artikel " & Me.Artike_1___Kundennummer & " im System nicht vorhanden!")   'TC0258
'            MsgBox (getTC_FCFromCode("TC0258") & " " & Me.Controls("Mastertab#Artikel-ID 1#4") & " " & getTC_FCFromCode("TC0400"))
'        Else
'            'Response = acDataErrContinue
'            Me.Controls("Mastertab#Artikel-ID 1#4") = lBG_ID
'            Response = acDataErrAdded
'        End If
'        Me.Controls("Mastertab#Artikel-ID 1#5") = Me.Controls("Mastertab#Artikel-ID 1#4")
'
'        Call Mastertab_Artikel_ID_1_5_AfterUpdate
'
'    End If

End Sub


Public Function CreateOutlookTask(sdate, Optional EDate, _
    Optional subject = "", Optional body = "", Optional Owner = "")

End Function

Private Sub Auftragsdatum_Exit(Cancel As Integer)

'    MsgBox "Bitte notieren Sie sich die Fehlerbericht-Nummer!", vbInformation, "Wichtig!  Bei Neuerfassung!"  'TC0233
    MsgBox getTC_FCFromCode("TC0233"), vbInformation, getTC_FCFromCode("TC0393")

End Sub


Private Sub Mastertab_Artikel_ID_1_5_DblClick(Cancel As Integer)

    If Me.Mastertab_Kategorie_Status = "Abgeschlossen" Then Exit Sub    '20210428-ROS

    If Not bArtikelNotInList Then
        If Not IsNull(Me.Controls("Mastertab#Artikel-ID 1#5").Column(1)) Then
            sLQArtikelTemp = Me.Controls("Mastertab#Artikel-ID 1#5").Column(1)
            sCurrArtikelSource = "Artikel"
            DoCmd.OpenForm "Artikel-avista_outTemp", , , , , acDialog
            Me.Controls("Mastertab#Artikel-ID 1#5") = lLQArtikelTemp
            Call Mastertab_Artikel_ID_1_5_AfterUpdate
'            getLQArtNrAvista (Me.[Artikel-ID_1].Column(1))
        End If
    End If

    bArtikelNotInList = False

ende:
    Exit Sub
Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0259
    Resume ende
End Sub
Private Sub Mastertab_Artikel_ID_1_5_AfterUpdate()

    Dim lrs As DAO.Recordset

    Me.Controls("Mastertab#Artikel-ID 1#4") = Me.Controls("Mastertab#Artikel-ID 1#5")
    Me.Controls("Artikel-all_avista_out#Variante") = Me.Controls("Mastertab#Artikel-ID 1#5").Column(2)
    If Me.Controls("Mastertab#Artikel-ID 1#5").Column(4) = "" Then
        Me.Controls("txtAvistaArtikelTyp_save") = 0
    Else
        Me.Controls("txtAvistaArtikelTyp_save") = Me.Controls("Mastertab#Artikel-ID 1#5").Column(4)
    End If

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Artikel-ID 1#5"))
    Set lrs = Nothing

End Sub
Private Sub Mastertab_Auftragsdatum_1_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Auftragsdatum#1"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Auftragsdatum_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Auftragsdatum"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Auftragsnummer_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Auftragsnummer"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Auftragsnummer_DblClick(Cancel As Integer)
'20260218-ROS: Abrufen von AB-Daten aus AVISTA

    Call Daten_AB_Avista(Me.Mastertab_Auftragsnummer)

End Sub

Private Sub Mastertab_Aufwand_TRUMPF_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Aufwand-TRUMPF"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Baugruppen_ID_2_AfterUpdate()

    Me.Controls("Mastertab#Baugruppen-ID#3") = Me.Controls("Mastertab#Baugruppen-ID#2")
    Call Mastertab_Baugruppen_ID_3_AfterUpdate

End Sub

Private Sub Mastertab_Baugruppen_ID_3_DblClick(Cancel As Integer)

    If Me.Mastertab_Kategorie_Status = "Abgeschlossen" Then Exit Sub    '20210428-ROS

    If Not bArtikelNotInList Then
        If Not IsNull(Me.Controls("Mastertab#Baugruppen-ID#3").Column(1)) Then
            sLQArtikelTemp = Me.Controls("Mastertab#Baugruppen-ID#3").Column(1)
            sCurrArtikelSource = "BG"
            DoCmd.OpenForm "Artikel-avista_outTemp", , , , , acDialog
            Me.Controls("Mastertab#Baugruppen-ID#3") = lLQArtikelTemp
            Call Mastertab_Baugruppen_ID_3_AfterUpdate
'            getLQArtNrAvista (Me.[Artikel-ID_1].Column(1))
        End If
    End If

    bArtikelNotInList = False
ende:
    Exit Sub
Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0234
    'MsgBox Err.Description, 16, "Fehler"  'TC0234"
    Resume ende


End Sub


Private Sub Mastertab_Baugruppen_ID_AfterUpdate()

    Dim lrs As DAO.Recordset

    Me.Controls("Mastertab#Herstellernummer - BG").Requery                      '20191203-ROS
    Me.Controls("Baugruppen-avista_out#Var Nr#1").Requery                        '20191203-ROS
    Me.Controls("Mastertab#Baugruppen-ID").Requery                                 '20191203-ROS

    Me.Controls("Mastertab#Herstellernummer - BG") = Me.Controls("Mastertab#Baugruppen-ID").Column(3)
    Me.Controls("Baugruppen-avista_out#Var Nr#1") = Me.Controls("Mastertab#Baugruppen-ID").Column(2)

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Baugruppen-ID"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Baugruppen_ID_DblClick(Cancel As Integer) '20191203-ROS
    If Not bArtikelNotInList Then
        If Not IsNull(Me.Controls("Mastertab#Baugruppen-ID").Column(1)) Then
            sLQArtikelTemp = Me.Controls("Mastertab#Baugruppen-ID").Column(1)
            sCurrArtikelSource = "BG"
            DoCmd.OpenForm "Artikel-avista_outTemp", , , , , acDialog
            Me.Controls("Mastertab#Baugruppen-ID") = lLQArtikelTemp
            Call Mastertab_Baugruppen_ID_AfterUpdate
'            getLQArtNrAvista (Me.[Artikel-ID_1].Column(1))
        End If
    End If

    bArtikelNotInList = False
ende:
    Exit Sub
Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0234
    'MsgBox Err.Description, 16, "Fehler"  'TC0234"
    Resume ende

End Sub

Private Sub Mastertab_Baugruppen_ID_NotInList(NewData As String, Response As Integer)   '20191203-ROS
'20150119-RSP:

    sLQArtikelTemp = NewData
    'BG Zurodnung
    sCurrArtikelSource = "BG"
    DoCmd.OpenForm "Artikel-avista_outTemp", , , , , acDialog
    Me.Controls("Mastertab#Baugruppen-ID") = lLQArtikelTemp
    Response = acDataErrAdded
    bArtikelNotInList = True
    Me.Controls("Mastertab#Baugruppen-ID").Requery
    Me.Requery


ende:
    Exit Sub

Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0259
    Resume ende
End Sub

Private Sub Mastertab_Belastungsbetrag_LQ_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Belastungsbetrag LQ"))
    Set lrs = Nothing

    Call PosKostenAnlegen("VerrechnungKD")

End Sub

Private Sub Mastertab_Bemerkungen___Maßnahmen_1_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Bemerkungen - Maßnahmen#1"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Bemerkungen___Maßnahmen_2_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Bemerkungen - Maßnahmen#2"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Bemerkungen___Maßnahmen_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Bemerkungen - Maßnahmen"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Bemerkungen_Vorgang_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Bemerkungen-Vorgang"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Bestellnummer_Kunde_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Bestellnummer Kunde"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Bestellnummer_Kunde_DblClick(Cancel As Integer)
'20260312-ROS: Abrufen von LI-Nr. und FA-Nr. anhand der Kunden Bestellnummer

    If Not IsNull(Me.Mastertab_Bestellnummer_Kunde) Then InfosZuBG "BstNr", , , , Me.Mastertab_Bestellnummer_Kunde

End Sub

Private Sub Mastertab_DefTaetigkeitAnalyseID_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#DefTaetigkeitAnalyseID"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_DefTaetigkeitErfasserID_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#DefTaetigkeitErfasserID"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_DefTaetigkeitID_1_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#DefTaetigkeitID#1"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_DefTaetigkeitID_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#DefTaetigkeitID"))
    Set lrs = Nothing


End Sub

Private Sub Mastertab_Einsatz_vor_Ort_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Einsatz vor Ort"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_ErfasserIntern_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#ErfasserIntern"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_FALosgroesse_AfterUpdate()
    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#FALosgroesse"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Fehlergruppe_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Fehlergruppe"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Fehlertexte_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Fehlertexte"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Fertigungsauftrag_FA_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Fertigungsauftrag-FA"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Fracht_Kosten_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Fracht-Kosten"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Gutschriftsnummer_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Gutschriftsnummer"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Herstellernummer___BG_1_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Herstellernummer - BG#1"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Herstellernummer___BG_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Herstellernummer - BG"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Herstellernummer___UBG_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Herstellernummer - UBG"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Herstellungsdatum___BG_1_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Herstellungsdatum - BG#1"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Herstellungsdatum___BG_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Herstellungsdatum - BG"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Herstellungsdatum___UBG_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Herstellungsdatum - UBG"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Kategorie_Bearbeitung_I_AfterUpdate()

    Dim lrs As DAO.Recordset
    Me.Controls("Mastertab#Kategorie-Bearbeitung II") = -1

    Call Mastertab_Kategorie_Bearbeitung_II_AfterUpdate

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Kategorie-Bearbeitung I"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Kategorie_Bearbeitung_II_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Kategorie-Bearbeitung II"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Kategorie_Problem_3_AfterUpdate()

    Dim sMemo As String
    Dim lrs As DAO.Recordset

    sMemo = Chr(13) & Chr(10) & "Neuer Status:  " & Me.Controls("Mastertab#Kategorie-Problem#3")
    sMemo = InsertMemoDS(Me.Controls("Mastertab#ID"), sMemo)


    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Kategorie-Problem#3"))
    Set lrs = Nothing


End Sub

Private Sub Mastertab_Kategorie_Problem_4_AfterUpdate()

    Dim sMemo As String
    Dim lrs As DAO.Recordset

    sMemo = Chr(13) & Chr(10) & "Neuer Status:  " & Me.Controls("Mastertab#Kategorie-Problem#4")
    sMemo = InsertMemoDS(Me.Controls("Mastertab#ID"), sMemo)


    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Kategorie-Problem#4"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Kategorie_Problem_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Kategorie-Problem"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Kategorie_Status_AfterUpdate()

    Dim lrs As DAO.Recordset
    Dim lKosten As String

    lKosten = Nz(DLookup("[KostenaufstellungID]", "tblKostenaufstellung", "[ReklID] = " & Me.Controls("Mastertab#ID") & ""), "")

        If (Me.Mastertab_Kategorie_Status = "Abgeschlossen") And (Me.Mastertab_extern = True) Then
            If lKosten <> "" Then
                MsgBox getTC_FCFromCode("TC4257"), vbInformation, getTC_FCFromCode("TC4258")

                Call Status_AfterUpdate(ls_AktiverStatus)      '20190830-ROS_2 Protokolierung der Zustansänderung wird eingetragen

                Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Kategorie-Status"))
                Set lrs = Nothing
                Me.RegisterStr246.Pages(9).SetFocus
            Else
                Call Status_AfterUpdate(ls_AktiverStatus)      '20190830-ROS_2 Protokolierung der Zustansänderung wird eingetragen

                Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Kategorie-Status"))
                Set lrs = Nothing
            End If
        Else
            Call Status_AfterUpdate(ls_AktiverStatus)      '20190830-ROS_2 Protokolierung der Zustansänderung wird eingetragen

            Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Kategorie-Status"))
            Set lrs = Nothing
        End If

End Sub

Private Sub Mastertab_KW_JJ_1_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#KW/JJ#1"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_KW_JJ_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#KW/JJ"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_langfristige_Maßnahmen_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#langfristige Maßnahmen"))

    Set lrs = Nothing

End Sub

Private Sub Mastertab_LiefAnsprechp_ID_3_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Lieferanten-Code 1"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Lieferscheinnummer_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Lieferscheinnummer"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_LINK_Nr___UBG_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#LINK-Nr - UBG"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_LQBereichID_AfterUpdate()

    Dim lrs As DAO.Recordset

    Me.Controls("Mastertab#LQWorkstationID") = -1
    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#LQBereichID"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_LQEinheitID_AfterUpdate()

    Dim lrs As DAO.Recordset

    Me.Controls("Mastertab#LQBereichID") = -1
    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#LQEinheitID"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_LQWorkstationID_AfterUpdate()

    Dim lrs As DAO.Recordset

    Me.Controls("Mastertab#TaetigkeitID") = -1
    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#LQWorkstationID"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Maschinennummer_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Maschinennummer"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Material_Kosten_AfterUpdate()

    Dim lrs As DAO.Recordset

    If IsNull(Me.Controls("Mastertab#Material-Kosten")) Then Me.Controls("Mastertab#Material-Kosten") = "0,00"

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Material-Kosten"))
    Set lrs = Nothing

    Call PosKostenAnlegen("MatKosten")

End Sub

Private Sub Mastertab_Meldungsnummer_Kunde_AfterUpdate()

    Dim lrs As DAO.Recordset
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from Mastertab where [Meldungsnummer Kunde] = '" & Me.Controls("Mastertab#Meldungsnummer Kunde") & "' ;")

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        MsgBox getTC_FCFromCode("TC0247") & " " & getTC_FCFromCode("TC0246") & " " & rs("ID"), vbExclamation, getTC_FCFromCode("TC0408")
    End If

    Set rs = Nothing

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Meldungsnummer Kunde"))
    Set lrs = Nothing

    '20200210-ROS soll verhindern das der Vorgangsordner doppelt angelet wird.
    Call VorgangsOrdner(Me.Controls("Mastertab#ID"), Me.Controls("Mastertab#Meldungsnummer Kunde"), Me.Controls("Mastertab#Auftragsdatum#3"), False)

End Sub


Private Sub Mastertab_Menge_betroffen___Kunde_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Menge betroffen - Kunde"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Menge_betroffen___LQ_1_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Menge betroffen - LQ#1"))
    Set lrs = Nothing


End Sub

Private Sub Mastertab_Menge_betroffen___LQ_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Menge betroffen - LQ"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_mit_Rücksendung_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#mit Rücksendung"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Mitarbeiter_Kurzzeichen_1_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Mitarbeiter-Kurzzeichen#1"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Mitarbeiter_Kurzzeichen_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Mitarbeiter-Kurzzeichen"))
    Set lrs = Nothing

End Sub



Private Sub Mastertab_per_E_Mail_Fax_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#per E-Mail/Fax"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_per_Telefon_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#per Telefon"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Prüfkosten_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Prüfkosten"))
    Set lrs = Nothing


End Sub



Private Sub Mastertab_Rechnungsnummer_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Rechnungsnummer"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Reparaturbetrag_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Reparaturbetrag"))
    Set lrs = Nothing

    Call PosKostenAnlegen("RepKosten")

End Sub

Private Sub Mastertab_Rückliefer_Nummer_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Rückliefer-Nummer"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Rückmeldung_Kunde_Ist_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Rückmeldung_Kunde_Ist"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Rückmeldung_Kunde_Verschiebung_1_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Rückmeldung_Kunde_Verschiebung#1"))
    Set lrs = Nothing

End Sub


Private Sub Mastertab_Rückmeldung_Kunde_Verschiebung_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Rückmeldung_Kunde_Verschiebung"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Sachbearbeiter_1_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Sachbearbeiter#1"))

    Set lrs = Nothing

End Sub

Private Sub Mastertab_Sachbearbeiter_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Sachbearbeiter"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Serial_Nr___BG_1_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Serial-Nr - BG#1"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Serial_Nr___BG_1_DblClick(Cancel As Integer)

    If Not IsNull(Me.Mastertab_Serial_Nr___BG_1) Then
        Call Info_SNR(Me.Mastertab_Serial_Nr___BG_1)
    End If

End Sub

Private Sub Mastertab_Serial_Nr___BG_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Serial-Nr - BG"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Serial_Nr___UBG_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Serial-Nr - UBG"))
    Set lrs = Nothing


End Sub

Private Sub Mastertab_Serial_Nr___UBG_DblClick(Cancel As Integer)

    If Not IsNull(Me.Mastertab_Serial_Nr___UBG) Then
        Call Info_SNR(Me.Mastertab_Serial_Nr___UBG)
    End If

End Sub

Private Sub Mastertab_Seriennummer_1_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Seriennummer#1"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Seriennummer_1_DblClick(Cancel As Integer)

    If Not IsNull(Me.Mastertab_Seriennummer_1) Then
        Call Info_SNR(Me.Mastertab_Seriennummer_1)
    End If

End Sub

Private Sub Mastertab_Seriennummer_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Seriennummer"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_TaetigkeitID_1_AfterUpdate()

    Dim lrs As DAO.Recordset
    'Abhängige Werte setzen
    Me.Controls("Mastertab#DefTaetigkeitID#1") = -1


    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#TaetigkeitID#1"))


    Set lrs = Nothing


End Sub


Private Sub Mastertab_TaetigkeitID_1_LostFocus()

    If IsNull(Me.Controls("Mastertab#DefTaetigkeitID#1")) Then
        Field_FehlerartGrob = 0
    Else
        Field_FehlerartGrob = Me.Controls("Mastertab#DefTaetigkeitID#1")
    End If


End Sub

Private Sub Mastertab_TaetigkeitErfasserID_AfterUpdate()

    Dim lrs As DAO.Recordset


    Me.Controls("Mastertab#DefTaetigkeitErfasserID") = -1
    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#TaetigkeitErfasserID"))
    Set lrs = Nothing

'    Me.Controls("Mastertab#DefTaetigkeitErfasserID") = -1
'    Me.Controls("Mastertab#DefTaetigkeitErfasserID").Requery
End Sub




Private Sub Mastertab_Baugruppen_ID_3_AfterUpdate()

    Dim lrs As DAO.Recordset

'    Me.Controls("Mastertab#Baugruppen-ID#4") = Me.Controls("Mastertab#Baugruppen-ID#3").Column(2)
    Me.Controls("Baugruppen-avista_out#Var Nr") = Me.Controls("Mastertab#Baugruppen-ID#3").Column(2)
    Me.Controls("Mastertab#Baugruppen-ID#2") = Me.Controls("Mastertab#Baugruppen-ID#3")

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Baugruppen-ID#3"))
    Set lrs = Nothing


End Sub



Private Sub Befehl6068_Click()
On Error GoTo err_Befehl6068_Click

    '1. Inhalt Dictionary sichern und anschließend leeren
    UpdatetFormTab Me, "[ID]=" & Me.Controls("Mastertab#ID")

    'DoCmd.Save acForm, "Fehlererfassung - Register"
'    DoCmd.RunCommand acCmdSaveRecord

    Exit Sub

err_Befehl6068_Click:
    If Err = 2046 Then
        MsgBox getTC_FCFromCode("TC0235")
        'MsgBox "Im Zustand 'Abgeschlossen' ist ein Speichern nicht möglich!"   'TC0235
        Exit Sub
    Else
        MsgBox "Error: " & Err & " - " & Error   'TC0236
        Resume Next
    End If


End Sub


Private Sub Befehl6084_Click()

    Dim ldb As DAO.Database
    Dim sSQL As String
    Dim lFilter As String

    If Me.Controls("txtTaetigkeit").Value = "" Or Me.Controls("Text6146") = "" Then
'        MsgBox ("Es sind keine Tätigkeiten und keine Fehlerbilder ausgewählt!")
        MsgBox getTC_FCFromCode("TC3223"), vbExclamation, getTC_FCFromCode("TC3224")
        Exit Sub
    End If

    If IsNull(Me.Controls("ReklMemo#FertigungsDatum#1")) Then
'        MsgBox ("Fertigungs-Datum fehlt!")
        MsgBox getTC_FCFromCode("TC3243"), vbInformation, getTC_FCFromCode("TC3220")
        Exit Sub
    End If



    psMassnahmeBearbeitung = "Ohne XD-Report"

    Set ldb = openAccDBConnection("openDB")

    ldb.Execute ("Delete * from MaßnahmenTempDruck;")
'    sSQL = "INSERT INTO MaßnahmenTempDruck SELECT Maßnahmen.* FROM Maßnahmen INNER JOIN MaßnahmenFehlerarten ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerarten.MaßnahmeID " & _
'           "WHERE (MaßnahmenFehlerarten.Fehlerklasse=(Select Taetigkeit from Taetigkeiten where Taetigkeiten.ID = " & Me.[TaetigkeitID] & ") AND Maßnahmen.StartUmsetzung is not null and " & _
'                "(Maßnahmen.EndeUmsetzung > #" & Format(Me.Auftragsdatum, "M\/D\/YYYY") & "# or Maßnahmen.EndeUmsetzung is null));"
    sSQL = "INSERT INTO MaßnahmenTempDruck SELECT Maßnahmen.* FROM Maßnahmen INNER JOIN MaßnahmenFehlerarten ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerarten.MaßnahmeID " & _
       "WHERE (MaßnahmenFehlerarten.Fehlerklasse=(Select Taetigkeit from Taetigkeiten where Taetigkeiten.ID = " & Me.Controls("Mastertab#TaetigkeitID#2") & ") AND Maßnahmen.StartUmsetzung is not null);"


    ldb.Execute (sSQL)
     ' Set ldb = openAccDBConnection("closeDB")
    DoCmd.OpenForm "MaßnahmenTempDruck"

End Sub


Private Sub Befehl6182_Click() '20160302-ROS
    If IsNull(Me.Controls("Mastertab#Zuständiger Service-MA")) Then
        MahnwesenKunde Me.Controls("Mastertab#ID"), ""
    Else
        MahnwesenKunde Me.Controls("Mastertab#ID"), Me.Controls("Mastertab#Zuständiger Service-MA")
    End If
    Me.frmMahnwesenKunde.Requery
    Me.[frmMahnwesenKunde].SetFocus
    DoCmd.GoToRecord , , acLast
End Sub


Private Sub Befehl6263_Click()

    Dim stDocName1 As String

    stDocName1 = "KundenUebersicht"

    If IsNull(Me.Controls("Mastertab#Kunden-ID#1")) Then        '20190828_ROS_1
            DoCmd.OpenForm stDocName1
        Else
            DoCmd.OpenForm stDocName1, , , "KndNr_Avista='" & Me.frmUFKundenDaten.Controls("qryKundenDaten#KndNr_Avista") & "'"
    End If

    DoCmd.SelectObject acForm, Me.Name, False
    DoCmd.SelectObject acForm, stDocName1, False
    Forms(stDocName1).Controls("FilterEingabe").SetFocus

End Sub

Private Sub BefehlAlleReklArtikel_Click()

    If IsNull(Me.Controls("Mastertab#Artikel-ID 1#5").Column(1)) Then
        MsgBox getTC_FCFromCode("TC4186"), vbInformation, getTC_FCFromCode("TC4187")
'        ReklasZuArtikel Me.Controls("Mastertab#Baugruppen-ID#3").Column(1)
    Else
        ReklasZuArtikel Me.Controls("Mastertab#Artikel-ID 1#5").Column(1)
    End If

End Sub

Private Sub Mastertab_Belastungsbetrag_Kunde_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Belastungsbetrag Kunde"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Bemerkungen_Vorgang_DblClick(Cancel As Integer)

'    Me.RegisterStr246.Pages(4).SetFocus
    DoCmd.OpenForm "FehlerRecherchePE_PopUp", acNormal, , , , acDialog
    Me.Requery

End Sub

Private Sub Mastertab_Bemerkungen_Vorgang_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub



Private Sub ButDatenUebernehmen_Click()

    Me.Controls("Mastertab#Bemerkungen - Maßnahmen#1") = Me.Controls("Mastertab#Welcher Fehler tritt auf#1")
    Call Mastertab_Bemerkungen___Maßnahmen_1_AfterUpdate

End Sub


Private Sub ButNewTag_Click()

    Set l_Form = New Form_DialogZuordnungTagZuRekl
    l_Form.parentDialog = Me
    l_Form.Visible = True

End Sub

Private Sub cboRibbonOff_Click()
        Select Case fOSUserName
            Case "ralf sprotte", "roberto santarelli"
                DoCmd.ShowToolbar "Ribbon", acToolbarNo
        End Select

End Sub


Private Sub cboRibbonOn_Click()

Dim ldb As DAO.Database
Dim rs As DAO.Recordset

If fOSUserName = "ralf sprotte" Or fOSUserName = "roberto santarelli" Then
    'Init Ribbon-Button
    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select [Value] from KonfigurationReklDBLocal where Konfiguration = 'RibbonOnOff';")
    If rs("Value") = "An" Then
        Me.cboRibbonOn.Caption = "Ribbon on"
        DoCmd.ShowToolbar "Ribbon", acToolbarNo
        ldb.Execute ("Update KonfigurationReklDBLocal set [Value] = 'Aus' where [Konfiguration] = 'RibbonOnOff';")
    Else
        DoCmd.ShowToolbar "Ribbon", acToolbarYes
        Me.cboRibbonOn.Caption = "Ribbon off"
        ldb.Execute ("Update KonfigurationReklDBLocal set [Value] = 'An' where [Konfiguration] = 'RibbonOnOff';")
    End If
End If

End Sub


Public Sub cmb2MahungVersenden_Click()

    Dim olApp As Object
    Dim strBody As String
    Dim strSachbeaEmail As String
    Dim AntwortEigeneMsgBox As Variant
    Dim Ansprache As String
    Dim strKontaktperson As String

    strSachbeaEmail = Nz(DLookup("[EmailName]", "Sachbearbeiter", "[Name] = '" & Me.Controls("Mastertab#Sachbearbeiter#1") & "'"), "")

'        AntwortEigeneMsgBox = EigeneMsgBox("Ansprache auswählen", vbYesNoCancel, "Ansprache")
        AntwortEigeneMsgBox = EigeneMsgBox(getTC_FCFromCode("TC3832"), vbYesNoCancel, getTC_FCFromCode("TC3833"))
        If AntwortEigeneMsgBox = vbYes Then
            Ansprache = getTC_FCFromCode("TC0498")
        End If
        If AntwortEigeneMsgBox = vbNo Then
            Ansprache = getTC_FCFromCode("TC0499")
        End If
        If AntwortEigeneMsgBox = vbCancel Then
            Exit Sub
        End If

        strKontaktperson = Left(Me.frmUFLiefernatDaten.Controls("qryLieferantenDaten#Kontaktperson"), InStr(Me.frmUFLiefernatDaten.Controls("qryLieferantenDaten#Kontaktperson"), ",") - 1)

    strBody = "<span style=""font-size:10pt; font-family:'Arial'"">" & _
            Ansprache & " " & strKontaktperson & "," & " <br><br>" & _
            "wir haben zur o.g. Reklamation, auch nach der 2.Mahnung, noch keine Stellungnahme erhalten. <br>" & _
            "Bitte teilen sie uns einen Termin mit, bis wann wir mit einer Stellungnahme rechnen können. <br>" & _
            "Wir werden dann den Termin dementsprechend anpassen. <br><br><br>" & _
            "Vielen Dank im Voraus. <br><br><br><br>" & _
            "Für Rückfragen stehen wir Ihnen gerne zur Verfügung.<br><br>" & _
            "Mit freundlichen Grüßen <br>" & _
            "i.A. " & "Roberto Santarelli <br><br>" & _
            "Mitarbeiter Service <br>" & _
            "Roberto.Santarelli@de.lq-group.com _Telefon +49.7143.968-36959 <br><br>" & _
            "LQ Mechatronik-Systeme GmbH<br>" & _
            "Carl-Benz-Straße 6 _D-74354 Besigheim<br>" & _
            "Fax +49.7143.9683-96 _www.lq-group.com<br><br>" & _
            "<span style=""font-size:7pt; font-family:'Arial'"">" & _
            "LQ Mechatronik-Systeme GmbH, Sitz: Freiberg am Neckar<br>" & _
            "Registergericht: Amtsgericht Stuttgart - HRB-Nr. 205651<br>" & _
            "Geschäftsführer: Friedrich P. Link, José Quero<br><br>" & _
            "Wichtiger Hinweis: Diese E-Mail und etwaige Anlagen enthalten<br>" & _
            "firmenvertrauliche Informationen. Sollten Sie diese E-Mail<br>" & _
            "irrtümlich erhalten haben, benachrichtigen Sie uns bitte durch<br>" & _
            "eine Antwort-Mail und löschen Sie diese E-Mail nebst Anlagen<br>" & _
            "von Ihrem System. Vielen Dank."

    Set olApp = CreateObject("Outlook.Application")
    With olApp.CreateItem(0)
        .SentOnBehalfOfName = "Roberto.Santarelli@de.lq-group.com"
        .To = Me.frmUFLiefernatDaten.Controls("AnsprechpartnerID#Email").Value
        .CC = strSachbeaEmail
        .subject = "Termin Stellungnahme Reklamation " & Me.Controls("Mastertab#ID").Value
        .HTMLBody = strBody
        .Display
    End With
    Set olApp = Nothing

End Sub


Private Sub cmb3DReport_Click()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim rs2 As DAO.Recordset
    Dim fs As Object
    Dim wb As Excel.Workbook
    Dim WS As Excel.Worksheet
    Dim objExcel As Object
    Dim lexcelprov As clsExcelProvider
    Dim bReportOeffnen As Boolean
    Dim sPfadDatei As String
    Dim intI As Integer
    Dim xlApp As Excel.Application
    Dim Fehlertext As String
    Dim FehlerUrsache As String
    Dim Maßnahmen As String
    Dim Sachbearbeiter As String
    Dim XDReportTemplate As String  '20160304-ROS
    Dim XDReportPfad As String  '20160304-ROS
    Dim AntwortEigeneMsgBox As Variant

    '1. Inhalt Dictionary sichern und anschließend leeren
    UpdatetFormTab Me, "[ID]=" & Me.Controls("Mastertab#ID")

'    If Not IsNull(Me.Controls("Mastertab#Referenznummer")) Then   '20190109-ROS
    If Me.Controls("Mastertab#Referenznummer") <> "" Then '20190109-ROS
'        MsgBox ("Zu Unterpositionen kann kein 3D-Report erstellt werden!")
        MsgBox getTC_FCFromCode("TC3225"), vbExclamation, getTC_FCFromCode("TC3226")
        Exit Sub
    End If

'    AntwortEigeneMsgBox = EigeneMsgBox("Bitte Sprache auswählen", vbYesNoCancel, "3D-Report")
    AntwortEigeneMsgBox = EigeneMsgBox(getTC_FCFromCode("TC3828"), vbYesNoCancel, getTC_FCFromCode("TC3829"))
        If AntwortEigeneMsgBox = vbYes Then
            XDReportTemplate = getTC_FCFromCode("DP0002") 'Template Deutsch
        End If
        If AntwortEigeneMsgBox = vbNo Then
            XDReportTemplate = getTC_FCFromCode("DP0007") 'Template Englisch
        End If
        If AntwortEigeneMsgBox = vbCancel Then
            Exit Sub
        End If

    DoCmd.Hourglass True

    XDReportPfad = getTC_FCFromCode("DP0001")

    'Init
    bReportOeffnen = False

    'Erzeugen einer Excel Instanz
    Set lexcelprov = New clsExcelProvider
    Set xlApp = lexcelprov.getExcelInstance()
    'lexcelprov.OpenKillExcelInstance

    'Prüfen ob XD-Report schon existiert
    Set fs = CreateObject("Scripting.FileSystemObject")
    'ist Verzeichnis vorhanden
    'sPfadDatei = getReklVerzeichnis(Me.Controls("Mastertab#ID"), Format(Me.Auftragsdatum, "yyyy"), Me.[Meldungsnummer Kunde])
    If (IsNull(Me.Controls("Mastertab#Meldungsnummer Kunde#2")) Or Me.Controls("Mastertab#Meldungsnummer Kunde#2") = "") Then
        sPfadDatei = setReklVerzeichnis(Me.Controls("Mastertab#ID"), format(Me.Controls("Mastertab#Auftragsdatum#3"), "yyyy"), "")
    Else
        sPfadDatei = setReklVerzeichnis(Me.Controls("Mastertab#ID"), format(Me.Controls("Mastertab#Auftragsdatum#3"), "yyyy"), Me.Controls("Mastertab#Meldungsnummer Kunde#2"))
    End If
    If sPfadDatei <> "" Then
        If fs.FileExists(sPfadDatei & "\" & Me.Controls("Mastertab#ID") & "_" & XDReportTemplate) = False Then
            '1. Kopie von Template erstellen
            Set wb = lexcelprov.openExcelInstance(XDReportPfad & XDReportTemplate, sPfadDatei & "\" & Me.Controls("Mastertab#ID") & "_" & XDReportTemplate)
            wb.Activate
            '2. XD-Report generieren
            Set WS = wb.Worksheets("3D-Report")
            WS.Activate
            Set ldb = openAccDBConnection("openDB")
            Set rs = ldb.OpenRecordset("Select * from [ROS_3D-Report] where [ID] = " & Me.Controls("Mastertab#ID") & ";", dbOpenSnapshot)
            If rs.RecordCount > 0 Then
                WS.Range("reklNr") = rs![ID]            'Feldnamen von Abfrage
                WS.Range("Firma") = rs![Firma]
                WS.Range("Kontaktperson") = rs![Kontaktperson]
                WS.Range("ReklNr_KD") = rs![ReklNr_KD]
                WS.Range("Auftragsdatum") = rs![Auftragsdatum]
                WS.Range("Telefonnummer") = rs![Telefonnummer]
                WS.Range("EmailAdresse") = rs![EmailAdresse]
                WS.Range("Maschinennummer") = rs![Maschinennummer]
                WS.Range("FehlermengeKD") = rs![FehlermengeKD]
                WS.Range("BestNrKD") = rs![BestNrKD]
                WS.Range("BG_KD") = rs![BG_KD]
                WS.Range("BG_LQ") = rs![BG_LQ]
                WS.Range("Art_KD") = rs![Art_KD]
                WS.Range("Art_LQ") = rs![Art_LQ]
                WS.Range("Variante") = rs![Variante]
                WS.Range("Service_MA") = rs![Service_MA]
                If rs![Kategorie-Problem] = "berechtigt" Then WS.Range("Berechtigt") = "X"
                If rs![Kategorie-Problem] = "nicht berechtigt" Then WS.Range("nicht_Berechtigt") = "X"

                    If IsNull(Me.Controls("Mastertab#Referenznummer")) Then
                        Set rs2 = ldb.OpenRecordset("Select * from [ROS_3D-Report] where [Referenznummer]  = '" & Me.Controls("Mastertab#ID") & "';", dbOpenSnapshot)
                            If rs2.RecordCount Then
                                rs2.MoveLast
                                rs2.MoveFirst
                                For intI = 1 To rs2.RecordCount
                                    Fehlertext = "Pos." & rs2![Fehlerpositionen] & ":" & vbCrLf & rs2![Fehler]
                                    FehlerUrsache = "Pos." & rs2![Fehlerpositionen] & ":" & vbCrLf & rs2![FehlerUrsache]
                                    Maßnahmen = "Pos." & rs2![Fehlerpositionen] & ":" & vbCrLf & rs2![Maßnahmen]
                                    If IsNull(rs2![Sachbearbeiter]) Then
                                        Sachbearbeiter = ""
                                    Else
                                        Sachbearbeiter = rs2![Sachbearbeiter]
                                    End If
                                        If intI <= rs2.RecordCount Then

                                                WS.Range("Fehler") = Fehlertext
                                                WS.Range("Fehler").EntireRow.Select
                                                xlApp.Selection.Copy
                                                xlApp.Selection.Insert Shift:=xlDown

                                                WS.Range("Fehlerursache") = FehlerUrsache
                                                WS.Range("Fehlerursache").EntireRow.Select
                                                xlApp.Selection.Copy
                                                xlApp.Selection.Insert Shift:=xlDown

                                                WS.Range("Sachbearbeiter") = Sachbearbeiter
                                                WS.Range("Korrekturmaßnahme") = Maßnahmen
                                                WS.Range("Korrekturmaßnahme").EntireRow.Select
                                                xlApp.Selection.Copy
                                                xlApp.Selection.Insert Shift:=xlDown

                                        End If
                                    rs2.MoveNext
                                Next intI
                            Else
                                WS.Range("Fehler") = rs![Fehler]
                                WS.Range("Fehlerursache") = rs![FehlerUrsache]
                                WS.Range("Korrekturmaßnahme") = rs![Maßnahmen]
                                If IsNull(rs![Sachbearbeiter]) Then
                                    WS.Range("Sachbearbeiter") = ""
                                Else
                                    WS.Range("Sachbearbeiter") = rs![Sachbearbeiter]
                                End If
                            End If
                    Else
                        WS.Range("Fehler") = rs![Fehler]
                        WS.Range("Fehlerursache") = rs![FehlerUrsache]
                        WS.Range("Korrekturmaßnahme") = rs![Maßnahmen]
                        If IsNull(rs![Sachbearbeiter]) Then
                            WS.Range("Sachbearbeiter") = ""
                        Else
                            WS.Range("Sachbearbeiter") = rs![Sachbearbeiter]
                        End If
                    End If
            End If
            Set WS = Nothing
            bReportOeffnen = True
        Else
            If (MsgBox(getTC_FCFromCode("TC4210"), vbYesNo, XDReportTemplate) = vbYes) Then    'TC0256
                bReportOeffnen = True
            End If
        End If
    Else
        MsgBox getTC_FCFromCode("TC0257")  'MsgBox ("Reklamationsvorgangspfad nicht vorhanden!")  'TC0257
    End If

    Set fs = Nothing
    Set rs = Nothing
    Set rs2 = Nothing

    lexcelprov.closeExcelInstance
    Set lexcelprov = Nothing

    If bReportOeffnen Then
        ShellExecute 0, "open", sPfadDatei & "\" & Me.Controls("Mastertab#ID") & "_" & XDReportTemplate, vbNullString, vbNullString, 1
    End If

    DoCmd.Hourglass False

End Sub


Private Sub cmbAbholauftrag_Click()

    '1. Inhalt Dictionary sichern und anschließend leeren
    UpdatetFormTab Me, "[ID]=" & Me.Controls("Mastertab#ID")
    DoCmd.OpenReport "Abholauftrag", acViewReport, , "[ID]=" & Me.Controls("Mastertab#ID")
    DoCmd.OpenForm ("frmLager")

End Sub

Private Sub cmbDial_Click()

'    If IsNull(Me.frmUFKundenDaten.Controls("qryKundenDaten#Telefonnummer")) Then
''        MsgBox ("Kein Telefonnummer vorhanden!")
'        MsgBox getTC_FCFromCode("TC3227"), vbInformation, getTC_FCFromCode("TC3228")
'        Exit Sub
'    Else
'        DialCOM Me.frmUFKundenDaten.Controls("qryKundenDaten#Telefonnummer")
'    End If

    Forms("Fehlererfassung - Register").frmUFKundenDaten.SetFocus
    Forms("Fehlererfassung - Register").frmUFKundenDaten.Controls("qryKundenDaten#Telefonnummer").SetFocus

'    Me.txtTelNr.SetFocus
    SendKeys "{F8}"
    SendKeys "{NUMLOCK}"

End Sub


Private Sub cmbDial2_Click()

''    Dim ldb As DAO.Database
''    Dim rs As DAO.Recordset
'    Dim strTelefon As String
'
''    Set ldb = openAccDBConnection("openDB")
''    Set rs = ldb.OpenRecordset(("Select * from Sachbearbeiter where Name='" & Me.[Sachbearbeiter QS] & "';"))
'
'    strTelefon = Nz(DLookup("[TelefonBeruflich]", "Sachbearbeiter", "[Name] = '" & Me.Controls("Mastertab#Sachbearbeiter#1") & "'"), "")
'
'    If strTelefon = "" Then
''    If rs.RecordCount = 0 Then
''        MsgBox ("Kein Sachbearbeiter ausgewählt!")
'        MsgBox getTC_FCFromCode("TC3229"), vbInformation, getTC_FCFromCode("TC3230")
'        Exit Sub
'    Else
''        DialCOM rs("TelefonBeruflich")
'        DialCOM strTelefon
'    End If
'
''    Set rs = Nothing

    Me.Mastertab_Sachbearbeiter_1.SetFocus
    SendKeys "{F8}"
    SendKeys "{NUMLOCK}"

End Sub


Private Sub cmbEingangsbestätigung_Click()   '20160314-ROS

    Dim strKontakt As String
    Dim strBetreff As String
    Dim strBody As String
    Dim sPfadDatei As String
    Dim strTextHtml As String
    Dim strVon As String

    If MsgBox(getTC_FCFromCode("TC4211"), vbYesNo) = vbNo Then
        Exit Sub
    End If
        strBody = "<span style=""font-size:10pt; font-family:'Arial'"">" & _
                "Sehr geehrte Damen und Herren,<br><br>" & "hiermit bestätigen wir Ihnen den Eingang Ihrer Reklamation " & Me.Controls("Mastertab#Meldungsnummer Kunde#2") & " vom " & Me.Controls("Mastertab#Auftragsdatum#3").Value & ". <br><br>" & _
                "Der Vorgang wurde bei LQ unter der Nummer " & Me.Controls("Mastertab#ID") & " erfasst.<br><br>" & _
                "Für Rückfragen stehen wir Ihnen gerne zur Verfügung.<br><br>" & _
                "Mit freundlichen Grüßen <br>" & _
                "i.A. " & "xxxx <br><br>" & _
                "Mitarbeiter(in) Service <br>" & _
                "service@de.lq-group.com _Telefon +49.7143.968440 <br><br>" & _
                "LQ Mechatronik-Systeme GmbH<br>" & _
                "Carl-Benz-Straße 6 _D-74354 Besigheim<br>" & _
                "Fax +49.7143.9683-96 _www.lq-group.com<br><br>" & _
                "<span style=""font-size:7pt; font-family:'Arial'"">" & _
                "LQ Mechatronik-Systeme GmbH, Sitz: Freiberg am Neckar<br>" & _
                "Registergericht: Amtsgericht Stuttgart - HRB-Nr. 205651<br>" & _
                "Geschäftsführer: Friedrich P. Link, José Quero<br><br>" & _
                "Wichtiger Hinweis: Diese E-Mail und etwaige Anlagen enthalten<br>" & _
                "firmenvertrauliche Informationen. Sollten Sie diese E-Mail<br>" & _
                "irrtümlich erhalten haben, benachrichtigen Sie uns bitte durch<br>" & _
                "eine Antwort-Mail und löschen Sie diese E-Mail nebst Anlagen<br>" & _
                "von Ihrem System. Vielen Dank."

    strBetreff = "Eingangsbestätigung Reklamation " & Me.Controls("Mastertab#Meldungsnummer Kunde#2")
    strKontakt = Me.frmUFKundenDaten.Controls("qryKundenDaten#EmailAdresse")
    sPfadDatei = get_DateiPfad & "E-Mail\"
    strBody = Replace(strBody, "xxxx", Forms("Fehlererfassung - Register")![Mastertab#Zuständiger Service-MA])
    strVon = getTC_FCFromCode("TC0510")        'service@de.lq-group.com

    Call EMailSendenUndSpeichern(strVon, strKontakt, strBetreff, strBody, sPfadDatei)

End Sub


Private Sub cmbEMailAnKKI_Click()   '20180914-ROS
'Mail an KKI
'Bei Reklamationen von LQCN oder LQUS muss sich KKI um Ersatz und de finanziellen Ausgleich der Ersatzlieferung kümmern
'ERfassung = Ersatz
'Abschluss = Abrechnen

    Dim lclsMailProv As clsMailProviderOutlook
    Dim AntwortEigeneMsgBox As Variant
    Dim strKontakt As String
    Dim strBetreff As String
    Dim strBody As String
    Dim sPfadDatei As String
    Dim strTextHtml As String
    Dim strVon As String
    Dim Erfassung As String
    Dim Abschluss As String
    Dim KatStatus As String

    '1. Inhalt Dictionary sichern und anschließend leeren
    UpdatetFormTab Me, "[ID]=" & Me.Controls("Mastertab#ID")

'    AntwortEigeneMsgBox = EigeneMsgBox("Bitte E-Mail auswählen", vbYesNoCancel, "E-Mail an KKI")
    AntwortEigeneMsgBox = EigeneMsgBox(getTC_FCFromCode("TC3830"), vbYesNoCancel, getTC_FCFromCode("TC3831"))
    If AntwortEigeneMsgBox = vbCancel Then
        Exit Sub
    End If

    'Instanz auf einen Outlook MailProvider
    Set lclsMailProv = New clsMailProviderOutlook
    'Mailinhalte = Body füllen
    lclsMailProv.setBody = "Erfassung" & _
        "Für Rückfragen stehen wir Ihnen gerne zur Verfügung.<br><br>" & _
        "Mit freundlichen Grüßen <br>" & _
        "i.A. " & "xxxx <br><br>" & _
        "Mitarbeiter(in) Service <br>" & _
        "service@de.lq-group.com _Telefon +49.7143.968440 <br><br>" & _
        "LQ Mechatronik-Systeme GmbH<br>" & _
        "Carl-Benz-Straße 6 _D-74354 Besigheim<br>" & _
        "Fax +49.7143.9683-96 _www.lq-group.com<br><br>" & _
        "<span style=""font-size:7pt; font-family:'Arial'"">" & _
        "LQ Mechatronik-Systeme GmbH, Sitz: Freiberg am Neckar<br>" & _
        "Registergericht: Amtsgericht Stuttgart - HRB-Nr. 205651<br>" & _
        "Geschäftsführer: Friedrich P. Link, José Quero<br><br>" & _
        "Wichtiger Hinweis: Diese E-Mail und etwaige Anlagen enthalten<br>" & _
        "firmenvertrauliche Informationen. Sollten Sie diese E-Mail<br>" & _
        "irrtümlich erhalten haben, benachrichtigen Sie uns bitte durch<br>" & _
        "eine Antwort-Mail und löschen Sie diese E-Mail nebst Anlagen<br>" & _
        "von Ihrem System. Vielen Dank."

   KatStatus = Nz(DLookup("[Kategorie-Problem]", "[Mastertab]", "[ID] = " & Me.Controls("Mastertab#ID") & ""))

    If AntwortEigeneMsgBox = vbYes Then
        'ERsatz senden
        Erfassung = "<span style=""font-size:10pt; font-family:'Arial'"">" & _
            "Hallo,<br><br>" & "Reklamation " & Me.Controls("Mastertab#Meldungsnummer Kunde#2") & " - " & Form("frmUFKundenDaten").Controls("qryKundenDaten#Firma_Name") & _
            " vom " & Me.Controls("Mastertab#Auftragsdatum#3").Value & ". <br><br>" & _
            "Der Vorgang wurde bei LQ unter der Nummer " & Me.Controls("Mastertab#ID") & " erfasst.<br><br><br>"
        lclsMailProv.setBody = Replace(lclsMailProv.getBody, "Erfassung", Erfassung)
        lclsMailProv.setBetreff = "Erfassung Reklamation " & Me.Controls("Mastertab#Meldungsnummer Kunde#2") & " - " & _
                                  Form("frmUFKundenDaten").Controls("qryKundenDaten#Firma_Name") & "  LQ_" & Me.Controls("Mastertab#ID")
    Else
        'Finanzieller Ausgleich
        Abschluss = "<span style=""font-size:10pt; font-family:'Arial'"">" & _
            "Hallo,<br><br>" & "Reklamation " & Me.Controls("Mastertab#Meldungsnummer Kunde#2") & " - " & Form("frmUFKundenDaten").Controls("qryKundenDaten#Firma_Name") & _
            " vom " & Me.Controls("Mastertab#Auftragsdatum#3").Value & ". <br><br>" & _
            "Reklamation Vorgangs-Nummer " & Me.Controls("Mastertab#ID") & " ist " & KatStatus & ".<br><br><br>"
        lclsMailProv.setBody = Replace(lclsMailProv.getBody, "Erfassung", Abschluss)
        lclsMailProv.setBetreff = "Abschluss Reklamation " & Me.Controls("Mastertab#Meldungsnummer Kunde#2") & " - " & _
                                  Form("frmUFKundenDaten").Controls("qryKundenDaten#Firma_Name").Value & "  LQ_" & Me.Controls("Mastertab#ID")
    End If
    'Mailobjekt füllen, versenden und abspeichern
    lclsMailProv.setAn = "KKI@de.lq-group.com"
    lclsMailProv.setBody = Replace(lclsMailProv.getBody, "xxxx", Forms("Fehlererfassung - Register")![Mastertab#Zuständiger Service-MA])
    lclsMailProv.setVon = getTC_FCFromCode("TC0510")        'service@de.lq-group.com

    lclsMailProv.setBetreff = "Meldung " & Me.Controls("Mastertab#Meldungsnummer Kunde#2").Value & " /" & "  LQ " & Me.Controls("Mastertab#ID").Value
    lclsMailProv.fillMailItem
    lclsMailProv.SaveAsMailItem get_DateiPfad & "E-Mail\", get_pKuerzel()
    lclsMailProv.displayMailItem
'    lclsMailProv.sendMailItem
    Set lclsMailProv = Nothing

End Sub


Private Sub cmbMail_Click()
'Sendet anhand der Kundendaten eine Mail an den Absender
'Hierbei wird die Mail anhand der bekannten Absenderdaten gefüllt und der Mail als eigener Dialog zu Weiterbearbeitung geöffnet.
'Die Mail kann hierbei im Nachhinein mit den noch notwendigen Informatione verändert werden.
'Das Versenden der Mail wird über den Mail-Dialog gesteuert.

        Dim lclsMailProv As clsMailProviderOutlook
        Dim strBody As String
        Dim strKontaktperson As String
        Dim AntwortEigeneMsgBox As Variant
        Dim Ansprache As String

'        AntwortEigeneMsgBox = EigeneMsgBox("Ansprache auswählen", vbYesNoCancel, "Ansprache")
        'Über die MsgBox wird die Anrede Herr/Frau festgelegt
        AntwortEigeneMsgBox = EigeneMsgBox(getTC_FCFromCode("TC3832"), vbYesNoCancel, getTC_FCFromCode("TC3833"))
        If AntwortEigeneMsgBox = vbYes Then
            Ansprache = getTC_FCFromCode("TC0498")
        End If
        If AntwortEigeneMsgBox = vbNo Then
            Ansprache = getTC_FCFromCode("TC0499")
        End If
        If AntwortEigeneMsgBox = vbCancel Then
            Exit Sub
        End If

        'Name der Kontaktperson für die Anrede erstellen
        strKontaktperson = Left(Me.frmUFKundenDaten.Controls("qryKundenDaten#Kontaktperson"), InStr(Me.frmUFKundenDaten.Controls("qryKundenDaten#Kontaktperson"), ",") - 1)

        'Instanz auf einen Outlook MailProvider
        Set lclsMailProv = New clsMailProviderOutlook
        'Mailinhalte füllen
        lclsMailProv.setBody = "<span style=""font-size:10pt; font-family:'Arial'"">" & _
                Ansprache & " " & strKontaktperson & ",<br><br><br><br><br><br>" & _
                "Für Rückfragen stehen wir Ihnen gerne zur Verfügung.<br><br>" & _
                "Mit freundlichen Grüßen <br>" & _
                "i.A. " & "xxxx <br><br>" & _
                "Mitarbeiter(in) Service <br>" & _
                "service@de.lq-group.com _Telefon +49.7143.968440 <br><br>" & _
                "LQ Mechatronik-Systeme GmbH<br>" & _
                "Carl-Benz-Straße 6 _D-74354 Besigheim<br>" & _
                "Fax +49.7143.9683-96 _www.lq-group.com<br><br>" & _
                "<span style=""font-size:7pt; font-family:'Arial'"">" & _
                "LQ Mechatronik-Systeme GmbH, Sitz: Freiberg am Neckar<br>" & _
                "Registergericht: Amtsgericht Stuttgart - HRB-Nr. 205651<br>" & _
                "Geschäftsführer: Friedrich P. Link, José Quero<br><br>" & _
                "Wichtiger Hinweis: Diese E-Mail und etwaige Anlagen enthalten<br>" & _
                "firmenvertrauliche Informationen. Sollten Sie diese E-Mail<br>" & _
                "irrtümlich erhalten haben, benachrichtigen Sie uns bitte durch<br>" & _
                "eine Antwort-Mail und löschen Sie diese E-Mail nebst Anlagen<br>" & _
                "von Ihrem System. Vielen Dank."
        'Mailobjekt füllen und Mail über Outlook öffnen und anzeigen
        lclsMailProv.setBody = Replace(lclsMailProv.getBody, "xxxx", Me.Controls("Mastertab#Zuständiger Service-MA"))
        lclsMailProv.setVon = getTC_FCFromCode("TC0510")        'service@de.lq-group.com
        lclsMailProv.setAn = Me.frmUFKundenDaten.Controls("qryKundenDaten#EmailAdresse").Value
        lclsMailProv.setBetreff = "Meldung " & Me.Controls("Mastertab#Meldungsnummer Kunde#2").Value & " /" & "  LQ " & Me.Controls("Mastertab#ID").Value
        lclsMailProv.fillMailItem
        lclsMailProv.displayMailItem
        Set lclsMailProv = Nothing

End Sub

Private Sub cmbNotizErstellen_Click()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim WinAnmeldung As String
    Dim BearbKz As String
    Dim strParameter As String
    Dim strKunde As String
    Dim strKD As String
    Dim strKdDaten As String


    '1. Inhalt Dictionary sichern und anschließend leeren
    UpdatetFormTab Me, "[ID]=" & Me.Controls("Mastertab#ID")

    WinAnmeldung = fOSUserName
    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from Sachbearbeiter where [Name] = '" & WinAnmeldung & "';", dbOpenSnapshot)
        If rs.RecordCount Then
            BearbKz = rs("Kurzzeichen")
        Else
'            MsgBox ("Sacharbeiter ist nicht in Datenbank eingetragen")
            MsgBox getTC_FCFromCode("TC3213"), vbInformation, getTC_FCFromCode("TC3214")
            Exit Sub
        End If

    strKunde = Nz(DLookup("[Firma_Name]", "qryFirmenName", "[Kunden-ID]=" & Me.Mastertab_Kunden_ID_1.Value))
    strKD = Nz(DLookup("[KndNr_Avista]", "qryFirmenName", "[Kunden-ID]=" & Me.Mastertab_Kunden_ID_1.Value))
    strKdDaten = strKD & " / " & strKunde

'    strParameter = "" & Me.Controls("Mastertab#ID") & ", " & Me.txtAktiverStatus.Value & "," & BearbKz & "," & strKunde & ""
    strParameter = "" & Me.Controls("Mastertab#ID") & ", " & Me.txtAktiverStatus.Value & "," & BearbKz & "," & strKdDaten & ""

    DoCmd.OpenForm "frmNotizen", acNormal, , , acFormAdd, acDialog, strParameter

    Set rs = Nothing

End Sub


Private Sub cmbReklamationUnterpositionAnlegen_Click()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsFehlerPos As DAO.Recordset
    Dim rsneu As DAO.Recordset
    Dim reklnummer As String
    Dim MeldungsnummerKunde As String
    Dim la_zahlenpos() As String
    Dim sPfadDatei As String
    Dim Auftragsdatum As String
    Dim li_Fehlerpos As Integer
    Dim li_AnzRef As Integer


    '1. Inhalt Dictionary sichern und anschließend leeren
    UpdatetFormTab Me, "[ID]=" & Me.Controls("Mastertab#ID")

    Set ldb = openAccDBConnection("openDB")

    If IsNull(Me.Controls("Mastertab#Meldungsnummer Kunde#2")) Then
        MeldungsnummerKunde = "LQ" & Me.Controls("Mastertab#ID")
    Else
        MeldungsnummerKunde = Me.Controls("Mastertab#Meldungsnummer Kunde#2")
    End If


    unlockDataset Me.Controls("Mastertab#ID")

    'Benötigte Daten zum erstellen der Unterpositionen ermitteln keine Unterposition vorhanden
    If IsNull(Me.Controls("Mastertab#Referenznummer")) Or Me.Controls("Mastertab#Referenznummer") = "" Then
        Set rs = ldb.OpenRecordset("Select * from ROS_DatensatzKopieren where [ID] = " & Me.Controls("Mastertab#ID") & ";", dbOpenSnapshot)
        li_AnzRef = DCount("[ID]", "Mastertab", "[Referenznummer] = '" & Me.Controls("Mastertab#ID") & "'")
        reklnummer = Me.Controls("Mastertab#ID")
    Else
        Set rs = ldb.OpenRecordset("Select * from ROS_DatensatzKopieren where [ID] = " & Me.Controls("Mastertab#Referenznummer") & ";", dbOpenSnapshot)
        li_AnzRef = DCount("[ID]", "Mastertab", "[Referenznummer] = '" & Me.Controls("Mastertab#Referenznummer") & "'")
        reklnummer = Me.Controls("Mastertab#Referenznummer")
    End If
    If li_AnzRef = 0 Then
        li_Fehlerpos = 1
    Else
        li_Fehlerpos = li_AnzRef + 1
    End If
    'erhöhen der Fehlerpositionen in Basis-Meldung
    Set rsFehlerPos = ldb.OpenRecordset("Select * from Mastertab where [ID] = " & reklnummer & ";")
    rsFehlerPos.Edit
    rsFehlerPos("Fehlerpositionen") = li_Fehlerpos
    rsFehlerPos.Update

    NewReklamationE "Erfassen", ""

    'Neu angelegte Unterposition bearbeiten
    Set rsneu = ldb.OpenRecordset("Select * from ROS_DatensatzKopieren where ID = " & get_p_ReklNr() & ";")
            rsneu.Edit
            rsneu("Auftragsdatum") = rs("Auftragsdatum")
            rsneu("Kunden-ID") = rs("Kunden-ID")
            rsneu("Bestellnummer Kunde") = rs("Bestellnummer Kunde")
'            rsneu("VorgangstypID") = rs("VorgangstypID")
            rsneu("VorgangstypID") = 1      'Meldung    20181206-ROS
            rsneu("Vorgangs_Eingang") = rs("Vorgangs_Eingang")
            rsneu("Maschinennummer") = rs("Maschinennummer")
            rsneu("Rückmeldung_Kunde_Soll") = rs("Rückmeldung_Kunde_Soll")
            rsneu("Rückmeldung_Kunde_Verschiebung") = rs("Rückmeldung_Kunde_Verschiebung")
            rsneu("per Telefon") = rs("per Telefon")
            rsneu("per E-Mail/Fax") = rs("per E-Mail/Fax")
            rsneu("mit Rücksendung") = rs("mit Rücksendung")
            rsneu("Referenznummer") = reklnummer
            rsneu("Meldungsnummer Kunde") = MeldungsnummerKunde
            rsneu("Fehlerpositionen") = li_Fehlerpos
            rsneu("Welcher Fehler tritt auf") = rs("Welcher Fehler tritt auf")
            rsneu.Update

    Set rs = Nothing
    Set rsFehlerPos = Nothing
    Set rsneu = Nothing

    'Laufzeit-Ticket "LZ-Vorgang" erstellen
    Dim lTicketID As Integer

    lTicketID = TicketErstellen(14, get_p_ReklNr)
    TicketFuellen lTicketID, Date, "MailDatum"
    TicketFuellen lTicketID, "2", "TicketStatusID"
    TicketFuellen lTicketID, "Service", "MailAn"

'    NewReklamationE "Erfassen", ""
    Call AnsichtFehlerpositionen
    Call RegisterStr246_Change
    Call VorgangsOrdner(Me.Controls("Mastertab#ID"), MeldungsnummerKunde, Me.Controls("Mastertab#Auftragsdatum#3"), False)
    Me.Requery

End Sub

Private Sub cmbTicketVersenden_Click()

    Dim Kriterium As String
    Dim strTest As String

    '1. Inhalt Dictionary sichern und anschließend leeren
    UpdatetFormTab Me, "[ID]=" & Me.Controls("Mastertab#ID")


    If IsNull(Me.Controls("Mastertab#Sachbearbeiter#1")) And (Me.Controls("Mastertab#Verursacher#1") = "Lieferant") Then
'        MsgBox ("Bitte Sachbearbeiter auswählen!")
        MsgBox getTC_FCFromCode("TC3229"), vbInformation, getTC_FCFromCode("TC3214")
        Exit Sub
    End If

    strTest = Nz(DLookup("[TaetigkeitLieferantID]", "ROS_Daten_Einkauf", "[ID-ReklNr]=" & Me.Controls("Mastertab#ID")), 0)

    If CurrentProject.AllForms("frmTicketErstellen").IsLoaded = True Then
        DoCmd.Close acForm, "frmTicketErstellen", acSaveYes
    End If

    If IsNull(Me.Controls("Mastertab#LQEinheitID#1")) Or (Me.Controls("Mastertab#LQEinheitID#1")) = 0 Then
'        MsgBox ("Bitte Verursacher-Einheit auswählen, wir für die E-Mail Adressen benötigt")
        MsgBox getTC_FCFromCode("TC3232"), vbInformation, getTC_FCFromCode("TC3210")
        Exit Sub
    Else
        If Me.Controls("Mastertab#Verursacher#1") = "Lieferant" And (strTest = "-1" Or Me.Controls("Mastertab#LiefAnsprechp-ID#3") = "" Or IsNull(Me.Controls("Mastertab#Sachbearbeiter#1")) Or IsNull(Me.Controls("Mastertab#LQEinheitID#1"))) Then
'            MsgBox ("Ticket kann nicht erstellt werden Daten fehlen!" & vbNewLine & "Tätigkeiten-Lieferant?" & vbNewLine & _
'                    "Ansprechpartner-Lieferant?" & vbNewLine & "Sachbearbeiter?" & vbNewLine & "Verursacher-Einheit?")
            MsgBox getTC_FCFromCode("TC3233"), vbInformation, getTC_FCFromCode("TC3234")
            Exit Sub
        Else
            'Kriterium = "(ErledigtAM Is Null) And (ReklID = " & Me.Controls("Mastertab#ID") & " AND TicketTypID <> 11 AND TicketTypID <> 12)"
'            DoCmd.OpenForm "frmTicketListe", , , , , acDialog, "Neu erstellen"
            DoCmd.OpenForm "frmTicketErstellen", , , , , acDialog, "Neu erstellen"  '20251216-ROS
        End If
    End If

End Sub


Private Sub cmButDatenUebernehmen_Click()

    Me.Controls("Mastertab#Bemerkungen - Maßnahmen#2") = Nz(DLookup("[Welcher Fehler tritt auf]", "Mastertab", "[ID] = " & Me.Controls("Mastertab#ID") & ""))
    Call Mastertab_Bemerkungen___Maßnahmen_2_AfterUpdate

End Sub

Private Sub CmdCheckliste_Click()

On Error GoTo err_CmdCheckliste_Click

    Dim fs As Object
    Dim bReportOeffnen As Boolean
    Dim sPfadDatei As String


    Dim XDReportPfad As String  '20160304-ROS
    XDReportPfad = getTC_FCFromCode("DP0001")

    Dim ChecklisteTemplate As String  '20160701-ROS
    ChecklisteTemplate = getTC_FCFromCode("DP0006")

    'Init
    bReportOeffnen = False


    Set fs = CreateObject("Scripting.FileSystemObject")
    'ist Verzeichnis vorhanden
    'sPfadDatei = getReklVerzeichnis(Me.Controls("Mastertab#ID"), Format(Me.Auftragsdatum, "yyyy"), Me.[Meldungsnummer Kunde])
    If (IsNull(Me.Controls("Mastertab#Meldungsnummer Kunde#2")) Or Me.Controls("Mastertab#Meldungsnummer Kunde#2") = "") Then
        sPfadDatei = setReklVerzeichnis(Me.Controls("Mastertab#ID"), format(Me.Controls("Mastertab#Auftragsdatum#3"), "yyyy"), "")
    Else
        sPfadDatei = setReklVerzeichnis(Me.Controls("Mastertab#ID"), format(Me.Controls("Mastertab#Auftragsdatum#3"), "yyyy"), Me.Controls("Mastertab#Meldungsnummer Kunde#2"))
    End If
    If sPfadDatei <> "" Then
        If fs.FileExists(sPfadDatei & "\" & Me.Controls("Mastertab#ID") & "_Checkliste.docx") = False Then
            fs.CopyFile XDReportPfad & "\" & ChecklisteTemplate, sPfadDatei & "\" & Me.Controls("Mastertab#ID") & "_Checkliste.docx", False
        End If
        bReportOeffnen = True
    End If

    Set fs = Nothing

    If bReportOeffnen Then
        ShellExecute 0, "open", sPfadDatei & "\" & Me.Controls("Mastertab#ID") & "_Checkliste.docx", "", "", 1
    End If

end_CmdCheckliste_Click:
    Exit Sub

err_CmdCheckliste_Click:

    MsgBox Error

End Sub

Private Sub Form_AfterUpdate()

On Error GoTo err_Form_AfterUpdate

p_KommentarKopieren = True

DoCmd.RunCommand acCmdSaveRecord

err_Form_AfterUpdate:

End Sub


Private Sub Form_GotFocus()
'    'DoCmd.Maximize

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me, Me.Controls("RegisterStr246").Pages(getActivePage())
'    KeyDown KeyCode, Shift, Me, Me.Controls("RegisterStr246").Pages(Me.Controls("RegisterStr246").Value)

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub

Private Sub Form_Load()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim objMouseOverForm As clsMouseoverForm
    Dim ContollAktiv As Boolean
    Dim ls_activeReg As String
    Dim l_VorgangArtIntern As Integer
    Dim Test As Boolean

    Call FormEinUndAusBlenden(Me, False)
    '#####################################################
    'Beseitigt das Flackern bei Requery
'    Application.Echo False
    '#####################################################

    'Anlegen Formular in RefFormList
    FormOpenDictionary Me

    Me.Controls("Mastertab#ID") = get_p_ReklNr()

    ls_activeReg = set_FehlererfassungEnviroment
    Call RegisterStr246_Change
    Test = LoadControlsFormAfterOpen(Me, getRecordSource(), "[ID]=" & get_p_ReklNr)

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select Value from KonfigurationReklDBLocal where Konfiguration = 'RibbonOnOff';", dbOpenSnapshot)
    If rs("Value") = "An" Then
        Me.cboRibbonOn.Caption = "Ribbon off"
    Else
        Me.cboRibbonOn.Caption = "Ribbon on"
    End If

    If Not IsNull(Me.[ReklMemo#Portal]) Then
        Me.TruTicket.Enabled = True
    End If

    Call AnsichtFehlerpositionen

    setCmbViewFalse Me

    If (getCurrLanguage = "German") Then getCmbView Me, "German"
    If get_AdminZuordnung = True Then getCmbView Me, "Admin"
    If (getCurrLanguage = "German") And get_ServiceZuordnung = True Then getCmbView Me, "Service/German"
    If get_ServiceZuordnung = True Then
        If (getCurrLanguage = "German") Then Me.cmbEmailSendenDispo.Visible = True
        getCmbView Me, "Service"
        Me.Controls("ReklMemo#Vorgangs_Eingang").Locked = False
    End If

    If Me.Controls("Mastertab#intern") = True Then
        Dim ReklnrKnd As String
        If IsNull(Me.Controls("Mastertab#Meldungsnummer Kunde#2")) Then
            ReklnrKnd = ""
        Else
            ReklnrKnd = Me.Controls("Mastertab#Meldungsnummer Kunde#2")
        End If
        Call VorgangsOrdner(Me.Controls("Mastertab#ID"), ReklnrKnd, Me.Controls("Mastertab#Auftragsdatum#3"), False)        '20190820-ROS
    End If

    Me.UF_Fehlererfassung.Requery
    If (getCurrLanguage = "German") = True Then Me.txtEMailAblegen.SetFocus

    '#####################################################
    'Beseitigt das Flackern bei Requery
'    Application.Echo True
    '#####################################################

    Set rs = Nothing

    If IsLoaded("frmFormularROS") Then
        With Forms("frmFormularROS")
            .Form("frmÜbersichtNotizen").Form.Filter = "[ReklID] = " & Me.Mastertab_ID
            .Form("frmÜbersichtNotizen").Form.FilterOn = True
            .Refresh
        End With
    End If

    If IsLoaded("frmÜbersichtNotizen") Then
        With Forms("frmÜbersichtNotizen")
            .Filter = "[ReklID] = " & Me.Mastertab_ID
            .FilterOn = True
            .Refresh
        End With
    End If

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)
    Me.SetFocus

End Sub

Private Sub Fracht_Kosten_AfterUpdate()

    Me.Refresh

End Sub

Private Sub GeheZu_Click()
On Error GoTo err_GeheZu_Click

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim strKrit As String
Dim Reklnr As Long
Dim sMemo As String

    'Bestehenden DS sichern
'    DoCmd.RunCommand acCmdSaveRecord
    Call Befehl6068_Click
'    p_ReklNr = InputBox("Geben Sie die Reklamationsnummer ein!", "Eingabe Vorgang")  'IC0205
'    p_ReklNr = InputBox(getTC_FCFromCode("IC0200"), getTC_FCFromCode("IC0219"))
    p_ReklNr = getValueFromInputBox(getTC_FCFromCode("IC0200"), getTC_FCFromCode("IC0219"))
    If StrPtr(p_ReklNr) = 0 Then ' Prüfen ob Abbrechen gedrückt wurde
        Exit Sub
    End If

    strKrit = "ID = " & p_ReklNr
    If DCount("ID", "Mastertab", strKrit) = 0 Then
        MsgBox getTC_FCFromCode("TC0205") & p_ReklNr & getTC_FCFromCode("TC0379")  'MsgBox ("Reklamation mit der Nummer: " & p_ReklNr & " ist nicht existent!")  'TC0239
        GoTo end_GeheZu_Click
    Else

        Set ldb = openAccDBConnection("openDB")
        'Ermittlung ob es zu dem Vorgang schon ein Protokolleintrag vorhanden ist
        Reklnr = Me.Controls("Mastertab#ID")
        If p_KommentarKopieren Then
            Set rs = ldb.OpenRecordset("Select * from ReklMemo where [ID] = " & Reklnr & ";", dbOpenSnapshot)
            If rs.RecordCount Then
                If IsNull(rs("Protokoll").Value) Then
                    sMemo = ""
                Else
                    sMemo = rs("Protokoll").Value
                End If
                If Not IsNull(rs("Aktueller Eintrag").Value) Then
                    sMemo = rs("Aktueller Eintrag").Value & Chr(13) & Chr(10) & Chr(13) & Chr(10) & sMemo
                End If
                ldb.Execute ("Update ReklMemo set [Protokoll] = '" & sMemo & _
                            "' where ID = " & Reklnr & ";")
                sMemo = ""
                ldb.Execute ("Update ReklMemo set [Aktueller Eintrag] = NULL where ID = " & Reklnr & ";")
            End If
        End If

        unlockDataset (Reklnr)

        GeheZuVorgang p_ReklNr
        Set rs = Nothing
    End If


end_GeheZu_Click:
Exit Sub
err_GeheZu_Click:
    'MsgBox (Err)
    If Err = 3188 Or Err = 2046 Then
        Resume Next
    Else
        GoTo end_GeheZu_Click
    End If


End Sub
Private Sub GeheZuVorgang(fReklNr As String)
On Error GoTo err_GeheZuVorgang

    Dim lb_reload As Boolean

    If GetStateVorgang(fReklNr) = "Abgeschlossen" Then
        Me.AllowEdits = False
        Me.AllowAdditions = False
        Me.AllowDeletions = False

'            Me.RecordsetType = 2 'Snapshot
    Else
        'RSP - 2012-05-30 Filter wird hier zentral gesetzt.
        InsertNewReklMemoDS (fReklNr)
        p_KommentarKopieren = True
        If proofAndLockDataset(fReklNr) = "" Then
            Me.AllowEdits = True
            Me.AllowAdditions = True
            Me.AllowDeletions = True

'                Me.RecordsetType = 0 'Dynaset
'                ReklBearbeiten = True
        Else
            Me.AllowEdits = False
            Me.AllowAdditions = False
            Me.AllowDeletions = False
'                ReklBearbeiten = False
'                Me.RecordsetType = 2 'Snapshot
        End If
    End If

    p_ReklNr = fReklNr

    lb_reload = False
    If IsLoaded("Fehlererfassung - Register") Then lb_reload = True
    DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , , acFormEdit
    If lb_reload Then Call Forms("Fehlererfassung - Register").ReLoadForm

end_GeheZuVorgang:
Exit Sub
err_GeheZuVorgang:
    MsgBox (Err)


End Sub


Private Sub GeheZuReferenz_Click()

Dim ldb As DAO.Database
Dim BasisRekl() As String

    '1. Inhalt Dictionary sichern und anschließend leeren
    UpdatetFormTab Me, "[ID]=" & Me.Controls("Mastertab#ID")

'    If IsNull(Me.Controls("Mastertab#Referenznummer")) Then
    If Me.Controls("Mastertab#Referenznummer") = "" Or IsNull(Me.Controls("Mastertab#Referenznummer")) Then
'        MsgBox ("Ist Basis-Reklamation!")
        MsgBox getTC_FCFromCode("TC3235"), vbInformation, getTC_FCFromCode("TC3236")
        Exit Sub
    Else
        BasisRekl() = Split(Me.Controls("Mastertab#Referenznummer"), "_")
        unlockDataset Me.Controls("Mastertab#ID")
        GeheZuVorgang BasisRekl(0)
    End If

    Call AnsichtFehlerpositionen

End Sub


Private Sub Drucken_aktueller_Datensatz_Click()
On Error GoTo Err_Drucken_aktueller_Datensatz_Click             'RSP - 16.05.2011

    Dim ilReklID As Long

    '1. Inhalt Dictionary sichern und anschließend leeren
    UpdatetFormTab Me, "[ID]=" & Me.Controls("Mastertab#ID")

    'RSP - Neu erstellt 13.05.2011

    ilReklID = Me.Controls("Mastertab#ID")

    DoCmd.Save acForm, "Fehlererfassung - Register"             'RSP - 16.05.2011
    DoCmd.OpenForm "F06_Fehlerbericht_11_Sept08", acNormal, , "[ID]=" & ilReklID
'    Forms("").Requery
'    Forms("").Refresh
    Forms("F06_Fehlerbericht_11_Sept08")![Gefundener Datensatz] = ilReklID
    DoCmd.PrintOut acSelection, 1, 1, acMedium
    DoCmd.Close acForm, "F06_Fehlerbericht_11_Sept08"
    Exit Sub
Err_Drucken_aktueller_Datensatz_Click:
    Resume Next

End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim objMouseOverForm As clsMouseoverForm

    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With

End Sub


Private Sub Kontakt_Kurzzeichen_NotInList(NewData As String, Response As Integer)

    MsgBox getTC_FCFromCode("TC0242")  'MsgBox "Doppelklicken Sie auf dieses Feld, um einen neuen Ansprechpartner/Kontaktperon einrichten zu können!"  'TC0242
    Response = acDataErrContinue

End Sub

Private Sub Mastertab_Lieferanten_Code_1_AfterUpdate()
    'RSP - 20130725: Bei EIngabe eines Lieferanten wird in Abhängigkeit der Lieferantenverantwortung die entsprechende Fehlerklasse gesetzt
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsADO As ADODB.Recordset
    Dim iLiefAvistaID As Long
    Dim lrs As DAO.Recordset
    Dim Response As Integer
    Dim blIntern As Boolean
    Dim rsIntern As DAO.Recordset
    Dim lApAvistaLeer As Boolean

    Set ldb = openAccDBConnection("openDB")
    lApAvistaLeer = True

    If Me.Controls("Mastertab#Lieferanten-Code 1") = 0 Or IsNull(Me.Controls("Mastertab#Lieferanten-Code 1")) Then
        Me.Controls("Mastertab#LiefAnsprechp-ID#3") = 0
        Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#LiefAnsprechp-ID#3"))
        Set lrs = Nothing
        Me.frmUFLiefernatDaten.Controls("AnsprechpartnerID#Email") = ""
    Else
        'Liste aller Ansprechpartner des Lieferanten ermitteln und in Temptabelle eintragen ########20221124-ROS#########
        blIntern = Nz(DLookup("[Intern]", "[Lieferanten]", "[Lieferant-ID] = " & Me.Controls("Mastertab#Lieferanten-Code 1") & ""))
        If blIntern = True Then
            'Liste aus interner Tabelle "Ansprechpartner"
            Set rsIntern = ldb.OpenRecordset("Select * from [LieferantAnsprechpartner] where [LieferantID] = " & Me.Controls("Mastertab#Lieferanten-Code 1") & " ;")
            If rsIntern.RecordCount Then
                DoCmd.OpenForm "LieferantAnsprechpartnerTemp", acNormal, , , acFormEdit
                Do While Not rsIntern.EOF
                    ldb.Execute ("Insert into tblTempAnspLief ([pa_org_name1], [pa_pers_vorname], [asp_position], [adb_email], [adb_tel1],[pa_id])" & _
                    "values('" & rsIntern("Name") & "', '" & rsIntern("Vorname") & "' , '" & rsIntern("Position") & "', " & _
                    "'" & rsIntern("Email") & "', '" & rsIntern("Telefon") & "', '" & rsIntern("AnsprechpartnerID") & "');")
                rsIntern.MoveNext
                Loop
                Set rsIntern = Nothing
            End If
        Else
            'Liste Ansprechpartner aus Avista
            Set rs = ldb.OpenRecordset("Select * from Lieferanten where ([Lieferant-ID] = " _
                                                      & Me.Controls("Mastertab#Lieferanten-Code 1") & ");")
            If rs.RecordCount > 0 Then
                iLiefAvistaID = rs("L-Nummer/AVISTA")
            End If

            Set rs = Nothing

            'Avista DB Provider Instanz erzeugen
            Set adbprov = New clsAvistaProvider

            Set rs = ldb.OpenRecordset("Select * from KonfigurationReklDB where Konfiguration = 'AvistaReklUser';")
            rs.MoveFirst
            Response = acDataErrAdded
            'MsgBox ("Der eingegebene Wert: " & NewData)
            Set rsADO = adbprov.getRecordsetAvista(SelectAnsprechp & WhereStartLiefNr & "'" & Trim(Str(iLiefAvistaID)) & "' " & WherePara1_Pos & "'" & rs("Value") & "' " & WhereOrderAnsprp)

            If rsADO.RecordCount Then
                DoCmd.OpenForm "LieferantAnsprechpartnerTemp", acNormal, , , acFormEdit
                Do While Not rsADO.EOF
                    ldb.Execute ("Insert into tblTempAnspLief ([pa_org_name1], [pa_pers_vorname], [asp_position], [asp_info11], [adb_email], [adb_tel1], [pa_id])" & _
                    "values('" & rsADO("pa_org_name1") & "', '" & rsADO("pa_pers_vorname") & "' , '" & rsADO("asp_position") & "', '" & rsADO("asp_info11") & "', " & _
                    "'" & rsADO("adb_email") & "', '" & rsADO("adb_tel1") & "', '" & rsADO("pa_id") & "');")
                rsADO.MoveNext
                Loop
            Else
                MsgBox getTC_FCFromCode("TC4190"), vbInformation, getTC_FCFromCode("TC4191")
                lApAvistaLeer = False
                Me.Controls("Mastertab#LiefAnsprechp-ID#3") = 0
                Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#LiefAnsprechp-ID#3"))
                Set lrs = Nothing
            End If

            Response = acDataErrContinue
            rsADO.Close
            Set rsADO = Nothing
            Set adbprov = Nothing
            Set rs = Nothing

        End If

        If lApAvistaLeer = True Then
            Forms("LieferantAnsprechpartnerTemp").Requery
            Forms("LieferantAnsprechpartnerTemp").AllowAdditions = False
            Forms("LieferantAnsprechpartnerTemp").AllowEdits = False
            sCurrentFehlererfassung = "Fehlererfassung - Register"
        End If

    End If

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Lieferanten-Code 1"))
    Set lrs = Nothing

    Call Befehl6068_Click

ende:
    Exit Sub

Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0243
    Resume ende

End Sub
Private Sub Mastertab_LQBereichID_1_AfterUpdate()

    Dim lrs As DAO.Recordset
    'Abhängige Werte setzen
    Me.Controls("Mastertab#LQWorkstationID#1") = -1


    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#LQBereichID#1"))
    Set lrs = Nothing

End Sub


Private Sub Mastertab_LQBereichErfasserID_AfterUpdate()

    Dim lrs As DAO.Recordset

    Me.Controls("Mastertab#LQWorkstationErfasserID") = -1
    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#LQBereichErfasserID"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_LQEinheitID_1_AfterUpdate()

    Dim lrs As DAO.Recordset
    'Abhängige Werte setzen
    Me.Controls("Mastertab#LQBereichID#1") = -1


    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#LQEinheitID#1"))
    Set lrs = Nothing

End Sub


Private Sub Mastertab_LQEinheitErfasserID_AfterUpdate()

    Dim lrs As DAO.Recordset

    Me.Controls("Mastertab#LQBereichErfasserID") = -1
    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#LQEinheitErfasserID"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_LQWorkstationID_1_AfterUpdate()

    Dim lrs As DAO.Recordset
    'Abhängige Werte setzen
    Me.Controls("Mastertab#TaetigkeitID#1") = -1


    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#LQWorkstationID#1"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_LQWorkstationErfasserID_AfterUpdate()

    Dim lrs As DAO.Recordset

    Me.Controls("Mastertab#TaetigkeitErfasserID") = -1
    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#LQWorkstationErfasserID"))
    Set lrs = Nothing

End Sub


Private Sub Mastertab_Material_Kosten_1_AfterUpdate()

    Dim lrs As DAO.Recordset

    If IsNull(Me.Controls("Mastertab#Material-Kosten#1")) Then Me.Controls("Mastertab#Material-Kosten#1") = "0,00"

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Material-Kosten#1"))
    Set lrs = Nothing

End Sub

Private Sub Meldungsnummer_AfterUpdate()
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from Mastertab where ([Meldungsnummer Kunde] = '" & Me.Controls("Mastertab#Meldungsnummer Kunde#2") & "') AND ((Mastertab.Referenznummer) Is Null);")

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        MsgBox getTC_FCFromCode("TC0247") & " " & getTC_FCFromCode("TC0246") & " " & rs("ID"), vbExclamation, getTC_FCFromCode("TC0408")
    End If

    Set rs = Nothing

End Sub

Private Sub NewDataset_Click()
On Error GoTo NewDataset_Click

Dim Reklnr As String
Dim Kürzel As String
Dim sMemo As String
Dim sOut As String
Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim NewRekl As Boolean
Dim ReklBearbeiten As Boolean
Dim sPfadDatei As String
Dim loSachbID As String


'Set db = currentdb ()

If get_pKuerzel() = "" Then
    If setKuerzel() = "" Then
        GoTo end_NewDataset_Click
    End If
End If
'Bestehedne Datensatz bearbeiten: sichern DS, Service MA speichern, usw
DoCmd.RunCommand acCmdSaveRecord
Set ldb = openAccDBConnection("openDB")
'Ermittlung ob es zu dem Vorgang schon ein Protokolleintrag vorhanden ist
Reklnr = Me.Controls("Mastertab#ID")
If p_KommentarKopieren Then
    Set rs = ldb.OpenRecordset("Select * from ReklMemo where [ID] = " & Reklnr & ";", dbOpenSnapshot)

    If rs.RecordCount Then
        If IsNull(rs("Protokoll").Value) Then
            sMemo = ""
        Else
            sMemo = rs("Protokoll").Value
        End If
        If Not IsNull(rs("Aktueller Eintrag").Value) Then
            sMemo = rs("Aktueller Eintrag").Value & Chr(13) & Chr(10) & Chr(13) & Chr(10) & sMemo
        End If
        ldb.Execute ("Update ReklMemo set [Protokoll] = '" & sMemo & _
                    "' where ID = " & Reklnr & ";")
        sMemo = ""
        ldb.Execute ("Update ReklMemo set [Aktueller Eintrag] = NULL where ID = " & Reklnr & ";")
        'db.Execute ("Update ReklMemo set [Aktueller Eintrag] = '" & sMemo & _
        '            "' where ID = " & ReklNr & ";")
    End If
End If
Set rs = Nothing
 ' Set ldb = openAccDBConnection("closeDB")

loSachbID = Me.[Zuständiger Service-MA]
Me.AllowAdditions = True
unlockDataset (Me.Controls("Mastertab#ID"))

DoCmd.GoToRecord , , acNewRec
'Das aktuelle Datum wird als Auftragsdatum eingetragen
' -> Notwendig damit die Reklamtionsnummer sofort generiert wird und für weiter Zwecke genutzt werden kann
Me.Controls("Auftragsdatum") = format(Date, "dd.mm.yy")
Me.Controls("Status") = "Bearbeitung: Service"
Me.[Zuständiger Service-MA] = loSachbID

p_ReklNr = Me.Controls("ID")
'Anlegen des Datensatz zum Sperren der Bearbeitung durch einen Zweiten Benutzer
lockDataset (p_ReklNr)
ReklBearbeiten = True
'
InsertNewReklMemoDS (p_ReklNr)
'Beim Öffnen so setzen das der generierte Kommentar nicht kopiert wird. Kopie macht erst Sinn wenn Änderungen
'am Datensatz bzw. direkt im Kommentarfeld vollzogen wurden. In dem Fall wird die Variable wie folgt geändert:
'Bei einer Änderung im Steuerelement wird die Variable auf True gesetzt (
p_KommentarKopieren = False

NewPosAufwand p_ReklNr, Me.[Zuständiger Service-MA]
Me.UF_Fehlererfassung.Requery


NewRekl = True
'Ordner für die Dokumentenablage anlegen - Ordnername = ReklNr
With Forms("Fehlererfassung - Register")
    sPfadDatei = setReklVerzeichnis(.[ID], format(.Auftragsdatum, "yyyy"))
End With

'RSP - 2012-05-30 Filter wird hier zentral gesetzt.
If ReklBearbeiten Then
    InsertNewReklMemoDS (p_ReklNr)
    InsertNewEinkaufDS (p_ReklNr)
End If
With Forms("Fehlererfassung - Register")
    'DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , , acFormEdit
    .Filter = "[Mastertab].[ID] = " & p_ReklNr
    .FilterOn = True
    .Requery
    .Refresh
    .AllowAdditions = False
    'Set Interner Fehler
    .extern = True
    '05.09.2013 -RSP: Ende
'    'DoCmd.Maximize

    'Sichern des aktuellen Vorgang-Status
    If IsNull(Forms("Fehlererfassung - Register").Controls("Kategorie-Status")) Then
        p_VorgangStatus = ""
    Else
        p_VorgangStatus = Forms("Fehlererfassung - Register").Controls("Kategorie-Status")
    End If
    Forms("Fehlererfassung - Register").Controls("Kunden-ID").SetFocus  '20121121-RSP Änderung
End With
end_NewDataset_Click:
Exit Sub

NewDataset_Click:
    If Err = 2046 Then
        Resume Next
    Else
        GoTo end_NewDataset_Click
    End If

End Sub

Private Sub Mastertab_TaetigkeitID_AfterUpdate()

    Dim lrs As DAO.Recordset

    Me.Controls("Mastertab#DefTaetigkeitID") = -1
    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#TaetigkeitID"))

    Set lrs = Nothing

End Sub

Private Sub Mastertab_Transportschaden_Click()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Transportschaden"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Umbau_Reparatur_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Umbau/Reparatur"))
    Set lrs = Nothing

    'Ticket "LZ-Rücksendung" wird benendet wenn vorhanden, ID vom Tickettyp ist 16
    If Me.Mastertab_Umbau_Reparatur = 0 Then
        Dim lID As Long
        lID = getTicketID(16, Me.Mastertab_ID)
        If lID > 0 Then
            TicketFuellen lID, "3", "TicketStatusID"
            TicketFuellen lID, Date, "ErledigtAm"
            TicketFuellen lID, "System", "ErledigtVon"
        End If
    End If

End Sub

Private Sub Mastertab_Unterbaugruppe_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Unterbaugruppe"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Verursacher_1_AfterUpdate()

    Dim lrs As DAO.Recordset

    If Me.Controls("Mastertab#Verursacher#1") <> "Lieferant" Then
        Me.Controls("ROS_Daten_Einkauf#TaetigkeitLieferantID") = -1
    End If

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Verursacher#1"))
    Set lrs = Nothing

    Call Befehl6068_Click

End Sub


Private Sub Mastertab_Verursacher_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Verursacher"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_WE_Bericht_Menge_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#WE-Bericht-Menge"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Welcher_Fehler_tritt_auf_1_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Welcher Fehler tritt auf#1"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Welcher_Fehler_tritt_auf_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Welcher Fehler tritt auf"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Wie_Warum_passierte_der_Fehler_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Wie Warum passierte der Fehler"))
    Set lrs = Nothing


End Sub



Private Sub Mastertab_Zustand_AfterUpdate()     '20191028-ROS

    Dim lrs As DAO.Recordset

    Me.Controls("Mastertab#Kategorie-Status") = "Abgeschlossen"
    Call Status_AfterUpdate(ls_AktiverStatus)
    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Kategorie-Status"))

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Zustand"))
    Set lrs = Nothing

    If Me.Mastertab_Zustand = -1 Then
        FormKoorSave Me             '20260224-ROS
        DoCmd.Close acForm, Me.Name
    End If
    '20210909-ROS Übersicht wird nur aktuallisiert wen geöffnet
    If IsLoaded("Reklamationen_ÜbersichtIntern") Then
        Forms("Reklamationen_ÜbersichtIntern").Refresh
    End If

End Sub

Private Sub Mastertab_Zuständig_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Zuständig"))
    Set lrs = Nothing


End Sub

Private Sub Mastertab_Zuständiger_Service_MA_2_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Zuständiger Service-MA#2"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Zuständiger_Service_MA_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Zuständiger Service-MA"))
    Set lrs = Nothing

End Sub

Private Sub NewPos_Click()

    Dim lKuerzel As String

    If Me.Mastertab_Kategorie_Status = "Abgeschlossen" Then Exit Sub    '20210428-ROS
    lKuerzel = Nz(DLookup("[Kürzel]", "[qryGesperrteVorgänge]", "[ID] = " & Me.Controls("Mastertab#ID") & ""))
    If lKuerzel <> get_pKuerzel Then Exit Sub

    NewPosAufwand2 Me.Controls("Mastertab#ID"), get_pKuerzel

    Me.Controls("UF_Fehlererfassung").Requery
    Me.Controls("UF_Fehlererfassung").SetFocus
    DoCmd.GoToRecord , , acLast

    Forms("Fehlererfassung - Register").Controls("lstZeitaufwändeBereich").Requery

End Sub


Private Sub Objekt_schließen___Formular_Click()
On Error GoTo err_Objekt_schließen___Formular_Click
    Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")

    Dim lRückliefKunde As Integer
    Dim ITicketID As Integer
    Dim lIDTicket As Long
    Dim lIDVorgangText As Long

    Call Befehl6068_Click

    ' Wenn Flag "Rücksendeung vom Kunden erwartet" gesetzt ist und noch kein Ticket existiert, wird ein Ticket mit dem Ticket-Typ "LZ-Rücksendung" gestartet.
    ITicketID = TicketExist(16, Me.Controls("Mastertab#ID"))
    If ITicketID = 0 Then
        lRückliefKunde = Nz(DLookup("[Umbau/Reparatur]", "Mastertab", "[ID] = " & Me.Controls("Mastertab#ID") & ""))
        If lRückliefKunde = -1 Then
            lIDTicket = TicketErstellen(16, Me.Mastertab_ID)
            TicketFuellen lIDTicket, "2", "TicketStatusID"
            TicketFuellen lIDTicket, Date, "MailDatum"
            TicketFuellen lIDTicket, "Service", "MailAn"  'Zuständigkeit liegt beim Service
        End If
    Else
    'Wenn Ticket existiert und Flag "Rücksendeung vom Kunden erwartet" nicht gesetzt ist, wird Ticket abgeschlossen.
        lRückliefKunde = Nz(DLookup("[Umbau/Reparatur]", "Mastertab", "[ID] = " & Me.Controls("Mastertab#ID") & ""))
        If lRückliefKunde = 0 Then
            TicketFuellen lIDTicket, "3", "TicketStatusID"
            TicketFuellen lIDTicket, Date, "MailDatum"
            TicketFuellen lIDTicket, "System", "ErledigtVon" 'Zuständigkeit liegt beim Service
        End If
    End If

    lIDVorgangText = Nz(DLookup("[VorgangID]", "[AufwandVorgangPositionen]", "[MID] = " & Me.Controls("Mastertab#ID") & " and [Vorgang]= 55"))

    If lIDVorgangText > 0 Then
        ldb.Execute ("UPDATE [ReklMemo] set [Lagerpüfung]=  0 where [ID] = " & Me.Controls("Mastertab#ID") & " ;")
    Else
        ldb.Execute ("UPDATE [ReklMemo] set [Lagerpüfung]=  1 where [ID] = " & Me.Controls("Mastertab#ID") & " ;")
    End If

    FormKoorSave Me
    DoCmd.Close acForm, Me.Name

    Exit Sub

err_Objekt_schließen___Formular_Click:

    If Err = 2051 Then
        Resume Next
    End If

End Sub

Private Sub optAlleTicket_Click()

    Dim sSQL As String
    Dim iStatus As Integer

    sSQL = "SELECT tblTicketListe.ReklID, tblTicketTyp.Anzeige, tblTicketStatus.Anzeige, tblTicketListe.MailDatum, tblTicketListe.BearbeitungBis, " & _
        "tblTicketListe.VerlängertBis, tblTicketListe.ErledigtAm, tblTicketListe.TicketID, tblTicketListe.TicketStatusID, tblTicketListe.MailErstellt, " & _
        "tblTicketListe.TicketTypID, tblTicketListe.MailAn, tblTicketListe.ErledigtVon " & _
        "FROM (tblTicketListe LEFT JOIN tblTicketTyp ON tblTicketListe.TicketTypID = tblTicketTyp.TicketTypID) " & _
        "LEFT JOIN tblTicketStatus ON tblTicketListe.TicketStatusID = tblTicketStatus.TicketStatusID " & _
        "WHERE (((tblTicketStatus.TicketStatusID)<>iStatus));"

    If Me.optAlleTicket = -1 Then
        sSQL = Replace(sSQL, "iStatus", 0)
    Else
        sSQL = Replace(sSQL, "iStatus", 3)
    End If


    frmUFTicketListe.Form.RecordSource = sSQL

End Sub

Private Sub PEAufwand_AfterUpdate()

    If Me.PEAufwand Mod 15 = 0 Then
        Me.Controls("Mastertab#PE-Bearbeitung") = (Me.PEAufwand \ 15) * 0.25 + Me.Controls("Mastertab#PE-Bearbeitung")
    Else
        Me.Controls("Mastertab#PE-Bearbeitung") = (Me.PEAufwand \ 15 + 1) * 0.25 + Me.Controls("Mastertab#PE-Bearbeitung")
    End If

    Me.PEAufwand = 0

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#PE-Bearbeitung"))
    Set lrs = Nothing

End Sub

Private Sub QSAufwand_AfterUpdate()

    If Me.QSAufwand Mod 15 = 0 Then
        Me.Controls("Mastertab#QS-Bearbeitung") = (Me.QSAufwand \ 15) * 0.25 + Me.Controls("Mastertab#QS-Bearbeitung")
    Else
        Me.Controls("Mastertab#QS-Bearbeitung") = (Me.QSAufwand \ 15 + 1) * 0.25 + Me.Controls("Mastertab#QS-Bearbeitung")
    End If

    Me.QSAufwand = 0

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#QS-Bearbeitung"))
    Set lrs = Nothing

End Sub

Private Sub ReklMemo_Angebotsnummer_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ReklMemo#Angebotsnummer"))
    Set lrs = Nothing

End Sub

Private Sub ReklMemo_AnlieferterminRep_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ReklMemo#AnlieferterminRep"))
    Set lrs = Nothing

End Sub

Private Sub ReklMemo_BestellnrRepAuftrag_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ReklMemo#BestellnrRepAuftrag"))
    Set lrs = Nothing

End Sub

Private Sub ReklMemo_FertigungsDatum_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ReklMemo#FertigungsDatum"))
    Set lrs = Nothing

End Sub

Private Sub ReklMemo_KundenRefNummer_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ReklMemo#KundenRefNummer"))
    Set lrs = Nothing

End Sub

Private Sub ReklMemo_Lagerpüfung_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ReklMemo#Lagerpüfung"))
    Set lrs = Nothing

End Sub

Private Sub ReklMemo_Portal_AfterUpdate()

'RSP - Änderung: 07.05.2012
Dim sLink As String
If Me.Controls("ReklMemo#Portal") Then
'    sLink = InputBox("Bitte den Link für die Reklamation eingeben.", "Eingabe Meldungsbezogener TRUMPF Portal-Link")
'    sLink = InputBox(getTC_FCFromCode("IC3853"), getTC_FCFromCode("IC3854"))
    sLink = getValueFromInputBox(getTC_FCFromCode("IC3853"), getTC_FCFromCode("IC3854"))
    If Len(sLink) Then
        Me.Controls("ReklMemo#Portal-Link") = sLink
        Me.TruTicket.Enabled = True
    Else
        MsgBox getTC_FCFromCode("TC0469") & Chr(10) & Chr(13) & getTC_FCFromCode("TC0470")
        'MsgBox "Kein Link angeggeben!" & Chr(10) & Chr(13) & "Vorgang wiederholen. Auswahl wird deaktiviert!"
        Me.Controls("ReklMemo#Portal") = False
        Me.TruTicket.Enabled = False
    End If
Else
    Me.Controls("ReklMemo#Portal") = False
    Me.TruTicket.Enabled = False
End If

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ReklMemo#Portal"))
    Set lrs = getCtrlDataset(Me, Me.Controls("ReklMemo#Portal-Link"))
    Set lrs = Nothing

End Sub

Private Sub Mastertab_Kategorie_Problem_1_AfterUpdate()
    '26.06.2014 - RSP: Erweiterung: Wird der Vorgang auf nicht berechtigt gesetzt, wird in Abhängigkeit des Verursachers die Fehlerklasse gesetzt
    '             Verursacher = leer: "Ohne Befund / Kein Fehler feststellbar"
    '             Verursacher = Kunde: "Kundenanwendungsfehler" gesetzt oder bestätigt bzw. besthender Wert beibehalten, da die FK schon das Setzen des Verursacher 'Kunde' gesetzt wird.

    Dim sMemo As String
    Dim lrs As DAO.Recordset

    sMemo = Chr(13) & Chr(10) & "Neuer Status:  " & Me.Controls("Mastertab#Kategorie-Problem#1")
    sMemo = InsertMemoDS(Me.Controls("Mastertab#ID"), sMemo)


    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Kategorie-Problem#1"))
    Set lrs = Nothing



End Sub


Private Sub RegisterStr246_Change()

Dim ReklID As Long
Dim li_caption As String
Dim li_backColor As String
Dim li_ctl As Control
Dim li_recordsource As String
Dim lb_setRS As String
Dim Test As String

    DoCmd.Hourglass True

    '1. Inhalt Dictionary sichern und anschließend leeren
    UpdatetFormTab Me, "[ID]=" & Me.Controls("Mastertab#ID")

    If IsNull(Me.Controls("Mastertab#ID")) Then
        ReklID = get_p_ReklNr
    Else
        ReklID = Me.Controls("Mastertab#ID")
    End If

    Me.Controls("cmbTicketVersenden").Visible = False
    'Main Register Fehlererfassung
    Select Case Me.RegisterStr246.Value
        Case 0
            li_caption = Me.Bezeichnungsfeld6407.Caption
            li_backColor = Me.Bezeichnungsfeld6407.BackColor
            setActivePage (0)
            lb_setRS = setRecordSource("qryDictMelder-Daten-Intern")
        Case 1
            li_caption = Me.Bezeichnungsfeld4.Caption
            li_backColor = Me.Bezeichnungsfeld4.BackColor
            setActivePage (1)
            lb_setRS = setRecordSource("qryDictMelder-Daten1")
        Case 2
            li_caption = Me.Bezeichnungsfeld24.Caption
            li_backColor = Me.Bezeichnungsfeld24.BackColor
            setActivePage (2)
            lb_setRS = setRecordSource("qryDictMelder-Daten2")
        Case 3
            li_caption = Me.Bezeichnungsfeld79.Caption
            li_backColor = Me.Bezeichnungsfeld79.BackColor
            setActivePage (3)
            lb_setRS = setRecordSource("qryDictMelder-Daten3")
        Case 4
            li_caption = Me.Bezeichnungsfeld286.Caption
            li_backColor = Me.Bezeichnungsfeld286.BackColor
            setActivePage (4)
            lb_setRS = setRecordSource("qryDictBearbeiter-Daten1")
        Case 5
            li_caption = Me.Bezeichnungsfeld289.Caption
            li_backColor = Me.Bezeichnungsfeld289.BackColor
            setActivePage (5)
            lb_setRS = setRecordSource("qryDictBearbeiter-Daten2")
            If getKurzelZuRekla = True Then Me.Controls("cmbTicketVersenden").Visible = True
        Case 6
            li_caption = Me.Bezeichnungsfeld295.Caption
            li_backColor = Me.Bezeichnungsfeld295.BackColor
            setActivePage (6)
            lb_setRS = setRecordSource("qryDictBearbeiter-Daten3")
        Case 7
            li_caption = Me.Bezeichnungsfeld292.Caption
            li_backColor = Me.Bezeichnungsfeld292.BackColor
            setActivePage (7)
            lb_setRS = setRecordSource("qryDictBearbeiter-Daten4")
        Case 8
            li_caption = Me.Bezeichnungsfeld6209.Caption
            li_backColor = Me.Bezeichnungsfeld6209.BackColor
            setActivePage (8)
            lb_setRS = setRecordSource("qryDictEinkauf-Daten2")
        Case 9
            li_caption = Me.Bezeichnungsfeld296.Caption
            li_backColor = Me.Bezeichnungsfeld296.BackColor
            setActivePage (9)
            lb_setRS = setRecordSource("qryDictKostenrechnung-Daten")
        Case 10
            li_caption = Me.Bezeichnungsfeld297.Caption
            li_backColor = Me.Bezeichnungsfeld297.BackColor
            setActivePage (10)
            lb_setRS = setRecordSource("qryDictErgebnis-Daten")
        Case 11
            li_caption = Me.RegisterStr246.Pages(11).Name
            li_backColor = Me.Bezeichnungsfeld297.BackColor
            setActivePage (11)
            lb_setRS = setRecordSource("qryDictUebersicht-Daten")
    End Select

    Me.lblRegisterTitel.Caption = li_caption
    Me.lblRegisterTitel.BackColor = li_backColor



'    Test = LoadControlsPageAfterOpen(Me, cRecordSource, Me.RegisterStr246.Pages(li_activePage).Controls, "[ID]=" & ReklID)
    Test = LoadControlsFormAfterOpen(Me, getRecordSource(), "[ID]=" & ReklID)
    Test = LoadControlsPageAfterOpen(Me, getRecordSource(), Me.RegisterStr246.Pages(getActivePage()).Controls, "[ID]=" & ReklID)

    For Each li_ctl In Me.RegisterStr246.Pages(getActivePage()).Controls
'        If li_ctl.Name = "Mastertab#Verursacher#1" Then
'            Debug.Print "Inhaltlicher Breakpoint"
'        End If
        If li_ctl.ControlType = 111 Then
            li_ctl.Requery
        End If
    Next

    Me.Mastertab_Kategorie_Status.Requery   '20191015-ROS_1 eingefügt, da beim wecheln von Reklamationen über Button "Gehe zu" der Folge-Status nicht aktuallisiert wurde
    Me.Objekt_schließen___Formular.SetFocus

    DoCmd.Hourglass False

End Sub

Private Sub ReklMappe_Click()

    Dim ReklnrKnd As String

    If IsNull(Me.Controls("Mastertab#Meldungsnummer Kunde#2")) Then
        ReklnrKnd = ""
    Else
        ReklnrKnd = Me.Controls("Mastertab#Meldungsnummer Kunde#2")
    End If

    Call VorgangsOrdner(Me.Controls("Mastertab#ID"), ReklnrKnd, Me.Controls("Mastertab#Auftragsdatum#3"), True)     '20190820-ROS

End Sub

Private Sub Rücklieferschein_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CursorWechseln(IDC_HAND)
End Sub

Private Sub Mastertab_Rückmeldung_Kunde_Soll_AfterUpdate()

    Dim lrs As DAO.Recordset

    'Beim eintragen "Rückmeldedatum Kunde" wird auch "Datum verschoben auf" eingetragen, nur wenn der Datensatz leer ist     '20190821-ROS
    If IsNull(Me.Controls("Mastertab#Rückmeldung_Kunde_Verschiebung#1")) Then
        Me.Controls("Mastertab#Rückmeldung_Kunde_Verschiebung#1") = Me.Controls("Mastertab#Rückmeldung_Kunde_Soll")
    End If

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Rückmeldung_Kunde_Soll"))
    Set lrs = Nothing

End Sub


Private Sub SendAufgAsMailQSV_Click()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")

    Set frm_MailSenden = New Form_MailSendenAnNeu
    frm_MailSenden.Visible = True
    frm_MailSenden.parentForm = Me
    Set rs = ldb.OpenRecordset("Select * from qryAufgaben where [ID] = " & Me.Controls("Mastertab#ID") & ";")
    frm_MailSenden.loadForm rs ' Hier werden die Daten via Recordset übergeben

End Sub


Private Sub ReklMemo_VorgangstypID_AfterUpdate()

    Dim lrs As DAO.Recordset

    Me.Controls("ReklMemo#VorgangstypID#1") = Me.Controls("ReklMemo#VorgangstypID")

    Set lrs = getCtrlDataset(Me, Me.Controls("ReklMemo#VorgangstypID"))
    Set lrs = Nothing

End Sub

Private Sub ReklMemo_WABestellnummer_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ReklMemo#WABestellnummer"))
    Set lrs = Nothing

End Sub

Private Sub ReklMemo_WALieferschein_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ReklMemo#WALieferschein"))
    Set lrs = Nothing

End Sub

Private Sub ReklMemo_WAWE_AB_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ReklMemo#WAWE-AB"))
    Set lrs = Nothing

End Sub

Private Sub ReklMemo_WEBestellnummer_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ReklMemo#WEBestellnummer"))
    Set lrs = Nothing

End Sub

Private Sub ReklMemo_WELieferschein_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ReklMemo#WELieferschein"))
    Set lrs = Nothing

End Sub

Private Sub ROS_Daten_Einkauf_AnsprechpLieferant_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ROS_Daten_Einkauf#AnsprechpLieferant"))
    Set lrs = Nothing

End Sub

Private Sub ROS_Daten_Einkauf_Fazit_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ROS_Daten_Einkauf#Fazit"))
    Set lrs = Nothing

End Sub

Private Sub ROS_Daten_Einkauf_NichtabgeschlosseneMaßnahmendesLieferanten_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ROS_Daten_Einkauf#NichtabgeschlosseneMaßnahmendesLieferanten"))
    Set lrs = Nothing

End Sub

Private Sub ROS_Daten_Einkauf_Prüfzyklus_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ROS_Daten_Einkauf#Prüfzyklus"))
    Set lrs = Nothing

End Sub

Private Sub ROS_Daten_Einkauf_TaetigkeitLieferantID_1_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ROS_Daten_Einkauf#TaetigkeitLieferantID#1"))
    Set lrs = Nothing

End Sub

Private Sub ROS_Daten_Einkauf_TaetigkeitLieferantID_AfterUpdate()

    Dim lrs As DAO.Recordset

    Me.Controls("Mastertab#DefTaetigkeitAnalyseID") = -1
    If IsNull(Me.Controls("ROS_Daten_Einkauf#TaetigkeitLieferantID")) Then
        Me.Controls("ROS_Daten_Einkauf#TaetigkeitLieferantID") = -1
    End If
    Set lrs = getCtrlDataset(Me, Me.Controls("ROS_Daten_Einkauf#TaetigkeitLieferantID"))

    Set lrs = Nothing


End Sub

Private Sub ROS_Daten_Einkauf_ÜberprüfungsterminErgebnis_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ROS_Daten_Einkauf#ÜberprüfungsterminErgebnis"))
    Set lrs = Nothing

End Sub

Private Sub ROS_Daten_Einkauf_ÜberprüfungsterminLQ_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ROS_Daten_Einkauf#ÜberprüfungsterminLQ"))
    Set lrs = Nothing

End Sub

Private Sub ROS_Daten_Einkauf_Wiederholungsfehler_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ROS_Daten_Einkauf#Wiederholungsfehler"))
    Set lrs = Nothing

End Sub

Private Sub ROS_Daten_Einkauf_ZeitaufwandjePrüfung_AfterUpdate()

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ROS_Daten_Einkauf#ZeitaufwandjePrüfung"))
    Set lrs = Nothing

End Sub

Private Sub SendEmailGutschrift_Click()

Dim myOutlApp As Object
Dim myMail As Object
Dim ldb As DAO.Database
Dim rs As DAO.Recordset


Set myOutlApp = CreateObject("Outlook.Application")
Set myMail = myOutlApp.CreateItem(olMailItem)
Set ldb = openAccDBConnection("openDB")
Set rs = ldb.OpenRecordset("Select * from LQEinheit where [ID]=" & Me.Controls("Mastertab#LQEinheitID#2") & ";", dbOpenSnapshot) '20201027-ROS

    With myMail
            .To = rs("MailTOLQEinheit").Value
            .CC = getTC_FCFromCode("TC0510")        'service@de.lq-group.com
            .subject = getTC_FCFromCode("TC0427", ldb) & " " & Me.Controls("Mastertab#ID").Value

            '.HTMLBody = "Dein Text in HTML-Format"
            'oder
            .body = getTC_FCFromCode("TC0424", ldb) & ", " & Chr(10) & Chr(10) & getTC_FCFromCode("TC0425", ldb) & " " & Me.Controls("Mastertab#ID") & " " & getTC_FCFromCode("TC0426", ldb) & "." & Chr(10) & Chr(10) & "Gruß" & Chr(10) & "Einkauf"

            'ggf. Deine Anlage
            '.Attachments.Add Anlage_Datei

            'Über andere Email-Adresse/Konto senden...(im Auftrag von)
            '.SentOnBehalfOfName = "DeineAndereEmail@domäne.de"

            'Erst nur Anzeigen...
            '.Display

            'sonst gleich senden...
            .send
    End With

    MsgBox getTC_FCFromCode("TC0428", ldb)

    Set myMail = Nothing
    Set myOutlApp = Nothing
    Set rs = Nothing
     ' Set ldb = openAccDBConnection("closeDB")

End Sub
Private Sub SendXDRanLiefMail_Click()

On Error GoTo Err_Befehl10_Click



    DoCmd.OpenForm "MailSendenAn"

    Exit Sub

Exit_Befehl10_Click:

'    Set rs = Nothing
     ' Set ldb = openAccDBConnection("closeDB")
    Exit Sub

Err_Befehl10_Click:
    MsgBox Err.Description 'FC0136
    Resume Exit_Befehl10_Click

End Sub


Private Sub Status_AfterUpdate(AktiverStatus As String)

Dim sMemo As String
Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim Reklnr As String


 Set ldb = openAccDBConnection("openDB")
'Ermittlung ob es zu dem Vorgang schon ein Protokolleintrag vorhanden ist
Reklnr = Forms("Fehlererfassung - Register").Controls("Mastertab#ID")
Set rs = ldb.OpenRecordset("Select * from ReklMemo where [ID] = " & Reklnr & ";", dbOpenSnapshot)


If rs.RecordCount Then
    If IsNull(rs("Aktueller Eintrag").Value) Then
        sMemo = ""
    Else
        sMemo = rs("Aktueller Eintrag").Value
    End If
    sMemo = sMemo & Chr(13) & Chr(10)
    If Not (Me.Controls("Mastertab#Kategorie-Status") = p_VorgangStatus) Then
        sMemo = sMemo & getTC_FCFromCode("TC0429", ldb) & "  " & """" & AktiverStatus & """" & " -> " & """" & Me.Controls("Mastertab#Kategorie-Status") & """"
    End If
    ldb.Execute ("Update ReklMemo set [Aktueller Eintrag] = '" & sMemo & _
                "' where ID = " & Reklnr & ";")
End If
p_KommentarKopieren = True
DoCmd.RunCommand acCmdSaveRecord

 ' Set ldb = openAccDBConnection("closeDB")
Me.Refresh
Me.Requery

Set rs = Nothing

End Sub


Private Sub Mastertab_Kategorie_Status_Change()   '20170220-ROS / 20221220-ROS Anpassung Me.[Mastertab#Kategorie-Status]durchgeführt

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim rs1 As DAO.Recordset
Dim liHour As Integer
Dim liDays As Integer
Dim liMinute As Integer
Dim lNow As Date
Dim lDate2h As Date
Dim lKatStatus As String
Dim lExtern As Boolean
Dim lIDTicket As Long

    'Wird zum Zwischenspeichen des Kategorie-Status benötigt
    lKatStatus = DLookup("[Kategorie-Status]", "Mastertab", "[ID] = " & Me.Controls("Mastertab#ID") & "")

    #If SystemDebug = 1 Then
        'Setzen von einem fiktiven Datums- und Zeitwert zu Testzwecken, um das System auszutricksen
        lNow = InputBox("Datum / Urzeit (jjjj.mm.dd hh:mm:ss)", "Eingabe Datum / Urzeit")
    #Else
        lNow = Now
        'Aufrunden Sekunde auf die nächste Minute, um beim Statuswechsel min. 1 Minute Berarbeitungszeitraum eintragen zu können, ansonsten bleibt die Differenz
        '0 wenn der Statuswechsel in der gleichen Minute erfolgt.
        'Biei 59 Minuten muss die nächste Stunde ebenso hochgezählt werden
        If MID(lNow, 15, 2) = "59" Then
            'Inkrement Minute und Stunde
            lNow = Left(lNow, 11) & format(CInt(MID(lNow, 12, 2)) + 1, "##") & ":00:00"
        Else
            'Inkrement Minute
            lNow = Left(lNow, 14) & format(CInt(MID(lNow, 15, 2)) + 1, "##") & ":00"
        End If
    #End If

    'Hier wird der letzt Status, chronologisch geordnet ermittelt.
    'Dem Status wird die Dauer seines Zustandes bis zum aktuellen Statuswechsel eingetragen.
    'Zeitstempel / Dauer  ist der aktuelle Zeitpunkt des Statuswechsel - dem Zeitstempel in dem der letzte Status eingetragen wurde.
    Set ldb = openAccDBConnection("openDB")
    Set rs1 = ldb.OpenRecordset("SELECT * from tblStatusWechsel WHERE [ReklNr] = " & Me.Controls("Mastertab#ID") & " ORDER BY tblStatusWechsel.Datum;")
    If rs1.RecordCount > 0 Then
        rs1.MoveLast
        rs1.Edit
        rs1!Kürzel = get_pKuerzel()
        'Berechnung der Dauer in Minuten in dem der letzt Status aktiv war.
        'Die Minuten beziehen sich nur auf die Arbeitszeit, unter Berücksichtigung des Wochenende
        rs1!AT = CalcArbeitszeitDiffDate(rs1!Datum, lNow) 'Eintrag lokale Variable
        'Die berechneten Minuten werden in Arbeitstage, Stunden und Minuten zerlegt
        rs1!AT = getHourMinuteByAnzMinute(rs1!AT, liDays, liHour, liMinute)
        rs1!Zeit_AT = liDays
        rs1!Zeit_h = liHour
        rs1!Zeit_min = liMinute
        rs1.Update
'        Me.Refresh
'        Me.Requery
        'Eintrag des neuen Status mit aktuellem Zeitstempel.
        'Als Zusatzinformation wird der Vorgänger-Status mit eingetragen. Hat aber keine Auswirkung
        rs1.AddNew
        rs1!Reklnr = Me.Controls("Mastertab#ID")
        rs1!Status = Me.[Mastertab#Kategorie-Status]
        rs1!Datum = lNow
        rs1!Kürzel = get_pKuerzel()
        rs1!VorgaengerStatus = gsCurrState
        rs1!Kürzel = get_pKuerzel()
        rs1.Update
'        Me.Refresh
'        Me.Requery

    Else
'        'Nur bei alten Vorgängen relevant!
        rs1.AddNew
        rs1!Reklnr = Me.Controls("Mastertab#ID")
        rs1!Status = Me.[Mastertab#Kategorie-Status]
        rs1!Datum = lNow
        rs1!Kürzel = get_pKuerzel()
        rs1!VorgaengerStatus = gsCurrState
        rs1!Kürzel = get_pKuerzel()
        rs1!AT = CalcArbeitszeitDiffDate(rs1!Datum, lNow)
        rs1!AT = getHourMinuteByAnzMinute(rs1!AT, liDays, liHour, liMinute)
        rs1!Zeit_AT = liDays
        rs1!Zeit_h = liHour
        rs1!Zeit_min = liMinute
        rs1.Update
        Me.Refresh
        Me.Requery


    End If
    Set rs = Nothing
    Set rs1 = Nothing

    lExtern = Nz(DLookup("[extern]", "Mastertab", "[ID] = " & Me.Controls("Mastertab#ID") & ""))
    'Aktionen nur duchführen, wenn es ein Externer Vorgang ist, nicht für Interne Vorgänge vorgesehen
    If lExtern = True Then
        Select Case Me.Mastertab_Kategorie_Status
            Case "Stellungnahme Klärung Kunde"
                'Bei Wechsel in Status "Stellungnhame Klärung Kunde"
                'Übernahme Stellungnahme Ticket "LZ-Übernahme Kunde" erstellen
                lIDTicket = TicketErstellen(17, Me.Mastertab_ID)
                TicketFuellen lIDTicket, "2", "TicketStatusID"
                TicketFuellen lIDTicket, Date, "MailDatum"
                TicketFuellen lIDTicket, "Service", "MailAn"  'Zuständigkeit liegt beim Service

            Case "Abgeschlossen"
                'Alle offene Tickets abschließen
                Dim rsTicket As DAO.Recordset
                Set rsTicket = ldb.OpenRecordset("Select * from tblTicketListe where [ReklID] = " & Me.Mastertab_ID & " and [TicketStatusID] = 2 ;")
                If rsTicket.RecordCount Then
                    rsTicket.MoveFirst
                    Do While Not rsTicket.EOF
                        TicketFuellen rsTicket("TicketID"), "3", "TicketStatusID"
                        TicketFuellen rsTicket("TicketID"), Date, "ErledigtAm"
                        TicketFuellen rsTicket("TicketID"), "System", "ErledigtVon"
                        rsTicket.MoveNext
                    Loop
                End If
                Set rsTicket = Nothing

            Case "Rücklieferung Kunde"
                'Automatisch Flag setzten "Rücklieferung Kunde erwartet"
                ldb.Execute ("Update Mastertab set [Umbau/Reparatur] = 1 where ID = " & Me.Controls("Mastertab#ID") & ";")

            Case "Klärung: zurück Kunde"
                'Bei Wechsel in Status "Klärung: zurück Kunde"
                'Vorgang Klärung Kunde Ticket "LZ-Fragen an Kunden" erstellen
                lIDTicket = TicketErstellen(15, Me.Mastertab_ID)
                TicketFuellen lIDTicket, "2", "TicketStatusID"
                TicketFuellen lIDTicket, Date, "MailDatum"
                TicketFuellen lIDTicket, "Service", "MailAn"  'Zuständigkeit liegt beim Service

            Case "Gutschrift / Ersatz Lieferant"
                'Bei Wechsel in Status "Gutschrift / Ersatz Lieferant"
                'Vorgang Klärung Kunde Ticket "LZ-Kostenausgleich" erstellen ID18
                lIDTicket = TicketErstellen(18, Me.Mastertab_ID)
                TicketFuellen lIDTicket, "2", "TicketStatusID"
                TicketFuellen lIDTicket, Date, "MailDatum"
                TicketFuellen lIDTicket, "Service", "MailAn"  'Zuständigkeit liegt beim Service

            Case Else
                'Bei Wechsel zurück von Status "Stellungnhame Klärung Kunde" auf anderen Status
                'Übernahme Stellungnahme Ticket "LZ-Übernahme Kunde" abschliessen
                If lKatStatus = "Stellungnahme Klärung Kunde" Then
                    lIDTicket = getTicketID(17, Me.Mastertab_ID)
                    TicketFuellen lIDTicket, "3", "TicketStatusID"
                    TicketFuellen lIDTicket, Date, "ErledigtAm"
                    TicketFuellen lIDTicket, "System", "ErledigtVon"
                End If
                'Bei Wechsel zurück von Status "Klärung: zurück Kunde" auf anderen Status
                'Vorgang Klärung Kunde Ticket "LZ-Fragen an Kunden" abschliessen
                If lKatStatus = "Klärung: zurück Kunde" Then
                    lIDTicket = getTicketID(15, Me.Mastertab_ID)
                    TicketFuellen lIDTicket, "3", "TicketStatusID"
                    TicketFuellen lIDTicket, Date, "ErledigtAm"
                    TicketFuellen lIDTicket, "System", "ErledigtVon"
                End If

                'Bei Wechsel zurück von Status "Gutschrift / Ersatz Lieferant" auf anderen Status
                'Vorgang Klärung Kunde Ticket "LZ-Kostenausgleich" abschliessen
                If lKatStatus = "Klärung: zurück Kunde" Then
                    lIDTicket = getTicketID(18, Me.Mastertab_ID)
                    TicketFuellen lIDTicket, "3", "TicketStatusID"
                    TicketFuellen lIDTicket, Date, "ErledigtAm"
                    TicketFuellen lIDTicket, "System", "ErledigtVon"
                End If

        End Select
    End If

End Sub


Private Sub Mastertab_Kategorie_Status_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    gsCurrState = Me.Controls("Mastertab#Kategorie-Status")
    ls_AktiverStatus = Me.Controls("txtAktiverStatus")
'    Debug.Print ls_AktiverStatus

End Sub



Private Sub Text6109_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

        Call CursorWechseln(IDC_NO)

End Sub

Private Sub Text6602_DblClick(Cancel As Integer)

    Call LagerübersichtArtikel(Me.Text6602, "00000")

End Sub

Private Sub Text6604_DblClick(Cancel As Integer)

    Call LagerübersichtArtikel(Left(Me.Text6604, InStr(Me.Text6604, "-") - 1), Right(Me.Text6604, Len(Me.Text6604) - InStrRev(Me.Text6604, "-")))

End Sub

Private Sub TruTicket_Click()
On Error GoTo err_TruTicket_Click

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim objIE1 As Object
    Dim ChromeLocation As String
    Dim lb_FirstErr As Boolean
    Dim lUser As String
    Dim lComputername As String

    lUser = fOSUserName
    lComputername = Environ("COMPUTERNAME")

    lb_FirstErr = False

    If Not IsNull(Me.Controls("ReklMemo#Portal-Link")) And Not Me.Controls("ReklMemo#Portal-Link") = "" Then
        If lComputername = "LQTERMINAL02" Then
            ChromeLocation = "C:\Users\" & lUser & "\AppData\Local\Google\Chrome\Application\chrome.exe"
        Else
            ChromeLocation = "C:\Program Files\Google\Chrome\Application\chrome.exe"  '20220404-ROS: IE wird ab Sommer 2022 nicht mehr Unterstützt Umstieg auf Chrome
        End If
        shell (ChromeLocation & " -url -newtab " & Me.Controls("ReklMemo#Portal-Link").Value), vbNormalFocus
'        Set objIE1 = CreateObject("InternetExplorer.Application")
'        objIE1.Visible = True   ' Objekt für IE Fenster 1
'        objIE1.Navigate Me.Controls("ReklMemo#Portal-Link").Value
    Else
        MsgBox getTC_FCFromCode("TC0253", ldb)
    End If

    Exit Sub

err_TruTicket_Click:

    If Err = 53 And Not lb_FirstErr Then
        ChromeLocation = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"  '20220404-ROS: IE wird ab Sommer 2022 nicht mehr Unterstützt Umstieg auf Chrome
        lb_FirstErr = True
        Resume
    Else
        MsgBox getTC_FCFromCode("TC0253", ldb) & " - Fehlertext ergänzen !!"
    End If

End Sub


Private Sub Mastertab_Fertigungsauftrag_FA_1_AfterUpdate() '20180125-ROS

    Dim Response As Variant
    Dim s As String

    If Not (IsNull(Me.Controls("Mastertab#Fertigungsauftrag-FA#1")) Or Me.Controls("Mastertab#Fertigungsauftrag-FA#1") = "") Then
        'RSP - 20140222: Suchen aller Fertigungsaufträge in Avista anhand der Eingab eines Teilstrings
        Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
        Dim rsADO As ADODB.Recordset
        Dim iLiefAvistaID As Long
        'Avista DB Provider Instanz erzeugen
        Set adbprov = New clsAvistaProvider
        Response = acDataErrAdded
        'MsgBox ("Der eingegebene Wert: " & NewData)
        'Generierung Abfrage
        s = SelectFA2 & "'*" & Me.Controls("Mastertab#Fertigungsauftrag-FA#1") & "'"
        'MsgBox s
        Set rsADO = adbprov.getRecordsetAvista(SelectFA2 & "'%" & Me.Controls("Mastertab#Fertigungsauftrag-FA#1") & "'")
        If rsADO.RecordCount Then
            Me.Controls("ReklMemo#FertigungsDatum") = rsADO("DATUM7").Value                 '20190822-ROS_3
'            Me.Controls("Mastertab#Mitarbeiter-Kurzzeichen#1") = rsADO("NAME4").Value
        End If
        Response = acDataErrContinue
        rsADO.Close
        Set rsADO = Nothing

        Set adbprov = Nothing
    Else

    End If

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Fertigungsauftrag-FA#1"))

    Set lrs = Nothing


ende:
    Exit Sub

Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0267
    Resume ende

End Sub

Private Sub txtArtikelAvistaTyp_AfterUpdate()

    Dim i As Integer
    i = getAnzeigefromTable("Artikel-all_avista_out", "Artikel-ID", Me.Controls("Mastertab#Artikel-ID 1#5"), "AvistaArtikelTyp")
    If i = 1 Then Me.Controls("txtAvistaArtikelTyp") = "Leitersatz"

End Sub

Private Sub txtEMailAblegen_AfterUpdate()       '20190607-ROS
'E-Mail abspeichern per drag and drop in Textfeld
On Error GoTo err_Speichern

    Dim olApp As Outlook.Application
    Dim olExp As Outlook.Explorer
    Dim olSel As Outlook.Selection
    Dim Betreff As String
    Dim i As Integer
    Dim A As String
    Dim j As Integer
    Dim Email As String
    Dim Eingang As String
    Dim Anhang As String
    Dim strPfad As String

    Dim oAnhang As Object
    Dim oMail As Object
    Dim Att As Outlook.Attachment

    If Me.txtEMailAblegen <> "" Then
        Set olApp = GetObject("", "Outlook.Application")
        Set olExp = olApp.ActiveExplorer
        Set olSel = olExp.Selection

        'Pfad für E-Mail Ablage
'        strpfad = get_DateiPfad
        If IsNull(Me.Controls("Mastertab#Meldungsnummer Kunde#2")) Or Me.Controls("Mastertab#Meldungsnummer Kunde#2") = "" Then
            strPfad = setReklVerzeichnis(Me.Controls("Mastertab#ID"), format(Me.Controls("Mastertab#Auftragsdatum#3"), "yyyy"))
        Else
            strPfad = setReklVerzeichnis(Me.Controls("Mastertab#ID"), format(Me.Controls("Mastertab#Auftragsdatum#3"), "yyyy"), Me.Controls("Mastertab#Meldungsnummer Kunde#2"))
        End If

        strPfad = strPfad & "E-Mail" & "\"

        'E-Mail abspeichern und Betreff ändern wenn nötig
        For i = 1 To olSel.Count
            Eingang = format(olSel.Item(i).ReceivedTime, "yyyymmdd_hhmm")
            Betreff = olSel.Item(i).subject
'            Betreff = Replace(Betreff, ":", "_")  ' ersetzten von ":", "\", "/" im Betreff, da sonst eine Speicherung nicht möglich ist
'            Betreff = Replace(Betreff, "\", "_")
'            Betreff = Replace(Betreff, "/", "_")
'            Betreff = Replace(Betreff, "?", "_")
'            Betreff = Replace(Betreff, "->", "_")
'            Betreff = Replace(Betreff, "->", "_")
            Betreff = SonderzeichenEntfernen(Betreff)   '20190902-ROS entfernt alle Sonderzeichen

            Email = Dir(strPfad & Eingang & "-" & get_pKuerzel & "_" & Betreff & ".msg")   'Prüfen ob Dateipfad schon vorhanden ist
            If Email = "" Then
                olSel.Item(i).SaveAs strPfad & Eingang & "-" & get_pKuerzel & "_" & Betreff & ".msg", olMSG
            Else
                If MsgBox(getTC_FCFromCode("TC3916"), vbYesNo, getTC_FCFromCode("TC3917")) = vbYes Then
'                If MsgBox("E-Mail schon vorhanden, E-Mail abspeichern?", vbYesNo) = vbYes Then
'                    a = InputBox(Betreff, "Bitte Betreff ergänzen")
'                    A = InputBox(getTC_FCFromCode("IC3918") & " " & Betreff, getTC_FCFromCode("IC3919"))
                    A = getValueFromInputBox(getTC_FCFromCode("IC3918") & " " & Betreff, getTC_FCFromCode("IC3919"))
                    If A = "" Then Exit Sub
                    olSel.Item(i).SaveAs strPfad & Eingang & "-" & get_pKuerzel & "_" & A & "-" & Betreff & ".msg", olMSG  'wenn ja wird der Betreff geändert Variable a wird eingefügt
                Else
                    Exit Sub
                End If
            End If
            'Anhänge der gespeicherten E-Mail in Vorgangsordner ablegen bzw. speichern
            Set oMail = olSel.Item(i)
            With oMail.Attachments
                j = .Count
                Do While (j > 0)
                Set oAnhang = .Item(j)
                    Anhang = Dir(get_DateiPfad & oAnhang.DisplayName)   'Prüfen ob Anhang schon vorhanden
                    If Anhang = "" Then
'                        If oAnhang.FileName Like "*msg" Then
'                            oAnhang.SaveAsFile get_DateiPfad & oAnhang.FileName
'                        Else
'                            oAnhang.SaveAsFile get_DateiPfad & oAnhang.DisplayName
                            oAnhang.SaveAsFile get_DateiPfad & oAnhang.FileName
'                        End If
                    End If
                j = j - 1
                Loop
            End With
        Next i

        Set olApp = Nothing
        Set olExp = Nothing
        Set olSel = Nothing
    End If
    Me.txtEMailAblegen = ""
    Exit Sub

err_Speichern:
If Err = 52 Then
    Resume Next
Else
'    MsgBox "Es konnten nicht alle Daten abgelegt werden, bitte Daten Prüfen"
    MsgBox getTC_FCFromCode("TC4225"), vbInformation, getTC_FCFromCode("TC4226")
End If

End Sub

Private Sub txtEMailAblegen_Click() '20190607-ROS

    Me.txtAktiverStatus.SetFocus

End Sub


Private Sub txtEMailAblegen_Exit(Cancel As Integer) '20190607-ROS

'    Call txtEMailAblegen_AfterUpdate

End Sub


Private Sub txtEMailAblegen_KeyDown(KeyCode As Integer, Shift As Integer) '20190607-ROS

    Me.txtAktiverStatus.SetFocus

End Sub


Private Sub ReklMemo_Vorgangs_Eingang_AfterUpdate()


Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim rs1 As DAO.Recordset
Dim liHour, liDays, liMinute As Integer

    Set ldb = openAccDBConnection("openDB")
    'RSP_20240214
    Set rs1 = ldb.OpenRecordset("SELECT * from tblStatusWechsel WHERE [ReklNr] = " & Me.Controls("Mastertab#ID") & " and [Status] =  'Bearbeitung: Service' ORDER BY [Datum];")
    'Wenn Zustand 'Vorgang Eingang' schon vorhanden, wird die vorhandene Zeit mit der neue eingegebenen Zeit korriegiert
    If rs1.RecordCount > 0 Then
        rs1.Edit
        rs1!Datum = Me.Controls("ReklMemo#Vorgangs_Eingang")
        rs1!Kürzel = get_pKuerzel()
        rs1!AT = CalcArbeitszeitDiffDate(rs1!Datum, Me.Controls("Mastertab#Datum#1"))
        rs1!AT = getHourMinuteByAnzMinute(rs1!AT, liDays, liHour, liMinute)
        rs1!Zeit_AT = liDays
        rs1!Zeit_h = liHour
        rs1!Zeit_min = liMinute
        rs1.Update
        Me.Refresh
        Me.Requery

    Else
        'Eintrag 'Vorgang Eingang' wird neue berechnet und anhand der eingebenen Daten erzeugt
        rs1.AddNew
        rs1!Reklnr = Me.Controls("Mastertab#ID")
        rs1!Status = "Bearbeitung: Service"
1100        rs1!Datum = Me.Controls("ReklMemo#Vorgangs_Eingang")
        rs1!Kürzel = get_pKuerzel()
        rs1!AT = CalcArbeitszeitDiffDate(rs1!Datum, Me.Controls("Mastertab#Datum#1"))
'        rs1!AT = CalcArbeitszeitDiffDate(rs1!Datum, DateAdd("h", 2, rs1!Datum)) 'Diff = Vorgang-Eingang  und  Erfassungsdatum Mastertab
        rs1!AT = getHourMinuteByAnzMinute(rs1!AT, liDays, liHour, liMinute)
        rs1!Zeit_AT = liDays
        rs1!Zeit_h = liHour
        rs1!Zeit_min = liMinute
        rs1.Update
        Me.Refresh
        Me.Requery

    End If

    Dim lrs As DAO.Recordset

    Set lrs = getCtrlDataset(Me, Me.Controls("ReklMemo#Vorgangs_Eingang"))
    Set lrs = Nothing

    Set rs = Nothing
    Set rs1 = Nothing

    'Laufzeit-Ticket "LZ-Vorgang" erstellen
    Dim lTicketID As Integer

    lTicketID = TicketErstellen(14, Me.Mastertab_ID)
    TicketFuellen lTicketID, Date, "MailDatum"
    TicketFuellen lTicketID, "2", "TicketStatusID"
    TicketFuellen lTicketID, "Service", "MailAn"

End Sub



Private Sub VorgangAktivieren_Click()

Dim lID As Long
Dim iID As Long
Dim i As Boolean
Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")
    ldb.Execute ("Update Mastertab set [Status: erledigt] = 0 where ID = " & Me.Controls("Mastertab#ID") & ";")

    '20190828-ROS
    'Abgeschlossene Interne Reklamationen wieder öffnen ohne Passwort Eingabe
    '###################################################################################################################
    If Me.Controls("Mastertab#Kategorie-Status").Value = "Abgeschlossen" And Me.Controls("Mastertab#intern") = True Then
        iID = Me.Controls("Mastertab#ID")
        DoCmd.Close acForm, "Fehlererfassung - Register", acSaveNo
        Application.Echo False
        DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab#ID] = " & iID, acFormEdit
        lockDataset (iID)
        i = InsertNewReklMemoDS(iID, "Vorgang im abgeschlossenen Zustand in Bearbeitung genommen")
    Else
    '###################################################################################################################
        If Me.Controls("Mastertab#Kategorie-Status").Value = "Abgeschlossen" Then
            If get_ServiceZuordnung = True Then        '20170125-ROS
                iID = Me.Controls("Mastertab#ID")
                ldb.Execute ("Update Mastertab set [Kategorie-Status] = 'Bearbeitung: Service' where ID = " & iID & ";")

                'Nach öffnen des Vorgang, neue Laufzeitticket, LZ-Vorgang erstellen
                lID = TicketErstellen(14, iID)
                TicketFuellen lID, Date, "MailDatum"
                TicketFuellen lID, "2", "TicketStatusID"
                TicketFuellen lID, "Service", "MailAn"  'Zuständigkeit liegt beim Service

                DoCmd.Close acForm, "Fehlererfassung - Register", acSaveNo
                Application.Echo False
                DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab#ID] = " & iID, acFormEdit
                lockDataset (iID)
                i = InsertNewReklMemoDS(iID, "Vorgang im abgeschlossenen Zustand in Bearbeitung genommen")
            Else
'                If Not InputBox(getTC_FCFromCode("TC0430") & ": ", "Append DB") = getPWDBAdmin() Then
                If Not getValueFromInputBox(getTC_FCFromCode("TC0430") & ": ", "Append DB") = getPWDBAdmin() Then
                    MsgBox getTC_FCFromCode("TC0204")  'MsgBox ("Falsches Password!") 'TC0254
                    Application.Echo True
                    Exit Sub
                Else
                    iID = Me.Controls("Mastertab#ID")
                    ldb.Execute ("Update Mastertab set [Kategorie-Status] = 'Bearbeitung: Service' where ID = " & iID & ";")

                    'Nach öffnen des Vorgang, neue Laufzeitticket, LZ-Vorgang erstellen
                    lID = TicketErstellen(14, iID)
                    TicketFuellen lID, Date, "MailDatum"
                    TicketFuellen lID, "2", "TicketStatusID"
                    TicketFuellen lID, "Service", "MailAn"  'Zuständigkeit liegt beim Service

                    DoCmd.Close acForm, "Fehlererfassung - Register", acSaveNo
                    Application.Echo False
                    DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[Mastertab#ID] = " & iID, acFormEdit
                    lockDataset (iID)
                    i = InsertNewReklMemoDS(iID, "Vorgang im abgeschlossenen Zustand in Bearbeitung genommen")
                End If
            End If
        End If
    End If

    Application.Echo True


End Sub


Private Sub Mastertab_WE_Datum_AfterUpdate()

    Dim lrs As DAO.Recordset
    Dim ITicketID As Integer

    Me.Controls("Mastertab#Umbau/Reparatur") = 0

    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#WE-Datum"))
    Set lrs = getCtrlDataset(Me, Me.Controls("Mastertab#Umbau/Reparatur"))
    Set lrs = Nothing

    If Not IsNull(Me.Controls("Mastertab#WE-Datum")) Then
     ' Wenn Ticket-Typ "LZ-Rücksendung" existiert, dann beneden
        ITicketID = TicketExist(16, Me.Controls("Mastertab#ID"))
        If ITicketID > 0 Then
            Dim lIDTicket As Long
            lIDTicket = getTicketID(16, Me.Controls("Mastertab#ID"))
            TicketFuellen lIDTicket, "3", "TicketStatusID"
            TicketFuellen lIDTicket, Me.Controls("Mastertab#WE-Datum"), "ErledigtAm"
            TicketFuellen lIDTicket, "System", "ErledigtVon"
        End If
    End If

End Sub

Private Sub ReklMemo_Aktueller_Eintrag_DblClick(Cancel As Integer)

    DoCmd.OpenForm "FehlerRecherchePE_PopUp", acNormal, , , , acDialog
    'Me.Requery

End Sub


Private Sub ReklMemo_Aktueller_Eintrag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CursorWechseln(IDC_HAND)
End Sub

Private Sub XDReportLief_Click()

    Dim lVerursacher As String
    Dim lLief As Integer

    lVerursacher = Nz(DLookup("[Verursacher]", "Mastertab", "[ID] = " & get_p_ReklNr() & ""))
    lLief = Nz(DLookup("[Lieferanten-Code 1]", "Mastertab", "[ID] = " & get_p_ReklNr() & ""))

    If lVerursacher <> "Lieferant" Then
        MsgBox getTC_FCFromCode("TC4194"), vbInformation, getTC_FCFromCode("TC4195")
        Exit Sub
    End If

    If lLief = 0 Then
        MsgBox getTC_FCFromCode("TC4196"), vbInformation, getTC_FCFromCode("TC4197")
        Exit Sub
    End If

    '1. Inhalt Dictionary sichern und anschließend leeren
    UpdatetFormTab Me, "[ID]=" & Me.Controls("Mastertab#ID")

    DoCmd.OpenForm "DetailinformationenReklamationLief", , , , , acDialog

End Sub

Public Sub XDReportLief_DblClick(Cancel As Integer)
'##################################################################################################
'SUB: XDReportLief_DblClick
'Funktionsbeschreibung:
'Füllt das 3D-Report Template aus und kopiert dabei eine Kopie in den entsprechenden Vorgansordner
'
'##################################################################################################
On Error GoTo err_XDReportLief_DblClick

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsLief As DAO.Recordset
    Dim rs2 As DAO.Recordset
    Dim fs As Object
    Dim wb As Excel.Workbook
    Dim WS As Excel.Worksheet
    Dim objExcel As Object
    Dim lexcelprov As clsExcelProvider
    Dim bReportOeffnen As Boolean
    Dim sPfadDatei As String
    Dim bIntern As Boolean '20150825 - RSPTag
    Dim rsAufwand As DAO.Recordset '20150825 - RSPTag
    Dim GesamtZeit As String '20160224-ROS

    Dim XDReportPfad As String  '20160304-ROS
    XDReportPfad = getTC_FCFromCode("DP0001")

    Dim XDReportTemplate As String  '20160304-ROS
    XDReportTemplate = getTC_FCFromCode("DP0002")

    Dim XDReportTemplateLieferant As String  '20160304-ROS
    'Abfrage ob Reklamationsbericht in Englisch erstellt werden soll 20230419-ROS
    If Cancel = 2 Then
        XDReportTemplateLieferant = getTC_FCFromCode("DP0014")
    Else
        XDReportTemplateLieferant = getTC_FCFromCode("DP0003")
    End If

    'Init
    bReportOeffnen = False

    'Plausibilitätsprüfung
    If IsNull(Me.Controls("Mastertab#Sachbearbeiter#1")) Then
        MsgBox getTC_FCFromCode("TC0255")  'MsgBox ("Sachbearbeiter fehlt!")  'TC0255
        Exit Sub
    End If

    'Erzeugen einer Excel Instanz
    Set lexcelprov = New clsExcelProvider
    'Prüfen ob XD-Report schon existiert
    Set fs = CreateObject("Scripting.FileSystemObject")
    'ist Verzeichnis vorhanden
    'sPfadDatei = getReklVerzeichnis(Me.ID, Format(Me.Auftragsdatum, "yyyy"), Me.[Meldungsnummer Kunde])
    If (IsNull(Me.Controls("Mastertab#Meldungsnummer Kunde#2")) Or Me.Controls("Mastertab#Meldungsnummer Kunde#2") = "") Then
        sPfadDatei = setReklVerzeichnis(Me.Controls("Mastertab#ID"), format(Me.Controls("Mastertab#Auftragsdatum#3"), "yyyy"), "")      '20190910-ROS
    Else
        sPfadDatei = setReklVerzeichnis(Me.Controls("Mastertab#ID"), format(Me.Controls("Mastertab#Auftragsdatum#3"), "yyyy"), Me.Controls("Mastertab#Meldungsnummer Kunde#2"))
    End If
    If sPfadDatei <> "" Then
        If fs.FileExists(sPfadDatei & "\" & Me.Controls("Mastertab#ID") & "_QS-XD-Report.xlsm") = False Then
            '1. Detailinformationen angeben
'            DoCmd.OpenForm "DetailinformationenReklamationLief", , , , , acDialog            'Öffnen
            '1. Kopie von Template erstellen
            Set wb = lexcelprov.openExcelInstance(XDReportPfad & XDReportTemplateLieferant, sPfadDatei & "\" & Me.Controls("Mastertab#ID") & "_Reklamationsbericht_Lieferant.xlsm")
            '2. XD-Report generieren
            Set WS = wb.Worksheets("Report")
            Set ldb = openAccDBConnection("openDB")
            Set rs = ldb.OpenRecordset("Select * from RSP_Ausw_Basisdaten_1 where [Mastertab].[ID] = " & Me.Controls("Mastertab#ID") & ";", dbOpenSnapshot)
            If rs.RecordCount > 0 Then
                'Aktuelles Datum eintragen
                WS.Range("Datum") = format(Date, "dd.mm.yyyy")
                WS.Range("LQ_Baugruppe") = rs("BG")
                WS.Range("LQ_Artikelnummer") = rs("ArtNr")
                WS.Range("Herstellerteilenummer") = rs("Hersteller-Teilenummer")
                WS.Range("Serialnummer") = rs("Seriennummer")
                WS.Range("Herstellungsdatum") = rs("KW/JJ")
                WS.Range("MeldungsnummerKD") = rs("Meldungsnummer Kunde")
                'Weitere Einträge
                WS.Range("Bestellnummer") = sBestellnummerLQ
                WS.Range("LieferscheinNr") = sLieferscheinLieferant

                'Eintrag der Fehlerbeschreibung und Zellenhöhe am Fehlertext anpassen.
                WS.Range("Fehlerbeschreibung") = rs("Bemerkungen - Maßnahmen")
                'WS.Rows(WS.Range("Fehlerbeschreibung").Row & ":" & WS.Range("Fehlerbeschreibung").Row).Select
                WS.Rows(WS.Range("Fehlerbeschreibung").Row & ":" & WS.Range("Fehlerbeschreibung").Row).AutoFit
                WS.Range("Fehlerbeschreibung").Select

                WS.Range("LQ_SBEmail") = rs("EmailName")
                WS.Range("LQ_SBTelefon") = rs("TelefonBeruflich")
                WS.Range("AnsprechpartnerKunde") = rs("Name")
                WS.Range("Fehlermenge") = rs("Menge betroffen - Kunde") + rs("Menge betroffen - LQ")
                WS.Range("ReklNr") = CStr((rs("ID")))
                If rs("Zuständig") <> "" Then
                    WS.Range("RMA") = rs("Zuständig")
                End If

                '20150825-RSPTag - Aufwand der Reklamation eintragen
                'Extern: Mehraufwand Kunde + LQ
                'Intern: LQ
                Set rsAufwand = ldb.OpenRecordset("SELECT Sum(RSPDouble([Zeit])) AS GesamtZeit From AufwandVorgangPositionen WHERE (((AufwandVorgangPositionen.MID)= " & Me.Controls("Mastertab#ID") & " ));")
                bIntern = False
                If Me.Controls("Mastertab#intern") Then
                    'Wenn bei einem internen Vorgang Zeitpositionen eingetragen sind, werden diese anstelle PE-Bearbeitung hergenommen.
                    'Grund: Es handelt sich hierbei mit sehr hoher Wahrscheinlichkeit um einen Lieferantenfehler, bei dem einen Fehlerreport versendet wurde. In diesen Fall
                    'werden aktuell von den QS-Bearbeiter die Zeiten sträflicherweise im Fat-Client unter den Positionen festgehalten.
                    If rsAufwand.RecordCount And Not IsNull(rsAufwand("GesamtZeit")) Then
                        WS.Range("Aufwand") = (Me.Controls("Mastertab#PE-Bearbeitung") + rsAufwand("GesamtZeit")) * 60
                    Else
                        WS.Range("Aufwand") = (Me.Controls("Mastertab#QS-Bearbeitung") + Me.Controls("Mastertab#PE-Bearbeitung")) * 60
                    End If
                    bIntern = True
                End If
                If Not IsNull(Me.Controls("Mastertab#extern")) And Not bIntern Then
                    If rsAufwand.RecordCount Then
                        WS.Range("Aufwand") = CStr(rsAufwand("GesamtZeit") * 60) + Me.Controls("Mastertab#Aufwand-TRUMPF")

                    Else
                        WS.Range("Aufwand") = Me.Controls("Mastertab#Aufwand-TRUMPF")
                    End If
                End If
                'Pauschale zum Gesamtaufwand addieren. Die Pauschale beinhaltet die Nacharbeit eines Vorgangs nach dem dieser aus der Klärung Lieferant zurück kommt
                'Pauschale beinhaltet: Lagerbewegung Ersatz, Sellungnahme Kunde, Abschlauss Vorgang, usw.
                WS.Range("Aufwand") = WS.Range("Aufwand") + 30
                Set rsAufwand = Nothing

                'Muss neu Programmiert werden wegen unterschiedlichen Rückmeldedatum vom Lieferant
                '########################################################################################
                'Datum der erwarteten Rückantwort vom Lieferant eintragen -> akt. Datum + 1. Mahnstufe
'                Set rs2 = ldb.OpenRecordset("Select * from KonfigurationMahnung;")
'                WS.Range("Bearbeitungsdatum") = getDateAddTage(Date, rs2("Mahnstufe1Tage"))
'                Set rs2 = Nothing
                '########################################################################################

                'Neu unterschiedliches Bearbeitungsdatum vom Lieferant in 3D-Repoprt übernehmen 20190930-ROS
                '########################################################################################
                Set rs2 = ldb.OpenRecordset("Select * from tblReklBearbZeit where [LiefID] = " & Me.Controls("Mastertab#Lieferanten-Code 1") & ";")
                If rs2.RecordCount Then
                    WS.Range("Bearbeitungsdatum") = getDateAddTage(Date, rs2("ATQS"))
                Else
                    WS.Range("Bearbeitungsdatum") = getDateAddTage(Date, 5)
                End If
                Set rs2 = Nothing
                '########################################################################################

                If rs("Annahme mit Nacharbeit") Then
                    WS.Range("AnnahmeMitNacharbeit") = "X"
                End If
                If rs("Annahme mit Prüfausnahme") Then
                    WS.Range("AnnahmeMitPrüfausnahme") = "X"
                End If
                If rs("Rücklieferung an Lieferant") Then
                    WS.Range("WareWirdZurueckgeschickt") = "X"
                End If

    '            WS.Range("Datum") = Date
                'Liefernatenbezogene Daten eintragen
                'WS.Range("FehlerLieferant") = "X"
                Set rsLief = ldb.OpenRecordset("SELECT LieferantAnsprechpartner.AnsprechpartnerID, LieferantAnsprechpartner.Name, LieferantAnsprechpartner.Vorname, " & _
                                                    " LieferantAnsprechpartner.Email, LieferantAnsprechpartner.Telefon, Lieferanten.Lieferantenname, Lieferanten.Postleitzahl, " & _
                                                    " Lieferanten.Ort, Lieferanten.Adresse  " & _
                                                    " FROM LieferantAnsprechpartner INNER JOIN Lieferanten ON LieferantAnsprechpartner.LieferantID = Lieferanten.[Lieferant-ID]  " & _
                                                    " Where LieferantAnsprechpartner.AnsprechpartnerID = " & Me.Controls("Mastertab#LiefAnsprechp-ID#3") & ";")
                If rsLief.RecordCount > 0 Then
                    WS.Range("Lieferant") = rsLief("Lieferantenname")
                    WS.Range("LiefStrasse") = rsLief("Adresse")
                    WS.Range("LiefOrt") = rsLief("Postleitzahl") & " " & rsLief("Ort")
                    WS.Range("LiefAnsprechpartner") = rsLief("Vorname") & " " & rsLief("Name")
                    WS.Range("LiefAnsprechpMail") = rsLief("Email")
                    WS.Range("LiefTel") = rsLief("Telefon")
                    'WS.Range("LiefDatum") = Date
                Else
                End If
                Set rsLief = Nothing
                'Die Maßnahmenzeilen bis auf eine reduzieren
'                For i2 = 29 To (29 - 5) Step -1
'                    WS.Rows(i2 & ":" & i2).Select
'                    Selection.Delete Shift:=xlUp
'                Next i2
                Set WS = Nothing
                Set rs = Nothing
'                 ' Set ldb = openAccDBConnection("closeDB")
            End If
            bReportOeffnen = True
            '...............
        Else
'            If (MsgBox("XD-Report schon vorhanden! Report öffnen?", vbYesNo, "XD-Report") = vbYes) Then    'TC0256
'                bReportOeffnen = True
            If (MsgBox(getTC_FCFromCode("TC0256"), vbYesNo, getTC_FCFromCode("TC0399")) = vbYes) Then
                bReportOeffnen = True
            End If
        End If

    Else
        MsgBox getTC_FCFromCode("TC0257")  'MsgBox ("Reklamationsvorgangspfad nicht vorhanden!")  'TC0257
    End If
    'wb.Activate
    Set fs = Nothing

    lexcelprov.closeExcelInstance
    Set lexcelprov = Nothing
    If bReportOeffnen Then
        ShellExecute 0, "open", sPfadDatei & "\" & Me.Controls("Mastertab#ID") & "_Reklamationsbericht_Lieferant.xlsm", "", "", 1
    End If

end_XDReportLief_DblClick:
    Exit Sub

err_XDReportLief_DblClick:

    MsgBox Error
End Sub


Private Sub Mastertab_Artikel_ID_1_4_NotInList(NewData As String, Response As Integer)

'20160510-RSP

'    If Not bArtikelNotInList Then
'        If Not IsNull(Me.Controls("Mastertab#Artikel-ID 1#5").Column(1)) Then
'            sLQArtikelTemp = Me.Controls("Mastertab#Artikel-ID 1#5").Column(1)
'            sCurrArtikelSource = "Artikel"
'            DoCmd.OpenForm "Artikel-avista_outTemp", , , , , acDialog
'            Me.Controls("Mastertab#Artikel-ID 1#5") = lLQArtikelTemp
''            getLQArtNrAvista (Me.[Artikel-ID_1].Column(1))
'        End If
''
''        If Not IsNull(Me.[Artikel-ID_1].Column(1)) Then
''            getLQArtNrAvista (Me.[Artikel-ID_1].Column(1))
''        End If
'    End If
'    Me.Controls("Mastertab#Artikel-ID 1#5").Requery
'    Me.Controls("Mastertab#Artikel-ID 1#4").Requery
'    Me.Controls("Mastertab#Artikel-ID 1#6").Requery
'    Me.Controls("Mastertab#Artikel-ID 1#3").Requery
'
'    bArtikelNotInList = False


    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rs As ADODB.Recordset
    Dim ldb As DAO.Database
    Dim rs2 As DAO.Recordset
    Dim lArt_ID As Long

    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider
    'Fehlerausgabe unterdrücken
    Response = acDataErrAdded
    'Recordset-Instanz
    Set rs = New ADODB.Recordset
    Set rs = adbprov.getRecordsetAvista(AvistaKndArtNrSQL_Select & NewData & "' " & AvistaArtikelSQL_Order)
    If rs.RecordCount Then
        If DCount("[Kd Mat-Nr]", "Artikel-all_avista_out", "[Kd Mat-Nr] = '" & NewData & "'") = 0 Then
            Set ldb = openAccDBConnection("openDB")
            Set rs2 = ldb.OpenRecordset("Artikel-all_avista_out", dbOpenDynaset)
            rs2.AddNew
                rs2![LQ Art-Nr] = rs("ast_nummer")
                rs2![Variante] = rs("avar_nummer")
                rs2![Kd Mat-Nr] = rs("avkd_artnr_kd")
                rs2![Bezeichnung 1] = rs("ast_bezeichnung")
                rs2![Hersteller-Teilenummer] = rs("avar_herst_nummer")
                lArt_ID = rs2![Artikel-ID]
            rs2.Update
            rs2.Close
            Set rs2 = Nothing
            Response = acDataErrContinue
            Me.Controls("Mastertab#Artikel-ID 1#4") = lArt_ID
            Me.Requery
            Me.Refresh
'            ldb.Close
'            Set db = Nothing
        End If
    End If
    rs.Close
    Set rs = Nothing
    Set adbprov = Nothing


ende:
    Exit Sub
Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0259
    Resume ende
End Sub




Private Sub Mastertab_Artikel_ID_1_5_NotInList(NewData As String, Response As Integer)

'20160510-RSP
    sLQArtikelTemp = NewData
    'BG Zurodnung
    sCurrArtikelSource = "Artikel"
    DoCmd.OpenForm "Artikel-avista_outTemp", , , , , acDialog
    Me.Controls("Mastertab#Artikel-ID 1#5") = lLQArtikelTemp
    Response = acDataErrAdded
    bArtikelNotInList = True

ende:
    Exit Sub

Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0259
    Resume ende


End Sub

Private Sub Mastertab_Baugruppen_ID_3_NotInList(NewData As String, Response As Integer)

'20150119-RSP:

    sLQArtikelTemp = NewData
    'BG Zurodnung
    sCurrArtikelSource = "BG"
    DoCmd.OpenForm "Artikel-avista_outTemp", , , , , acDialog
    Me.Controls("Mastertab#Baugruppen-ID#3") = lLQArtikelTemp
    Response = acDataErrAdded
    bArtikelNotInList = True
    Me.Controls("Mastertab#Baugruppen-ID#3").Requery
    Me.Requery


ende:
    Exit Sub

Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0259
    Resume ende


End Sub

Private Sub Mastertab_Baugruppen_ID_2_NotInList(NewData As String, Response As Integer)

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rs As ADODB.Recordset
    Dim ldb As DAO.Database
    Dim rs2 As DAO.Recordset
    Dim lBG_ID As Long

    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider
    'Fehlerausgabe unterdrücken
    Response = acDataErrAdded
    'Recordset-Instanz
    Set rs = New ADODB.Recordset
    Set rs = adbprov.getRecordsetAvista(AvistaKndArtNrSQL_Select & NewData & "' " & AvistaArtikelSQL_Order)
    If rs.RecordCount Then
        If DCount("[Kunden-Nr]", "Baugruppen-avista_out", "[Kunden-Nr] = '" & NewData & "'") = 0 Then
            Set ldb = openAccDBConnection("openDB")
            Set rs2 = ldb.OpenRecordset("Baugruppen-avista_out", dbOpenDynaset)
            rs2.AddNew
                rs2![LQ Art-Nr] = rs("ast_nummer")
                rs2![Var Nr] = rs("avar_nummer")
                rs2![Kunden-Nr] = rs("avkd_artnr_kd")
                rs2![Bezeichnung 1] = rs("ast_bezeichnung")
                rs2![Produktgruppe] = rs("pg_bez")
                lBG_ID = rs2![Baugruppen-ID]
            rs2.Update
            rs2.Close
            Set rs2 = Nothing
            Response = acDataErrContinue
            Me.Controls("Mastertab#Baugruppen-ID#2") = lBG_ID
            Me.Requery
            Me.Refresh
'            ldb.Close
'            Set db = Nothing
        End If
    End If
    rs.Close
    Set rs = Nothing
    Set adbprov = Nothing

ende:
    Exit Sub

Fehler:
    MsgBox Err.Description, 16, getTC_FCFromCode("TC0228")  'TC0259
    Resume ende
End Sub


Public Sub ReparaturauftragExternErstellen()    '20180921-ROS
On Error GoTo err_ReparaturauftragExternErstellen

    Dim ldb As DAO.Database
    Dim rs, rsLief, rs2 As DAO.Recordset
    Dim fs As Object
    Dim wb As Excel.Workbook
    Dim WS As Excel.Worksheet
    Dim objExcel As Object
    Dim lexcelprov As clsExcelProvider
    Dim bReportOeffnen As Boolean
    Dim sPfadDatei As String
    Dim bIntern As Boolean '20150825 - RSPTag
    Dim rsAufwand As DAO.Recordset '20150825 - RSPTag
    Dim GesamtZeit As String '20160224-ROS
    Dim Anlieferwerk As String
    Dim Ort As String
    Dim PLZ As String
    Dim Werk As String

    Dim XDReportPfad As String  '20160304-ROS
    XDReportPfad = getTC_FCFromCode("DP0001")
'    XDReportPfad = "C:\Users\Roberto Santarelli\Documents\Entwicklung_ROS\Ordner Testdatenbank\Vorlagen\"

    Dim RepAuftragTemplate As String
    RepAuftragTemplate = getTC_FCFromCode("DP0009")
'    RepAuftragTemplate = "20170215-ROS_Reparaturauftrag_Extern.xlsm"

    Anlieferwerk = DLookup("Strasse", "Sachbearbeiter", "[Name] = '" & Me.Controls("Mastertab#Sachbearbeiter#1") & "'")
    Ort = DLookup("Ort", "Sachbearbeiter", "[Name] = '" & Me.Controls("Mastertab#Sachbearbeiter#1") & "'")
    PLZ = DLookup("PLZ", "Sachbearbeiter", "[Name] = '" & Me.Controls("Mastertab#Sachbearbeiter#1") & "'")
    Werk = DLookup("Bürostandort", "Sachbearbeiter", "[Name] = '" & Me.Controls("Mastertab#Sachbearbeiter#1") & "'")


    'Init
    bReportOeffnen = False

    'Plausibilitätsprüfung
    If IsNull(Me.Controls("Mastertab#Sachbearbeiter#1")) Then
        MsgBox getTC_FCFromCode("TC0255")  'MsgBox ("Sachbearbeiter fehlt!")  'TC0255
        Exit Sub
    End If
    'Erzeugen einer Excel Instanz
    Set lexcelprov = New clsExcelProvider
    'Prüfen ob XD-Report schon existiert
    Set fs = CreateObject("Scripting.FileSystemObject")
    'ist Verzeichnis vorhanden
    'sPfadDatei = getReklVerzeichnis(Me.ID, Format(Me.Auftragsdatum, "yyyy"), Me.[Meldungsnummer Kunde])
    If (IsNull(Me.Controls("Mastertab#Meldungsnummer Kunde#2")) Or Me.Controls("Mastertab#Meldungsnummer Kunde#2") = "") Then
        sPfadDatei = setReklVerzeichnis(Me.Controls("Mastertab#ID"), format(Me.Controls("Mastertab#Auftragsdatum#3"), "yyyy"), "")
    Else
        sPfadDatei = setReklVerzeichnis(Me.Controls("Mastertab#ID"), format(Me.Controls("Mastertab#Auftragsdatum#3"), "yyyy"), Me.Controls("Mastertab#Meldungsnummer Kunde#2"))
    End If
    If sPfadDatei <> "" Then
        If fs.FileExists(sPfadDatei & "\" & Me.Controls("Mastertab#ID") & "_Reparaturauftrag.xlsm") = False Then
            '1. Detailinformationen angeben
'            DoCmd.OpenForm "DetailinformationenReklamationLief", , , , , acDialog            'Öffnen
            '1. Kopie von Template erstellen
            Set wb = lexcelprov.openExcelInstance(XDReportPfad & RepAuftragTemplate, sPfadDatei & "\" & Me.Controls("Mastertab#ID") & "_Reparaturauftrag_Extern.xlsm")
            '2. XD-Report generieren
            Set WS = wb.Worksheets("Rep_Auftrag")
            Set ldb = openAccDBConnection("openDB")
            Set rs = ldb.OpenRecordset("Select * from RSP_Ausw_Basisdaten_1 where [Mastertab].[ID] = " & Me.Controls("Mastertab#ID") & ";")
            If rs.RecordCount > 0 Then
                'Aktuelles Datum eintragen
                WS.Range("Datum") = format(Date, "dd.mm.yyyy")
                WS.Range("Anlieferwerk") = Anlieferwerk & " " & "(" & Werk & ")"
                WS.Range("Ort") = PLZ & " " & Ort
                WS.Range("Meldungsnummer_Kunde") = rs("Meldungsnummer Kunde")
                WS.Range("LQ_Artikelnummer") = rs("ArtNr")
                WS.Range("Herstellerteilenummer") = rs("Hersteller-Teilenummer")
                WS.Range("Serialnummer") = rs("Seriennummer")
                WS.Range("Herstelldatum") = rs("KW/JJ")
                WS.Range("Anliefertermin") = sReparaturdatum
                WS.Range("Bestellnummer") = sBestellnummerLQ
'                WS.Range("LieferscheinNr") = sLieferscheinLieferant
                WS.Range("LQ_SBEmail") = rs("EmailName")
                WS.Range("LQ_SBTelefon") = rs("TelefonBeruflich")
                WS.Range("AnsprechpartnerKunde") = rs("Name")
                WS.Range("Fehlermenge") = rs("Menge betroffen - LQ")
                WS.Range("ReklNr") = CStr((rs("ID")))
                WS.Range("RepNr") = CStr((rs("ID")))
                If rs("Zuständig") <> "" Then
                    WS.Range("RMA") = rs("Zuständig")
                End If
                'Eintrag der Fehlerbeschreibung und Zellenhöhe am Fehlertext anpassen.
                WS.Range("Fehlerbeschreibung") = rs("Bemerkungen - Maßnahmen")
                'WS.Rows(WS.Range("Fehlerbeschreibung").Row & ":" & WS.Range("Fehlerbeschreibung").Row).Select
                WS.Rows(WS.Range("Fehlerbeschreibung").Row & ":" & WS.Range("Fehlerbeschreibung").Row).AutoFit
                WS.Range("Fehlerbeschreibung").Select
                Set rsLief = ldb.OpenRecordset("SELECT LieferantAnsprechpartner.AnsprechpartnerID, LieferantAnsprechpartner.Name, LieferantAnsprechpartner.Vorname, " & _
                                                    " LieferantAnsprechpartner.Email, LieferantAnsprechpartner.Telefon, Lieferanten.Lieferantenname, Lieferanten.Postleitzahl, " & _
                                                    " Lieferanten.Ort, Lieferanten.Adresse  " & _
                                                    " FROM LieferantAnsprechpartner INNER JOIN Lieferanten ON LieferantAnsprechpartner.LieferantID = Lieferanten.[Lieferant-ID]  " & _
                                                    " Where LieferantAnsprechpartner.AnsprechpartnerID = " & Me.Controls("Mastertab#LiefAnsprechp-ID#3") & ";")
                If rsLief.RecordCount > 0 Then
                    WS.Range("Lieferant") = rsLief("Lieferantenname")
                    WS.Range("LiefStrasse") = rsLief("Adresse")
                    WS.Range("LiefOrt") = rsLief("Postleitzahl") & " " & rsLief("Ort")
                    WS.Range("LiefAnsprechpartner") = rsLief("Vorname") & " " & rsLief("Name")
                    WS.Range("LiefAnsprechpMail") = rsLief("Email")
                    WS.Range("LiefTel") = rsLief("Telefon")
                Else
                End If
                Set rsLief = Nothing
                Set WS = Nothing
                Set rs = Nothing
            End If
            bReportOeffnen = True
        Else
'            If (MsgBox("XD-Report schon vorhanden! Report öffnen?", vbYesNo, "XD-Report") = vbYes) Then    'TC0256
'                bReportOeffnen = True
            If (MsgBox(getTC_FCFromCode("TC0256"), vbYesNo, getTC_FCFromCode("TC0399")) = vbYes) Then
                bReportOeffnen = True
            End If
        End If

    Else
        MsgBox getTC_FCFromCode("TC0257")  'MsgBox ("Reklamationsvorgangspfad nicht vorhanden!")  'TC0257
    End If
    Set fs = Nothing

    lexcelprov.closeExcelInstance
    Set lexcelprov = Nothing
    If bReportOeffnen Then
        ShellExecute 0, "open", sPfadDatei & "\" & Me.Controls("Mastertab#ID") & "_Reparaturauftrag_Extern.xlsm", "", "", 1
    End If

    Exit Sub

err_ReparaturauftragExternErstellen:

    MsgBox Error
End Sub
Sub privFormClose()


On Error GoTo err_Form_Close

Dim sMemo As String
Dim ldb As DAO.Database
Dim rs, rsReklVorgang As DAO.Recordset
Dim Reklnr, sSQL As String

Dim cPERechercheHistorie As String  '20160304-ROS
cPERechercheHistorie = getTC_FCFromCode("DP0005")

 Set ldb = openAccDBConnection("openDB")
'Ermittlung ob es zu dem Vorgang schon ein Protokolleintrag vorhanden ist
'ReklNr = Forms("Fehlererfassung - Register").Controls("ID")
'If p_KommentarKopieren Then
    Set rs = ldb.OpenRecordset("Select * from ReklMemo where [ID] = " & p_ReklNr & ";", dbOpenSnapshot)

    If rs.RecordCount Then
        If Len(rs("Aktueller Eintrag").Value) > 27 Then
            If IsNull(rs("Protokoll").Value) Then
                sMemo = ""
            Else
                sMemo = rs("Protokoll").Value
            End If
            sMemo = rs("Aktueller Eintrag").Value & Chr(13) & Chr(10) & Chr(13) & Chr(10) & sMemo
            sMemo = Replace(sMemo, "'", "")
            sMemo = Replace(sMemo, """", "")
            ldb.Execute ("Update ReklMemo set [Protokoll] = '" & sMemo & _
                        "' where ID = " & p_ReklNr & ";")
            sMemo = ""
            ldb.Execute ("Update ReklMemo set [Aktueller Eintrag] = NULL where ID = " & p_ReklNr & ";")
            'db.Execute ("Update ReklMemo set [Aktueller Eintrag] = '" & sMemo & _
            '            "' where ID = " & ReklNr & ";")
        End If
    End If
'End If
'p_KommentarKopieren = False
'Datensatz zu nächste Bearbeitung freigeben
unlockDataset (p_ReklNr)
'db.Execute ("Update [Mastertab-DSBerechtigung] set [Zustand] = 'Frei' where ID = " & reklNr & ";")

end_Form_Close:
   Set rs = Nothing
'   Set db = Nothing
   Exit Sub

err_Form_Close:
    If Err = 3188 Then
        MsgBox getTC_FCFromCode("TC0431"), vbOKOnly
        If getFileExists(cPERechercheHistorie & Me.Controls("Mastertab#ID") & "_Fehlerprotokoll.txt") Then
            sMemo = readTextFile(cPERechercheHistorie & Me.Controls("Mastertab#ID") & "_Fehlerprotokoll.txt")
            sMemo = rs("Aktueller Eintrag").Value & Chr(13) & Chr(10) & Chr(13) & Chr(10) & sMemo
            sMemo = Replace(sMemo, "'", "")
            sMemo = Replace(sMemo, """", "")
        End If
        sMemo = writeTextFile(cPERechercheHistorie & Me.Controls("Mastertab#ID") & "_Fehlerprotokoll.txt", sMemo, "Overwrite")


        'Funktion zum Erstellen und Eintrag des externen Protokollfiles
        '
        Resume Next
    Else
        MsgBox getTC_FCFromCode("TC0432"), vbCritical, "Error"
        Resume end_Form_Close

    End If

End Sub

'########################################################################
'#######################Public Function and Subs#########################
'########################################################################
Public Sub closeInstanceForm(fForm As Form)

    l_Form.Visible = False
    Set l_Form = Nothing

End Sub

Public Sub ReLoadForm()

    Call Form_Load

End Sub
Public Sub frmFRControlEventAfterUpdate(ByVal fCtrlName As String)

    Select Case fCtrlName
        Case "Mastertab#LQEinheitID#1"
            Call Mastertab_LQEinheitID_1_AfterUpdate
        Case "Mastertab#LQBereichID#1"
            Call Mastertab_LQBereichID_1_AfterUpdate
        Case "Mastertab#LQWorkstationID#1"
            Call Mastertab_LQWorkstationID_1_AfterUpdate
        Case "Mastertab#TaetigkeitID#1"
            Call Mastertab_TaetigkeitID_1_AfterUpdate
        Case "Mastertab#DefTaetigkeitID#1"
            Call Mastertab_DefTaetigkeitID_1_AfterUpdate
        Case "ROS_Daten_Einkauf#TaetigkeitLieferantID"
            Call ROS_Daten_Einkauf_TaetigkeitLieferantID_AfterUpdate
        Case "Mastertab#DefTaetigkeitAnalyseID"
            Call Mastertab_DefTaetigkeitAnalyseID_AfterUpdate


        Case "Mastertab#LQEinheitErfasserID"
            Call Mastertab_LQEinheitErfasserID_AfterUpdate
        Case "Mastertab#LQBereichErfasserID"
            Call Mastertab_LQBereichErfasserID_AfterUpdate
        Case "Mastertab#LQWorkstationErfasserID"
            Call Mastertab_LQWorkstationErfasserID_AfterUpdate
        Case "Mastertab#TaetigkeitErfasserID"
            Call Mastertab_TaetigkeitErfasserID_AfterUpdate
        Case "Mastertab#DefTaetigkeitErfasserID"
            Call Mastertab_DefTaetigkeitErfasserID_AfterUpdate

        Case "Mastertab#LQEinheitID"
            Call Mastertab_LQEinheitID_AfterUpdate
        Case "Mastertab#LQBereichID"
            Call Mastertab_LQBereichID_AfterUpdate
        Case "Mastertab#LQWorkstationID"
            Call Mastertab_LQWorkstationID_AfterUpdate
        Case "Mastertab#TaetigkeitID"
            Call Mastertab_TaetigkeitID_AfterUpdate
        Case "Mastertab#DefTaetigkeitID"
            Call Mastertab_DefTaetigkeitID_AfterUpdate

        Case "RegisterStr246_Change"
            Call RegisterStr246_Change

    End Select

End Sub



'########################################################################
'#### Bereich Melder Daten 1 ############################################
'########################################################################

Public Function setNewAnsprechpartner(fID As Long)

    Dim db

    Me.Controls("Mastertab#Kunden-ID#1") = fID
    Me.Requery
    Me.frmUFKundenDaten.Requery

End Function

Private Sub ReklasZuArtikel(Optional ByVal ArtikelNr As String, Optional ByVal FaNr As String)      '20191002_2-ROS


'On Error GoTo err_ReklasZuArtikel

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim sÜbersicht As String
    Dim Dat1 As Date
    Dim Dat2 As Date
    Dim Dat3 As String
    Dim Dat4 As String
    Dim bFilterOn As Boolean
    Dim sfilter As String


    Dat1 = DateAdd("yyyy", -1, Date)
    Dat2 = Date
    Dat3 = format(Dat1, "mm\/dd\/yyyy")
'    Dat4 = Format(Dat2, "dd\/mm\/yyyy")
    Dat4 = format(Dat2, "mm\/dd\/yyyy")


    If Me.Controls("Mastertab#extern") = True Then sÜbersicht = "Reklamationen_Übersicht"
    If Me.Controls("Mastertab#intern") = True Then sÜbersicht = "Reklamationen_ÜbersichtIntern"

    bFilterOn = False

'20191118-ROS
    If ArtikelNr <> "" Then
        sfilter = "(([Kategorie-Problem] ='berechtigt' or [Kategorie-Problem] ='nicht berechtigt') And ([Auftragsdatum] between #" & Dat3 & "# and #" & Dat4 & "#) And ( [Artikel-all_avista_out].[LQ Art-Nr] like '" & ArtikelNr & "*'))"
    End If

    If FaNr <> "" Then
'        sfilter = "(([Kategorie-Problem] ='berechtigt' or [Kategorie-Problem] ='nicht berechtigt') And ([Auftragsdatum] between #" & Dat3 & "# and #" & Dat4 & "#) And ( [Fertigungsauftrag-FA] like '" & FaNr & "*'))"
        '20260407-ROS weitere Reklamationen zu FA geändert ohne Beschränkung Auftragsdatum
        sfilter = "(([Kategorie-Problem] ='berechtigt' or [Kategorie-Problem] ='nicht berechtigt') And ( [Fertigungsauftrag-FA] like '" & FaNr & "*'))"
    End If

    Select Case sÜbersicht
        Case "Reklamationen_Übersicht"
            If Not IsLoaded(sÜbersicht) Then
                DoCmd.OpenForm sÜbersicht
            End If
            If IsLoaded("Reklamationen_ÜbersichtIntern") Then
                DoCmd.Close acForm, "Reklamationen_ÜbersichtIntern"
            End If

        Case "Reklamationen_ÜbersichtIntern"
            If Not IsLoaded(sÜbersicht) Then
                DoCmd.OpenForm sÜbersicht
            End If
            If IsLoaded("Reklamationen_Übersicht") Then
                DoCmd.Close acForm, "Reklamationen_Übersicht"
            End If
    End Select

    If Not IsLoaded(sÜbersicht) Then
        DoCmd.OpenForm sÜbersicht
    End If

    Forms(sÜbersicht).FilterOn = False
    'Filter in Übersicht setzen und aktualisieren
    Forms(sÜbersicht).Controls("FilterView") = sfilter
    Forms(sÜbersicht).Filter = sfilter
    Forms(sÜbersicht).FilterOn = True
    If Forms(sÜbersicht).RecordsetClone.RecordCount = 0 Then
        '20260407-ROS Ausgabe-Text angepasst
        If FaNr <> "" Then
            MsgBox getTC_FCFromCode("TC4283"), vbInformation, getTC_FCFromCode("TC4284")
        Else
            MsgBox getTC_FCFromCode("TC4223"), vbInformation, getTC_FCFromCode("TC4224")
        End If
        Forms("Fehlererfassung - Register").SetFocus
'        DoCmd.Close acForm, sÜbersicht
    End If

'    Forms(sÜbersicht).Requery
'    Forms(sÜbersicht).Refresh

    'Query zum Exportieren Filter setzen (Where)
    'Dim ldb As DAO.Database
'    Dim QS As DAO.QueryDef
'    Set ldb = openAccDBConnection("openDB")
'    Set QS = ldb.QueryDefs("RSP_Ausw_Basisdaten_1_Export")
'    If Not bFilterOn Then
'        If InStr(QS.sql, "Where") Then
'            QS.sql = Left(QS.sql, InStr(QS.sql, "Where") - 1) & sfilter & ";"
'        End If
'    Else
'        If InStr(QS.sql, "Where") Then
'            QS.sql = Left(QS.sql, InStr(QS.sql, "Where") + 5) & sfilter & ";"
'        Else
'            QS.sql = Left(QS.sql, InStr(QS.sql, ";") - 1) & " Where " & sfilter & ";"
'        End If
'    End If

'exit_ReklasZuArtikel:
'    Set rs = Nothing
'
''    DoCmd.OpenForm sÜbersicht
'    Forms(sÜbersicht).SetFocus
'    Exit Sub
'
'err_ReklasZuArtikel:
'    If Err = 3021 Then
'        Resume Next
'    Else
'        GoTo exit_ReklasZuArtikel
'    End If
End Sub
Private Sub cmbPosKosten_Click()    '20220125-ROS

    Call PosKostenAnlegen("")

End Sub
Public Sub PosKostenAnlegen(ByVal fMatKosten As String)

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim rs1 As DAO.Recordset
    Dim sLastPos As String
    Dim lKuerzel As String
    Dim lKosten As Double
    Dim IKostenTyp As Integer

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from tblKostenaufstellung where ReklID=" & Me.Mastertab_ID & " ORDER BY KostenaufstellungID;")

    If fMatKosten = "Frachtkosten" Then
        'Pauschal-Frachtkosten können Individuell für das Land angepasst werden für DE sind 10€ eingestellt
        lKosten = getTC_FCFromCode("TC0502")
        IKostenTyp = 9
    End If

    If fMatKosten = "MatKosten" Then
        lKosten = Forms("Fehlererfassung - Register").Controls("Mastertab#Material-Kosten")
        IKostenTyp = 5
    End If

    If fMatKosten = "RepKosten" Then
        lKosten = Forms("Fehlererfassung - Register").Controls("Mastertab#Reparaturbetrag")
'        If lKosten = 0 Then Exit Sub
        IKostenTyp = 11
    End If

    If fMatKosten = "VerrechnungKD" Then
        lKosten = Forms("Fehlererfassung - Register").Controls("Mastertab#Belastungsbetrag LQ")
'        If lKosten = 0 Then Exit Sub
        IKostenTyp = 10
    End If

    Set rs1 = ldb.OpenRecordset("Select * from tblKostenaufstellung where ReklID=" & Me.Mastertab_ID & " and KostentypID= " & IKostenTyp & "  ORDER BY KostenaufstellungID;")

    If rs.RecordCount = 0 Then
        'Anlegen Erfassungsposition 10
        rs.AddNew
        rs("ReklID") = Me.Mastertab_ID
        rs("Pos") = "10"
        rs("Datum") = Date
        rs("KZ") = get_pKuerzel()
        rs("KostentypID") = IKostenTyp
        rs("Kosten") = lKosten
        rs.Update
    Else
        If IsNull(rs("Datum")) Then
            MsgBox getTC_FCFromCode("TC0333", ldb)  'MsgBox "Kein neuer Datensatz möglich, da in der Letzten Position Datum fehlt!"   'TC0333
        Else
            If rs1.RecordCount Then
                If lKosten = 0 Then
                    rs1.Delete
                Else
                    rs1.Edit
                    rs1("Kosten") = lKosten
                    rs1.Update
                End If
            Else
                rs.MoveLast
                sLastPos = rs("Pos")
                rs.AddNew
                rs("ReklID") = Me.Mastertab_ID
                rs("Datum") = Date
                rs("Pos") = CStr(CInt(sLastPos) + 10) 'Pos um 10 erhöhen
                rs("KZ") = get_pKuerzel()
                rs("KostentypID") = IKostenTyp
                rs("Kosten") = lKosten
                rs.Update
            End If
        End If
    End If

    Me.Controls("UF_Kostenaufstellung").Requery
    DoCmd.GoToRecord , , acLast

    Set rs1 = Nothing
    Set rs = Nothing
End Sub
