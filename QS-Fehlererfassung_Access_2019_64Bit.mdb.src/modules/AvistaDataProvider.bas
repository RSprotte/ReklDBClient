Attribute VB_Name = "AvistaDataProvider"
Option Compare Database
Option Explicit

'20160302-ROS
Public Const AvistaArtikelSQL_Select = "SELECT tm_artikel.ast_nummer, tm_artikel_var.avar_nummer, tm_artikel_var.avar_herst_nummer, tm_artikelvar_kd.avkd_artnr_kd, tm_artikel.ast_bezeichnung, va_produktgruppe.pg_bez " & _
                                "FROM tm_artikel " & _
                                "INNER JOIN tm_artikel_var ON tm_artikel.ast_id = tm_artikel_var.ast_id " & _
                                "INNER JOIN tx_artikelstatus ON tm_artikel.artstat_id = tx_artikelstatus.artstat_id " & _
                                "LEFT OUTER JOIN tm_artikelvar_kd ON tm_artikelvar_kd.ast_id = tm_artikel.ast_id AND tm_artikelvar_kd.avar_id = tm_artikel_var.avar_id " & _
                                "LEFT OUTER JOIN va_produktgruppe ON tm_artikel.pg_id = va_produktgruppe.pg_id AND va_produktgruppe.pg_gueltig = 1 AND va_produktgruppe.ma_id = 2 AND va_produktgruppe.sprach_id = 1 " & _
                                "WHERE tm_artikel.ast_nummer <> '-' AND tx_artikelstatus.artstat_archiv = 0 AND tm_artikel.ast_nummer like '"
Public Const AvistaKndArtNrSQL_Select = "SELECT tm_artikel.ast_nummer, tm_artikel_var.avar_nummer, tm_artikel_var.avar_herst_nummer, tm_artikelvar_kd.avkd_artnr_kd, tm_artikel.ast_bezeichnung, va_produktgruppe.pg_bez " & _
                                "FROM tm_artikel " & _
                                "INNER JOIN tm_artikel_var ON tm_artikel.ast_id = tm_artikel_var.ast_id " & _
                                "INNER JOIN tx_artikelstatus ON tm_artikel.artstat_id = tx_artikelstatus.artstat_id " & _
                                "LEFT OUTER JOIN tm_artikelvar_kd ON tm_artikelvar_kd.ast_id = tm_artikel.ast_id AND tm_artikelvar_kd.avar_id = tm_artikel_var.avar_id " & _
                                "LEFT OUTER JOIN va_produktgruppe ON tm_artikel.pg_id = va_produktgruppe.pg_id AND va_produktgruppe.pg_gueltig = 1 AND va_produktgruppe.ma_id = 2 AND va_produktgruppe.sprach_id = 1 " & _
                                "WHERE tm_artikel.ast_nummer <> '-' AND tx_artikelstatus.artstat_archiv = 0 AND tm_artikelvar_kd.avkd_artnr_kd = '"



'Public Const AvistaArtikelSQL_Order = "GROUP BY tm_artikel.ast_nummer, tm_artikel_var.avar_nummer, tm_artikel_var.avar_herst_nummer, tm_artikelvar_kd.avkd_artnr_kd, tm_artikel.ast_bezeichnung, va_produktgruppe.pg_bez ORDER BY tm_artikel.ast_nummer"
Public Const AvistaArtikelSQL_Order = "GROUP BY tm_artikel.ast_nummer, tm_artikel_var.avar_nummer, tm_artikel_var.avar_herst_nummer, tm_artikelvar_kd.avkd_artnr_kd, tm_artikel.ast_bezeichnung, va_produktgruppe.pg_bez ORDER BY tm_artikel_var.avar_nummer"
'Ansprechpartner Lieferant
Public Const SelectAnsprechp = "SELECT DISTINCT tm_partei_firma.pa_id AS pa_id_firma, tm_partei.pa_org_name1, tm_partei.pa_pers_vorname, tm_partei.pa_titel, tm_partei.pa_anrede_brief, tm_partei.pa_anrede_adresse, tm_partei.anrede_id, tm_adress_basis.*, tm_ansprechpart.*, pa_rel_asp.pa_id_to, tm_partei_firma.pa_org_name1 AS pa_org_name_firma, tm_kunde.adk_nummer, tm_kunde.adk_id, tm_lieferant.adl_nummer, tm_lieferant.adl_id, tm_partei_erg.* " & _
                        "FROM (((((((DBA.tm_partei INNER JOIN DBA.tm_ansprechpart ON tm_partei.pa_id = tm_ansprechpart.asp_id INNER JOIN DBA.tm_partei_erg ON tm_partei.pa_id = tm_partei_erg.pa_id) INNER JOIN DBA.tm_adress_basis ON tm_partei.pa_id = tm_adress_basis.pa_id AND adt_id = 1) " & _
                        "INNER JOIN DBA.tx_partei_relation AS pa_rel_asp ON tm_partei.pa_id = pa_rel_asp.pa_id_from ) INNER JOIN DBA.tm_partei AS tm_partei_firma ON tm_partei_firma.pa_id = pa_rel_asp.pa_id_to ) " & _
                        "LEFT OUTER JOIN DBA.tx_partei_relation AS pa_rel_mandant_left ON tm_partei_firma.pa_id = pa_rel_mandant_left.pa_id_from ) LEFT OUTER JOIN DBA.tx_partei_relation AS pa_rel_mandant_right ON tm_partei_firma.pa_id = pa_rel_mandant_right.pa_id_to ) " & _
                        "LEFT OUTER JOIN DBA.tm_kunde ON tm_partei_firma.pa_id = tm_kunde.adk_id ) LEFT OUTER JOIN DBA.tm_lieferant ON tm_partei_firma.pa_id = tm_lieferant.adl_id "
Public Const WhereStartLiefNr = "WHERE ( pa_rel_asp.reltyp_id = 4 OR pa_rel_asp.reltyp_id = 7 OR pa_rel_asp.reltyp_id = 25) AND (pa_rel_mandant_left.pa_id_to = 2 OR pa_rel_mandant_right.pa_id_from = 2) AND tm_lieferant.adl_nummer = "
'Public Const WherePara1_Pos = "AND tm_partei_erg.pae_info12 LIKE "  'tm_ansprechpart.asp_info12
Public Const WherePara1_Pos = "AND tm_ansprechpart.asp_info12 LIKE "  'tm_ansprechpart.asp_info12
Public Const WhereOrderAnsprp = "AND tm_adress_basis.adt_id = 1 ORDER BY tm_partei.pa_org_name1"

'Fertigungsaufträge
Public Const SelectFA = "SELECT aint_nr, ast_nummer, avar_nummer, ast_bezeichnung, avar_bezeichnung, aint_bedarf_menge, ast_matchcode, aint_abr_lager_id, tm_lager.lager_bez, aint_text3 " & _
                        "FROM (DBA.tt_auftrag_intern " & _
                        "LEFT OUTER JOIN DBA.tm_lager ON tt_auftrag_intern.aint_abr_lager_id = tm_lager.lager_id) " & _
                        "LEFT OUTER JOIN DBA.tm_artikel ON tt_auftrag_intern.ast_id = tm_artikel.ast_id " & _
                        "LEFT OUTER JOIN DBA.tm_artikel_var ON tt_auftrag_intern.ast_id = tm_artikel_var.ast_id AND tt_auftrag_intern.avar_id = tm_artikel_var.avar_id " & _
                        "WHERE tt_auftrag_intern.aint_typ = 1 AND tt_auftrag_intern.vgst_id = 500 AND tt_auftrag_intern.aint_nr LIKE "

'20180125-ROS 'Fertigungsaufträge
Public Const SelectFA2 = "SELECT aint_nr, ast_nummer, avar_nummer, ast_bezeichnung, avar_bezeichnung, aint_bedarf_menge, ast_matchcode, aint_abr_lager_id, tm_lager.lager_bez, aint_text3, um_produktionsauftrag.DATUM7, um_produktionsauftrag.NAME4 " & _
                        "FROM (DBA.tt_auftrag_intern " & _
                        "LEFT OUTER JOIN DBA.um_produktionsauftrag ON tt_auftrag_intern.aint_id = um_produktionsauftrag.aint_id " & _
                        "LEFT OUTER JOIN DBA.tm_lager ON tt_auftrag_intern.aint_abr_lager_id = tm_lager.lager_id) " & _
                        "LEFT OUTER JOIN DBA.tm_artikel ON tt_auftrag_intern.ast_id = tm_artikel.ast_id " & _
                        "LEFT OUTER JOIN DBA.tm_artikel_var ON tt_auftrag_intern.ast_id = tm_artikel_var.ast_id AND tt_auftrag_intern.avar_id = tm_artikel_var.avar_id " & _
                        "WHERE tt_auftrag_intern.aint_typ = 1 AND tt_auftrag_intern.vgst_id = 500 AND tt_auftrag_intern.aint_nr LIKE "

'Mitabrbeiter LQ
Public Const SelectMAKuerzel = "SELECT tm_partei.pa_id, tm_partei.pa_org_name1, tm_partei.pa_pers_vorname, tm_partei.pa_org_name2, cmgroup_name, tm_adress_basis.adb_tel1,tm_adress_basis.adb_fax, tm_adress_basis.adb_email " & _
                        "FROM tm_partei " & _
                        "INNER JOIN tm_user ON tm_user.cmuser_id = tm_partei.pa_id INNER JOIN DBA.tm_adress_basis ON tm_partei.pa_id = tm_adress_basis.pa_id LEFT OUTER JOIN tm_user_group ON tm_user.cmgroup_id = tm_user_group.cmgroup_id " & _
                        "WHERE adt_id = 1 AND cmuser_gesperrt = 0 AND tm_partei.pa_org_name2 LIKE '"
Public Const SelectMAName = "SELECT tm_partei.pa_org_name1, tm_partei.pa_pers_vorname, tm_partei.pa_org_name2, cmgroup_name " & _
                        "FROM tm_partei " & _
                        "INNER JOIN tm_user ON tm_user.cmuser_id = tm_partei.pa_id INNER JOIN DBA.tm_adress_basis ON tm_partei.pa_id = tm_adress_basis.pa_id LEFT OUTER JOIN tm_user_group ON tm_user.cmgroup_id = tm_user_group.cmgroup_id " & _
                        "WHERE adt_id = 1 AND cmuser_gesperrt = 0 AND tm_partei.pa_org_name1 LIKE "


' Constante für aktuelle Stückliste
Public Const cStueli = "SELECT tm_stueckliste_pos.stlp_nummer as Pos,tm_stueckliste_pos.stlp_menge as Menge, " & _
            "tm_artikel.ast_nummer as Artikel,tm_artikel_var.avar_nummer AS Var,tm_artikelvar_kd.avkd_artnr_kd as Kd_Nr,tm_artikel.ast_bezeichnung as Bez, " & _
            "tm_artikel.ast_merkmal_4 as Typ,tm_stueckliste_pos_memo.stlp_memo as Label, tm_artikel_var.avar_hersteller as Hersteller " & _
            "FROM (((((((tm_stueli_posvers INNER JOIN tm_stueckliste_pos ON tm_stueli_posvers.stlp_id = tm_stueckliste_pos.stlp_id) " & _
            "INNER JOIN tm_stueckliste ON tm_stueckliste.stl_id = tm_stueli_posvers.stl_id AND tm_stueckliste.stl_version = tm_stueli_posvers.stl_version) " & _
            "LEFT OUTER JOIN ta_mengeneinheit ON tm_stueckliste_pos.me_id = ta_mengeneinheit.me_id) " & _
            "LEFT OUTER JOIN tm_artikel ON tm_stueckliste_pos.stlp_member_id = tm_artikel.ast_id AND  (tm_stueckliste_pos.stlp_typ = 1 OR tm_stueckliste_pos.stlp_typ = 4)) " & _
            "LEFT OUTER JOIN tm_artikel_var ON tm_stueckliste_pos.stlp_member_id = tm_artikel_var.ast_id " & _
            "AND tm_stueckliste_pos.stlp_member_sub_id = tm_artikel_var.avar_id AND  (tm_stueckliste_pos.stlp_typ = 1 OR tm_stueckliste_pos.stlp_typ = 4)) " & _
            "LEFT OUTER JOIN tm_artikelvar_kd ON tm_artikel_var.ast_id = tm_artikelvar_kd.ast_id AND tm_artikel_var.avar_id = tm_artikelvar_kd.avar_id " & _
            "AND tm_artikelvar_kd.adk_id = ###AVISTA_KD_ID### AND tm_artikelvar_kd.avkd_default = 1 " & _
            "LEFT OUTER JOIN tm_arbeitsplatz ON tm_stueckliste_pos.stlp_member_id = tm_arbeitsplatz.arpl_id AND tm_stueckliste_pos.stlp_typ = 3) " & _
            "LEFT OUTER JOIN tm_artikel_var AS artvar2 ON tm_stueckliste.ast_id = artvar2.ast_id AND tm_stueckliste_pos.stlp_avar_id = artvar2.avar_id) " & _
            "LEFT OUTER JOIN tm_artikel AS art2 ON tm_stueckliste.ast_id = art2.ast_id " & _
            "LEFT OUTER JOIN tm_artikel_ek ON tm_stueckliste_pos.stlp_member_id = tm_artikel_ek.ast_id AND tm_artikel_ek.avar_id = tm_stueckliste_pos.stlp_member_sub_id " & _
            "LEFT OUTER JOIN tm_stueckliste_pos_memo ON tm_stueckliste_pos_memo.stlp_id = tm_stueckliste_pos.stlp_id " & _
            "LEFT OUTER JOIN xm_stueckliste_pos ON xm_stueckliste_pos.stlp_id = tm_stueckliste_pos.stlp_id " & _
            "LEFT OUTER JOIN um_qz_parse_hist ON um_qz_parse_hist.stlp_id = tm_stueckliste_pos.stlp_id " & _
            "WHERE tm_stueckliste_pos.stlp_gueltig_bis >= '###Datum###' AND tm_stueckliste_pos.stlp_gueltig_ab <= '###Datum###' " & _
            "AND art2.ast_nummer = '###Art###' ORDER BY Pos;"

' Constante für FA Stückliste
Public Const cFaStueli = " SELECT tt_auftrag_intern_pos.aint_pos_nummer as Pos, tt_auftrag_intern_pos.aint_pos_sollmenge as Menge, " & _
            "tm_artikel.ast_nummer as Artikel, tm_artikel_var.avar_nummer AS Var, tm_artikelvar_kd.avkd_artnr_kd as Kd_Nr, tm_artikel.ast_bezeichnung as Bez, " & _
            "tm_artikel.ast_merkmal_4 as Typ, tt_auftrag_intern_pos_text.aint_posm_memo as Label , tm_artikel_var.avar_hersteller as Hersteller " & _
            "FROM (((tt_auftrag_intern " & _
            "INNER JOIN tt_auftrag_intern_pos ON tt_auftrag_intern.aint_id = tt_auftrag_intern_pos.aint_id) " & _
            "INNER JOIN tm_artikel ON tt_auftrag_intern_pos.aint_pos_member_id = tm_artikel.ast_id AND (tt_auftrag_intern_pos.aint_pos_typ = 1 )) " & _
            "INNER JOIN tm_artikel_var ON tm_artikel.ast_id = tm_artikel_var.ast_id AND tt_auftrag_intern_pos.aint_pos_member_sub_id = tm_artikel_var.avar_id) " & _
            "LEFT OUTER JOIN tt_auftrag_intern_pos_text ON tt_auftrag_intern_pos.aint_id = tt_auftrag_intern_pos_text.aint_id " & _
            "AND tt_auftrag_intern_pos.aint_pos_id = tt_auftrag_intern_pos_text.aint_pos_id " & _
            "LEFT OUTER JOIN tm_artikelvar_kd ON tm_artikel_var.ast_id = tm_artikelvar_kd.ast_id AND tm_artikel_var.avar_id = tm_artikelvar_kd.avar_id " & _
            "AND tm_artikelvar_kd.adk_id = ###AVISTA_KD_ID### AND tm_artikelvar_kd.avkd_default = 1 " & _
            "WHERE tt_auftrag_intern_pos.aint_pos_alternative = 0 AND tt_auftrag_intern.vgst_id = 500 and tt_auftrag_intern.aint_nr LIKE '%###FA###%' " & _
            "ORDER BY aint_pos_anz_maschinen, aint_pos_order;"

Public Function GetDBConnection() As ADODB.Connection

    Dim cnn As ADODB.Connection

    Set cnn = New ADODB.Connection

    ' Open a connection using an ODBC DSN "Pubs".
    'cnn.ConnectionString = "DSN=av_test_he;UID=dba;PWD=phoenix;"
    cnn.ConnectionString = "DSN=avista;UID=dba;PWD=phoenix;"
    cnn.Open

    Set GetDBConnection = cnn
    'Set cnn = Nothing

End Function


Public Function closeDBConnection(ByRef fcnn As ADODB.Connection) As Boolean
    closeDBConnection = False
    fcnn.Close
    Set fcnn = Nothing
    closeDBConnection = True

End Function

'Public Function get_RecordsetAvista(f_sql As String, f_rs As ADODB.Recordset) As Boolean
Public Function get_RecordsetAvista(ByVal f_cnn As ADODB.Connection, f_sql As String) As ADODB.Recordset
On Error GoTo err_SetBG

    Dim rs As ADODB.Recordset

    'Set f_rs = cnn.Execute(AvistaKndArtNrSQL_Select & f_BG & "'" & AvistaArtikelSQL_Order)
    Set rs = f_cnn.Execute(f_sql)
    Set get_RecordsetAvista = rs
    'Set rs = Nothing

close_db:
   Exit Function

err_SetBG:
    MsgBox "error: " & Err & " - " & Error
    Resume Next

End Function

Public Function LQPP_IF_Interne_Fehlererfassung()
'Funktion liest in Avista die tabelle zlqpp_fed_daten_kopf aus und übernimmt alle eingetragen intern erfassten Reklamation in die Mastertab
'Return:
'OK:    Rekl-Nr wird in Tabellen-DS Avista eingetragen und Status 100 gesetzt
'NOK:   Status wird auf 200 gesetzt (Fehler) und in das ReklNr Datenfeld wird der Fehlercode eingetragen


'Fehlercodes
' FC1 = user_id aus Avista nicht in Access gefunden
' FC2 = Parameter Melderkürzel aus Avista nicht bekannt in Access
' FC3 = Fehler beim Ermitteln Baugruppendaten in Avista
' FC4 = Fehler beim Ermitteln des FA Daten anhand aint_id in Avista
' FC5 = Fehler beim Ermitteln Artikeldaten in Avista
' FC6 = Workstation konnten in Access nicht ermittelt werden
' FC7 = Tätigkeit konnte nicht ermittelt werden

    Const FC1 = 1
    Const FC2 = 2
    Const FC3 = 4
    Const FC4 = 8
    Const FC5 = 16
    Const FC6 = 32
    Const FC7 = 64
    Const FC8 = 128

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim dbAccess As DAO.Database
    'Recordset Avista
    Dim rsAvistaUser As ADODB.Recordset
    Dim rsAvistaUserData As ADODB.Recordset
    Dim rsAvistaOpenReklDS As ADODB.Recordset
    Dim rsAvistaFAData As ADODB.Recordset
    Dim rsAvistaBGData As ADODB.Recordset
    Dim rsAvistaArtikelData As ADODB.Recordset
    'Recordset Access
    Dim rsAccessUserID As Long
    Dim rsAccessBGID As Long
    Dim rsAccessArtikelID As Long
    Dim rsAccessWs As Long
    Dim rsAccessTaetID As Long

    Dim l_error As Integer
    Dim ls_sql As String

    l_error = 0
    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider
    'Daten aus Access auslesen
    Set dbAccess = openAccDBConnection("openDB")

    'Erster Schritt alle offenen Reklamationen aus LQPP ermitteln
    Set rsAvistaOpenReklDS = adbprov.getRecordsetAvista("SELECT * FROM zlqpp_fe_daten_kopf where [fed_status] = 10;")
    If rsAvistaOpenReklDS.RecordCount Then
        rsAvistaOpenReklDS.MoveFirst
        While Not rsAvistaOpenReklDS.EOF
            'Master Insert Tabelle initialisieren
            'Drop Wertetabelle
            dbAccess.Execute ("DROP TABLE InsertTabTemplateWerte;")
            DoCmd.CopyObject , "InsertTabTemplateWerte", acTable, "InsertTabTemplate"


            '#################################################################
            '######## suche von Melderdaten in Acces über cmuser_id in Avista
            '#################################################################

            '1. User in Access anhand Avista User_ID ermitteln

            rsAccessUserID = DLookup("[Kunden-ID]", "Kunden", "[Avista-ID] = " & rsAvistaOpenReklDS("cmuser_id"))
            If IsNull(rsAccessUserID) Then
                Set rsAvistaUser = adbprov.getRecordsetAvista("SELECT pa_org_name2 FROM tm_partei where pa_id = " & rsAvistaOpenReklDS("cmuser_id") & ";")
                If rsAvistaUser.RecordCount Then
                    rsAccessUserID = DLookup("Kunden-ID", "Kunden", "[Kurzzeichen] = '" & rsAvistaUser("pa_org_name2") & "' and [Kundennummer] = 01052 and [Aktiv] = True")
                    If IsNull(rsAccessUserID) Then
                        'Anlegen User in Kundentabelle
                        'Ermitteln aller Melderdaten aus Avista
                        Set rsAvistaUserData = adbprov.getRecordsetAvista("SELECT tm_partei.pa_id, tm_partei.pa_org_name1, tm_partei.pa_pers_vorname, tm_partei.pa_org_name2, " & _
                                                                  "cmgroup_name, tm_adress_basis.adb_tel1,tm_adress_basis.adb_fax, tm_adress_basis.adb_email " & _
                                                                  "From tm_partei " & _
                                                                  "INNER JOIN tm_user ON tm_user.cmuser_id = tm_partei.pa_id INNER JOIN DBA.tm_adress_basis ON tm_partei.pa_id = tm_adress_basis.pa_id " & _
                                                                  "LEFT OUTER JOIN tm_user_group ON tm_user.cmgroup_id = tm_user_group.cmgroup_id " & _
                                                                  "WHERE adt_id = 1 AND cmuser_gesperrt = 0 AND tm_partei.pa_org_name2 LIKE '" & rsAvistaUser("pa_org_name2") & "';")
                        If rsAvistaUser.RecordCount Then
                            'Ermittelte  Melderdaten in Kundentabelle eintragen
                            dbAccess.Execute ("Insert into Kunden ([Kontaktperson], [Firma], [Kurzzeichen], [Telefonnummer], [EmailAdresse], [Anmerkungen], [FirmaID], [Kundennummer], " & _
                                              "[Avista-ID]) " & _
                                              "values ('" & rsAvistaUserData("pa_org_name1") & "', '" & rsAvistaUserData("pa_pers_vorname") & "', 'LQ Mechatronik-Systeme GmbH', '" & _
                                              rsAvistaUserData("pa_org_name2") & "', '" & rsAvistaUserData("adb_tel1") & "', '" & rsAvistaUserData("adb_email") & "', '" & _
                                              rsAvistaUserData("cmgroup_name") & "', 62, '01052', " & rsAvistaUserData("pa_id") & ");")
                            'User Id des neu angelegten Melder ermitteln
                            rsAccessUserID = DLookup("Kunden-ID", "Kunden", "[Kurzzeichen] = '" & rsAvistaUser("pa_org_name2") & "' and [Kundennummer] = 01052 and [Aktiv] = True")
                        Else
                            MsgBox "Fehler beim Ermitteln des Melderdaten anahnd des Melderkürzel in Avista!"
                            l_error = l_error + FC1
                        End If
                    End If
                Else
                    MsgBox "Fehler beim Ermitteln des Melderkürzel anhand user_id in Avista!"
                    l_error = l_error + FC2
                End If
            End If
            If Not IsNull(rsAccessUserID) Then
                dbAccess.Execute ("update InsertTabTemplateWerte set [FieldValue] = '" & rsAccessUserID & "' Where [FieldName] = '[ErfasserIntern]';")
                dbAccess.Execute ("update InsertTabTemplateWerte set [FieldValue] = '" & rsAccessUserID & "' Where [FieldName] = '[Kunden-ID]';")
            End If
            '#################################################################
            '######## Ermitteln aller FA-spezifischen Daten
            '#################################################################
            'Ermitteln Baugruppen-Daten FA
            Set rsAvistaFAData = adbprov.getRecordsetAvista("SELECT tt_auftrag_intern.aint_nr as [FA-Nr], SUM(tt_auftrag_intern.aint_bedarf_menge) as Losgroesse, tm_artikel.ast_nummer, " & _
                                                            "tm_artikel_var.avar_nummer, tt_auftrag_intern.avar_id , tt_auftrag_intern.ast_id " & _
                                                            "From DBA.tt_auftrag_intern " & _
                                                            "LEFT OUTER JOIN tm_artikel ON tt_auftrag_intern.ast_id = tm_artikel.ast_id " & _
                                                            "LEFT OUTER JOIN tm_artikel_var ON tt_auftrag_intern.ast_id = tm_artikel_var.ast_id AND " & _
                                                            "tt_auftrag_intern.avar_id = tm_artikel_var.avar_id " & _
                                                            "WHERE tt_auftrag_intern.aint_id = " & rsAvistaOpenReklDS("aint_id") & _
                                                            "GROUP BY tt_auftrag_intern.aint_nr, tt_auftrag_intern.avar_id, tt_auftrag_intern.ast_id, tm_artikel.ast_nummer, tm_artikel_var.avar_nummer, tt_auftrag_intern.aint_id;")
            If rsAvistaFAData.RecordCount Then
                rsAccessBGID = Nz(DLookup("[Baugruppen-ID]", "[Baugruppen-avista_out]", "[LQ Art-Nr] = '" & rsAvistaFAData("ast_nummer") & "' AND [Var Nr] = '" & rsAvistaFAData("avar_nummer") & "'"))
                If IsNull(rsAccessBGID) Or rsAccessBGID = 0 Then
                    'Baugruppe Anlegen
                    'Ermitteln aller Baugruppendaten aus Avista

                    Set rsAvistaBGData = adbprov.getRecordsetAvista("SELECT tm_artikel.ast_id, tm_artikelvar_kd.avar_id, tm_artikel.ast_nummer, tm_artikel_var.avar_nummer, " & _
                                                                    "tm_artikel_var.avar_herst_nummer,  tm_artikel.ast_bezeichnung, va_produktgruppe.pg_bez " & _
                                                                    "From tm_artikel " & _
                                                                    "INNER JOIN tm_artikel_var ON tm_artikel.ast_id = tm_artikel_var.ast_id " & _
                                                                    "INNER JOIN tx_artikelstatus ON tm_artikel.artstat_id = tx_artikelstatus.artstat_id " & _
                                                                    "LEFT OUTER JOIN tm_artikelvar_kd ON tm_artikelvar_kd.ast_id = tm_artikel.ast_id AND " & _
                                                                    "tm_artikelvar_kd.avar_id = tm_artikel_var.avar_id AND tm_artikelvar_kd.adk_id = 73 " & _
                                                                    "LEFT OUTER JOIN va_produktgruppe ON tm_artikel.pg_id = va_produktgruppe.pg_id AND va_produktgruppe.pg_gueltig = 1 AND " & _
                                                                    "va_produktgruppe.ma_id = 2 AND va_produktgruppe.sprach_id = 1 " & _
                                                                    "WHERE tx_artikelstatus.artstat_archiv = 0 AND tm_artikel.ast_id = " & rsAvistaFAData("ast_id") & _
                                                                    " AND tm_artikelvar_kd.avar_id = " & rsAvistaFAData("avar_id") & ";")
                    If rsAvistaBGData.RecordCount Then
                        'Baugruppe in Access anlegen
                        dbAccess.Execute ("Insert into [Baugruppen-avista_out] ([LQ Art-Nr], [Var Nr], [Kunden-Nr], [Bezeichnung 1]) " & _
                                          "values ('" & rsAvistaBGData("ast_nummer") & "', '" & rsAvistaBGData("avar_nummer") & "', '" & rsAvistaBGData("avar_id") & _
                                          "', '" & rsAvistaBGData("ast_bezeichnung") & "');")
                        'Baugruppen-ID ermitteln
                        rsAccessBGID = DLookup("[Baugruppen-ID]", "Baugruppen-avista_out", "[LQ Art-Nr] = '" & rsAvistaFAData("ast_nummer") & "' and [Var Nr] ='" & rsAvistaFAData("avar_nummer") & "'")
                    Else
                        MsgBox "Fehler beim Ermitteln Baugruppendaten in Avista!"
                        l_error = l_error + FC3
                    End If

                End If
                dbAccess.Execute ("update InsertTabTemplateWerte set [FieldValue] = '" & rsAvistaFAData("FA-Nr") & "' Where [FieldName] = '[Fertigungsauftrag-FA]';")
                dbAccess.Execute ("update InsertTabTemplateWerte set [FieldValue] = '" & rsAvistaOpenReklDS("fed_menge") & "' Where [FieldName] = '[Menge betroffen - LQ]';")
                dbAccess.Execute ("update InsertTabTemplateWerte set [FieldValue] = '" & rsAvistaFAData("Losgroesse") & "' Where [FieldName] = '[FALosgroesse]';")
            Else
                MsgBox "Fehler beim Ermitteln des FA Daten anhand aint_id in Avista!"
                l_error = l_error + FC4
            End If
            If Not IsNull(rsAccessBGID) Then
                dbAccess.Execute ("update InsertTabTemplateWerte set [FieldValue] = '" & rsAccessBGID & "' Where [FieldName] = '[Baugruppen-ID]';")
            End If

            'Ermitteln Artikel-Daten des betroffenen Artikel in der Meldung
            rsAccessArtikelID = Nz(DLookup("[Artikel-ID]", "Artikel-all_avista_out", "[LQ Art-Nr] = '" & rsAvistaFAData("ast_nummer") & "' and [Variante] ='" & rsAvistaFAData("avar_nummer") & "'"))
            If IsNull(rsAccessArtikelID) Or rsAccessArtikelID = 0 Then
                'Anlegen Artikel
                'Ermitteln aller Artikeldaten aus Avista

                Set rsAvistaArtikelData = adbprov.getRecordsetAvista("SELECT tm_artikel.ast_id, tm_artikelvar_kd.avar_id, tm_artikel.ast_nummer, tm_artikel_var.avar_nummer, " & _
                                                                "tm_artikel_var.avar_herst_nummer,  tm_artikel.ast_bezeichnung, va_produktgruppe.pg_bez " & _
                                                                "From tm_artikel " & _
                                                                "INNER JOIN tm_artikel_var ON tm_artikel.ast_id = tm_artikel_var.ast_id " & _
                                                                "INNER JOIN tx_artikelstatus ON tm_artikel.artstat_id = tx_artikelstatus.artstat_id " & _
                                                                "LEFT OUTER JOIN tm_artikelvar_kd ON tm_artikelvar_kd.ast_id = tm_artikel.ast_id AND " & _
                                                                "tm_artikelvar_kd.avar_id = tm_artikel_var.avar_id AND tm_artikelvar_kd.adk_id = 73 " & _
                                                                "LEFT OUTER JOIN va_produktgruppe ON tm_artikel.pg_id = va_produktgruppe.pg_id AND va_produktgruppe.pg_gueltig = 1 AND " & _
                                                                "va_produktgruppe.ma_id = 2 AND va_produktgruppe.sprach_id = 1 " & _
                                                                "WHERE tx_artikelstatus.artstat_archiv = 0 AND tm_artikel.ast_id = " & rsAvistaBGData("ast_id") & _
                                                                " AND tm_artikelvar_kd.avar_id = " & rsAvistaBGData("avar_id") & ";")
                If rsAvistaArtikelData.RecordCount Then
                    'Artikel in Access eintragen
                    dbAccess.Execute ("Insert into [Artikel-all_avista_out] ([LQ Art-Nr], Variante, [Kd Mat-Nr], [Bezeichnung 1], [Hersteller-Teilenummer], [AvistaArtikelTyp]) " & _
                                      "values ('" & rsAvistaArtikelData("ast_nummer") & "', '" & rsAvistaArtikelData("avar_nummer") & "', '" & rsAvistaArtikelData("avar_id") & _
                                      "', '" & rsAvistaArtikelData("ast_bezeichnung") & "', '" & rsAvistaArtikelData("avar_herst_nummer") & "', 0);")
                    'Neu angelegter Artikel ermitteln
                    rsAccessArtikelID = DLookup("[Artikel-ID]", "[Artikel-all_avista_out]", "[LQ Art-Nr] = '" & rsAvistaFAData("ast_nummer") & "' and [Variante] ='" & rsAvistaFAData("avar_nummer") & "'")
                Else
                    MsgBox "Fehler beim Ermitteln Artikeldaten in Avista!"
                    l_error = l_error + FC5
                End If

            End If
            If Not IsNull(rsAccessArtikelID) Then
                dbAccess.Execute ("update InsertTabTemplateWerte  set [FieldValue] = '" & rsAccessArtikelID & "' Where [FieldName] = '[Artikel-ID 1]';")
            End If
            '#################################################################
            '######## Worksationermitteln
            '#################################################################

            rsAccessWs = DLookup("ID", "LQWorkstation", "[WSAvistaID] = " & rsAvistaOpenReklDS("lqpp_ws_id"))
            If IsNull(rsAccessWs) Then
                l_error = l_error + FC6
            Else
                dbAccess.Execute ("update InsertTabTemplateWerte set [FieldValue] = '" & rsAccessWs & "' Where [FieldName] = '[LQWorkstationErfasserID]';")
            End If

            '##################################################################
            '########## Weitere Parameter in Insert Tabelle eintragen
            '#################################################################
            dbAccess.Execute ("update InsertTabTemplateWerte set [FieldValue] = '" & rsAvistaOpenReklDS("feb_id") & "' Where [FieldName] = '[LQBereichErfasserID]';")
            dbAccess.Execute ("update InsertTabTemplateWerte set [FieldValue] = '" & rsAvistaOpenReklDS("fee_id") & "' Where [FieldName] = '[LQEinheitErfasserID]';")
            'Ermitteln Tätigkeit anhand Fehlerbild
            rsAccessTaetID = DLookup("TaetigkeitID", "ZuordnungDefTaetigkeit", "[ID] = " & rsAvistaOpenReklDS("fef_id"))
            If Not IsNull(rsAccessTaetID) Then
                dbAccess.Execute ("update InsertTabTemplateWerte set [FieldValue] = '" & rsAccessTaetID & "' Where [FieldName] = '[TaetigkeitErfasserID]';")
                dbAccess.Execute ("update InsertTabTemplateWerte Set [FieldValue] = '" & rsAvistaOpenReklDS("fef_id") & "' Where [FieldName] = '[DefTaetigkeitErfasserID]';")
            Else
                l_error = l_error + FC7
            End If
            'Umrechnen Min in Std.
            Dim l_z  As Double
            Dim l_h As Integer
            Dim l_m As Double
            l_h = rsAvistaOpenReklDS("fed_zeitaufwand") / 60
            l_m = rsAvistaOpenReklDS("fed_zeitaufwand") Mod 60
            l_z = (l_m / 0.6) / 100
            dbAccess.Execute ("update InsertTabTemplateWerte set [FieldValue] = '" & Replace(CStr(l_z), ",", ".") & "' Where [FieldName] = '[PE-Bearbeitung]';")
            dbAccess.Execute ("update InsertTabTemplateWerte set [FieldValue] = '" & rsAvistaOpenReklDS("fed_info_text") & "' Where [FieldName] = '[Welcher Fehler tritt auf]';")
            dbAccess.Execute ("update InsertTabTemplateWerte set [FieldValue] = '" & rsAvistaOpenReklDS("fed_datum") & "' Where [FieldName] = '[Datum]';")
            dbAccess.Execute ("update InsertTabTemplateWerte set [FieldValue] = '" & rsAvistaOpenReklDS("fed_datum") & "' Where [FieldName] = '[Auftragsdatum]';")


            '#################################################################
            '######## Fehlerticket erzeugen
            '#################################################################
            '1. Max Rekl-ID ermitteln
            p_ReklNr = getLastDSMastertab()
            If get_p_ReklNr() > 0 Then
                'ReklNr erhöhen = Neuer Vorgang
                p_ReklNr = get_p_ReklNr() + 1
                dbAccess.Execute ("update InsertTabTemplateWerte set [FieldValue] = '" & p_ReklNr & "' Where [FieldName] = '[ID]';")
            End If

            ls_sql = getSQLStatInsertTable("Mastertab", "Insert")
            'Eintrag Reklamation in die Mastertab
            dbAccess.Execute (ls_sql)
            'Zeitstempel der Erfassung in die Status Wechseltabelle eintragen
            dbAccess.Execute ("Insert into tblStatusWechsel ([ReklNr],[Datum],[Status],[Kürzel]) " & _
                              "values (" & p_ReklNr & ", " & _
                                        "'" & format(rsAvistaOpenReklDS("fed_datum"), "dd.mm.yyyy hh:mm:ss") & "', " & _
                                        "'Vorgang Eingang Intern', " & _
                                        "'" & DLookup("[Kurzzeichen]", "Kunden", "[Avista-ID] = " & rsAvistaOpenReklDS("cmuser_id")) & "'" & _
                                       ");")

            InsertNewReklMemoDS (get_p_ReklNr())
            InsertNewEinkaufDS (get_p_ReklNr())
            'ReklID in Avista als OK eintragen
            If l_error = 0 Then
                adbprov.CommandExecuteAvista ("Update zlqpp_fe_daten_Kopf set [fed_rekla_id] = " & p_ReklNr & ", [fed_status] = 100 where fed_id = " & rsAvistaOpenReklDS("fed_id") & ";")
                'Statsu in Avista eintragen
            Else
                'Fehlercode und Status in Avista eintragen
                adbprov.CommandExecuteAvista ("Update zlqpp_fe_daten_Kopf set [fed_rekla_id] = " & l_error & ", [fed_status]  = 200 where fed_id = " & rsAvistaOpenReklDS("fed_id") & ";")
                'Fehlercode in Tabelle tblFehlerCodeLQPP eintragen
                DecToBin (l_error), rsAvistaOpenReklDS("fed_id")
            End If
            rsAvistaOpenReklDS.MoveNext
        Wend

    End If

    'Recordset Avista
    Set rsAvistaUser = Nothing
    Set rsAvistaUserData = Nothing
    Set rsAvistaOpenReklDS = Nothing
    Set rsAvistaFAData = Nothing
    Set rsAvistaBGData = Nothing
    Set rsAvistaArtikelData = Nothing

    Set adbprov = Nothing

End Function

Public Function DecToBin(lngDecimal As Long, ID_Avista As Long) As String

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim strBinaer As String
    Dim i As Integer
    Dim Feldname As String

    Set ldb = openAccDBConnection("openDB")

    ldb.Execute ("Insert into tblFehlerCodeLQPP (AvistaID) values (" & ID_Avista & ");")
    Set rs = ldb.OpenRecordset("Select * from tblFehlerCodeLQPP where [AvistaID] = " & ID_Avista & " ;")
    If rs.RecordCount Then
        Do While lngDecimal > 0
            i = i + 1
            rs.Edit
            Feldname = "FC" & i
            If lngDecimal Mod 2 = 1 Then
                strBinaer = "1" & strBinaer
                lngDecimal = lngDecimal - 1
                lngDecimal = lngDecimal / 2
                rs(Feldname) = True
            Else
                strBinaer = "0" & strBinaer
                lngDecimal = lngDecimal / 2
                rs(Feldname) = False
            End If
            rs.Update
        Loop
    End If
    Set rs = Nothing

End Function
