Attribute VB_Name = "Konstanten"
Option Compare Database
Option Explicit

'********** Bereich automatische Aufgabe versenden *********
Public Const cBetreff = "!Textwert! (!ReklNr!)" & " - "
Public Const cTrennStrich = "________________________________________"
Public Const cVAR_Kuerzel = "Kuerzel"
Public Const cVAR_LoggingClient = "LoggingClient"       '20180124-ROS
Public Const cVAR_LoggingClientAll = "LoggingClientAll"     '20180124-ROS

'#########################################################################
'20231020-ROS
'Reklamations-ID Konstante für Fehlererfassung über LQPP, Lokalisiert
Public Const cReklID_LQPP_DE = 500000
Public Const cReklID_LQPP_CN = 3000000
Public Const cReklID_LQPP_EN = 5000000


#Const SystemDebug = 1

'##############################################################################################################################################################################
'Konstanten für den Abruf der Teileverwendung
Public Const cs_TVN = "SELECT tm_artikel.ast_id, tm_artikel_var.avar_id " & _
                "FROM (DBA.tm_artikel INNER JOIN DBA.tm_artikel_text ON tm_artikel.ast_id = tm_artikel_text.ast_id " & _
                "INNER JOIN DBA.tm_artikel_zuord ON tm_artikel_zuord.ast_id = tm_artikel.ast_id) " & _
                "INNER JOIN DBA.tm_artikel_var ON tm_artikel_var.ast_id = tm_artikel.ast_id " & _
                "WHERE sprach_id = 1 AND atxt_typ=1 AND tm_artikel.ast_nummer = '###ArtNr###';"


Public Const cs_TVN_1 = "SELECT tm_artikel.ast_id, tm_artikel_var.avar_id,tm_artikel_var.avar_nummer " & _
                "FROM (DBA.tm_artikel INNER JOIN DBA.tm_artikel_text ON tm_artikel.ast_id = tm_artikel_text.ast_id " & _
                "INNER JOIN DBA.tm_artikel_zuord ON tm_artikel_zuord.ast_id = tm_artikel.ast_id) " & _
                "INNER JOIN DBA.tm_artikel_var ON tm_artikel_var.ast_id = tm_artikel.ast_id " & _
                "WHERE sprach_id = 1 AND atxt_typ=1 AND tm_artikel.ast_nummer = '###ArtNr###'and avar_nummer = '###ArtVar###';"

Public Const cs_TVN_2 = "SELECT DISTINCT tm_artikel.ast_nummer as ArtNr, tm_artikel.ast_bezeichnung as Bez, tm_artikel.ast_merkmal_4 as Merkmal " & _
                "FROM ((( DBA.tm_stueckliste INNER JOIN DBA.tm_stueli_posvers ON tm_stueckliste.stl_id = tm_stueli_posvers.stl_id) " & _
                "INNER JOIN DBA.tm_stueckliste_pos ON tm_stueli_posvers.stlp_id = tm_stueckliste_pos.stlp_id) " & _
                "INNER JOIN DBA.tm_artikel ON tm_stueckliste.ast_id = tm_artikel.ast_id ) " & _
                "INNER JOIN DBA.tx_artikelstatus ON tm_artikel.artstat_id = tx_artikelstatus.artstat_id " & _
                "WHERE tx_artikelstatus.artstat_archiv = 0 And tm_stueckliste_pos.stlp_member_id = ###ast_id### And tm_stueckliste_pos.stlp_member_sub_id = ###avar_id### " & _
                "ORDER BY ast_nummer;"
Public Const cs_TVN_3 = "SELECT DISTINCT tm_artikel.ast_nummer as ArtNr, tm_artikel.ast_bezeichnung as Bez, tm_artikel.ast_merkmal_4 as Merkmal " & _
                "FROM ((( DBA.tm_stueckliste INNER JOIN DBA.tm_stueli_posvers ON tm_stueckliste.stl_id = tm_stueli_posvers.stl_id) " & _
                "INNER JOIN DBA.tm_stueckliste_pos ON tm_stueli_posvers.stlp_id = tm_stueckliste_pos.stlp_id) " & _
                "INNER JOIN DBA.tm_artikel ON tm_stueckliste.ast_id = tm_artikel.ast_id ) " & _
                "INNER JOIN DBA.tx_artikelstatus ON tm_artikel.artstat_id = tx_artikelstatus.artstat_id " & _
                "WHERE tx_artikelstatus.artstat_archiv = 0 And tm_stueckliste_pos.stlp_member_id = ###ast_id### " & _
                "ORDER BY ast_nummer;"
'###############################################################################################################################################################################
'Konstanten zum Abrufen Infos zu Seriennummern
 Public Const cs_SNR_1 = "SELECT aint_nr as FA, ut_ppaf_fa_snr_hist.af_snr_text as SNR, avkd_artnr_kd as KdNr " & _
                        "From DBA.ut_ppaf_fa_snr_status " & _
                        "INNER JOIN DBA.tt_auftrag_intern ON tt_auftrag_intern.aint_id = ut_ppaf_fa_snr_status.aint_id " & _
                        "INNER JOIN DBA.ut_ppaf_fa_snr_hist ON ut_ppaf_fa_snr_status.aint_id = ut_ppaf_fa_snr_hist.aint_id " & _
                        "LEFT JOIN DBA.tm_artikelvar_kd ON tm_artikelvar_kd.ast_id = ut_ppaf_fa_snr_hist.ast_id " & _
                        "AND tm_artikelvar_kd.avar_id = ut_ppaf_fa_snr_hist.avar_id AND tm_artikelvar_kd.adk_id = 73 " & _
                        "WHERE ut_ppaf_fa_snr_hist.af_snr_text like '%###SNR###%';"
'%407877652%
 Public Const cs_SNR_2 = "SELECT aint_nummer_in as FA,tm_seriennum_pos.snrp_seriennummer as SNR, tm_artikel.ast_nummer + '-' + tm_artikel_var.avar_nummer as MatNr," & _
                        "tm_artikel.ast_bezeichnung as Bez, tbl_kunde.pa_org_name1 AS Kunde, lief_nummer as LI " & _
                        "FROM (((DBA.tm_seriennum_pos INNER JOIN DBA.tm_artikel_var ON tm_seriennum_pos.ast_id = tm_artikel_var.ast_id AND tm_seriennum_pos.avar_id = tm_artikel_var.avar_id) " & _
                        "INNER JOIN DBA.tm_artikel ON tm_seriennum_pos.ast_id = tm_artikel.ast_id) " & _
                        "LEFT OUTER JOIN DBA.tm_partei AS tbl_kunde ON tm_seriennum_pos.adk_id = tbl_kunde.pa_id) " & _
                        "WHERE tm_seriennum_pos.snrp_seriennummer like '%###SNR###%' " & _
                        "ORDER BY ast_nummer, avar_nummer, snrp_seriennummer DESC;"
'%2311200013189-LQ%
'##########################################################################################################################################################################################
