VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAvistaProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'Const cs_sql_ListBG_BMK_KN = "SELECt tm_artikel.*, tm_artikel_var.*, ean_code, ean_type, tm_artikel.ast_bezeichnung AS artikel_bezeichnung, tm_artikelvar_kd.avkd_artnr_kd, tm_artikel.ast_nummer AS ast_nummer_ueb, tx_artikelstatus.artstat_archiv, " & _
'                            "tm_artikel.ast_id, tm_artikel_var.wanr_key, tm_artikel_var.avar_zolltarifnr, tm_artikel_var.avar_nummer AS avar_nummer_ueb, tm_artikelvar_kd.adk_id AS adkid, tm_kunde.adk_nummer, tm_artvar_lief.avli_artnr_liefer, " & _
'                            "tm_sml_daten.sml_id, pa_org_name1, xm_artikel.ast_lqadapter_1q, xm_artikel.ast_lqadapter_2q, xm_artikel.ast_lqadapter_1z, xm_artikel.ast_lqadapter_2z, xm_artikel.ast_ul_kennung, tm_stueckliste.stuelistat_id, " & _
'                            "tm_stueckliste.stl_version, xm_artikel.fa_bemerkung, xm_artikel_var.avar_c14_import_kennung " & _
'                            "FROM (((DBA.tm_artikel LEFT OUTER JOIN DBA.xm_artikel ON xm_artikel.ast_id = tm_artikel.ast_id " & _
'                            "INNER JOIN DBA.tm_artikel_var ON tm_artikel.ast_id = tm_artikel_var.ast_id) " & _
'                            "LEFT OUTER JOIN DBA.tm_artikel_var_ean ON tm_artikel_var.ast_id = tm_artikel_var_ean.ast_id AND tm_artikel_var.avar_id = tm_artikel_var_ean.avar_id) " & _
'                            "LEFT OUTER JOIN DBA.tm_artikelvar_kd ON tm_artikelvar_kd.ast_id = tm_artikel.ast_id AND tm_artikelvar_kd.avar_id = tm_artikel_var.avar_id AND tm_artikelvar_kd.adk_id = 73 " & _
'                            "LEFT OUTER JOIN DBA.tm_partei ON tm_partei.pa_id = tm_artikelvar_kd.adk_id " & _
'                            "LEFT OUTER JOIN DBA.tm_kunde ON tm_artikelvar_kd.adk_id = tm_kunde.adk_id AND tm_kunde.ma_id = 2 " & _
'                            "LEFT OUTER JOIN DBA.tm_sml_daten ON tm_artikel.ast_id =tm_sml_daten.ast_id AND tm_sml_daten.sprach_id = 1) " & _
'                            "INNER JOIN DBA.tx_artikelstatus ON tm_artikel.artstat_id = tx_artikelstatus.artstat_id " & _
'                            "LEFT OUTER JOIN DBA.xm_artikel_var ON tm_artikel_var.ast_id = xm_artikel_var.ast_id AND tm_artikel_var.avar_id = xm_artikel_var.avar_id " & _
'                            "LEFT OUTER JOIN DBA.tm_artvar_lief ON tm_artikel_var.ast_id = tm_artvar_lief.ast_id AND tm_artikel_var.avar_id = tm_artvar_lief.avar_id AND tm_artvar_lief.avli_prioritaet = 1 " & _
'                            "LEFT OUTER JOIN DBA.tm_stueckliste ON tm_artikel.ast_id = tm_stueckliste.ast_id  WHERE ast_nummer <> '-' AND tm_artikel.ast_id " & _
'                            "IN ( SELECT tm_artikelvar_kd.ast_id FROM (DBA.tm_artikelvar_kd INNER JOIN DBA.tm_partei ON tm_artikelvar_kd.adk_id = tm_partei.pa_id ) " & _
'                            "INNER JOIN  DBA.tm_kunde ON tm_artikelvar_kd.adk_id = tm_kunde.adk_id " & _
'                            "LEFT OUTER JOIN DBA.xm_artikelvar_kd ON tm_artikelvar_kd.ast_id = xm_artikelvar_kd.ast_id AND tm_artikelvar_kd.avar_id = xm_artikelvar_kd.avar_id AND tm_artikelvar_kd.adk_id = xm_artikelvar_kd.adk_id " & _
'                            "AND tm_artikelvar_kd.avkd_artnr_kd = xm_artikelvar_kd.avkd_artnr_kd " & _
'                            "WHERE tm_artikelvar_kd.avkd_artnr_kd LIKE '%###MaschNr###%' ) AND tm_artikel.ast_id " & _
'                            "IN ( SELECT tm_stueckliste.ast_id FROM (DBA.tm_stueli_posvers INNER JOIN DBA.tm_stueckliste_pos ON tm_stueli_posvers.stlp_id = tm_stueckliste_pos.stlp_id " & _
'                            "INNER JOIN DBA.tm_stueckliste ON tm_stueckliste.stl_id = tm_stueli_posvers.stl_id AND tm_stueckliste.stl_version = tm_stueli_posvers.stl_version) " & _
'                            "LEFT OUTER JOIN DBA.tm_stueckliste_pos_memo ON tm_stueckliste_pos.stlp_id = tm_stueckliste_pos_memo.stlp_id " & _
'                            "WHERE stlp_gueltig_ab <= '2021-07-05' AND stlp_gueltig_bis >= '2021-07-05' AND tm_stueckliste_pos_memo.stlp_memo LIKE '%###BMK###%' ) " & _
'                            "AND tx_artikelstatus.artstat_archiv = 0 AND tm_artikel.ast_id <> 7688"

Const cs_sql_ListBG_BMK_KN = "SELECT ast_nummer + '-' + avar_nummer AS nr, tm_artikel.*, tm_artikel_var.avar_id, tm_artikel_var.*, ean_code, ean_type, tm_artikel.ast_bezeichnung " & _
                            "AS artikel_bezeichnung, tm_artikelvar_kd.avkd_artnr_kd, " & _
                            "tm_artikelvar_kd.avkd_default, tm_artikel.ast_nummer AS ast_nummer_ueb, tx_artikelstatus.artstat_archiv, smld_text01, smld_text02, smld_text03, smld_text04, smld_text05, " & _
                            "smld_text06, smld_text07, smld_text08, smld_text09, smld_text10, smld_text11, smld_text12, smld_text13, smld_text14, smld_text15, smld_text16, smld_text17, smld_text18, " & _
                            "smld_text19, smld_text20, smld_text21, smld_text22, smld_text23, smld_text24, smld_text25, smld_text26, smld_text27, smld_text28, smld_text29, smld_text30, smld_text31, " & _
                            "smld_text32, smld_text33, smld_text34, smld_text35, smld_text36, smld_text37, smld_text38, smld_text39, smld_text40,tm_artikel.ast_id, tm_artikel_var.wanr_key, " & _
                            "tm_artikel_var.avar_zolltarifnr, tm_artikel_var.avar_nummer AS avar_nummer_ueb, tm_artikelvar_kd.adk_id AS adkid, pa_org_name1, xm_artikel_var.avar_matnr_etik_anz, " & _
                            "xm_artikel_var.avar_fa_text, tm_artvar_lief.avli_artnr_liefer, xm_artikel.ast_lqadapter_1q, xm_artikel.ast_lqadapter_2q, xm_artikel.ast_lqadapter_1z,xm_artikel.ast_lqadapter_2z , " & _
                            "tm_stueckliste.stuelistat_id , xm_artikel.ast_rohteile_kategorie, xm_artikel.ast_dlz_sum " & _
                            "FROM (((DBA.tm_artikel LEFT OUTER JOIN DBA.xm_artikel ON xm_artikel.ast_id = tm_artikel.ast_id " & _
                            "INNER JOIN DBA.tm_artikel_var ON tm_artikel.ast_id = tm_artikel_var.ast_id) " & _
                            "LEFT OUTER JOIN DBA.tm_artikel_var_ean ON tm_artikel_var.ast_id = tm_artikel_var_ean.ast_id AND tm_artikel_var.avar_id = tm_artikel_var_ean.avar_id) " & _
                            "LEFT OUTER JOIN DBA.tm_artikelvar_kd ON tm_artikelvar_kd.ast_id = tm_artikel.ast_id AND tm_artikelvar_kd.avar_id = tm_artikel_var.avar_id AND tm_artikelvar_kd.adk_id = 73 " & _
                            "LEFT OUTER JOIN DBA.tm_partei ON tm_partei.pa_id = tm_artikelvar_kd.adk_id " & _
                            "LEFT OUTER JOIN DBA.tm_sml_daten ON tm_artikel.ast_id =tm_sml_daten.ast_id AND sprach_id = 1) " & _
                            "INNER JOIN DBA.tx_artikelstatus ON tm_artikel.artstat_id = tx_artikelstatus.artstat_id LEFT OUTER JOIN DBA.xm_artikel_var ON tm_artikel_var.ast_id = xm_artikel_var.ast_id AND tm_artikel_var.avar_id = xm_artikel_var.avar_id " & _
                            "LEFT OUTER JOIN DBA.tm_artvar_lief ON tm_artikel_var.ast_id = tm_artvar_lief.ast_id AND tm_artikel_var.avar_id = tm_artvar_lief.avar_id AND tm_artvar_lief.avli_prioritaet = 1 " & _
                            "LEFT OUTER JOIN DBA.tm_stueckliste ON tm_artikel.ast_id = tm_stueckliste.ast_id WHERE ast_nummer <> '-' AND tm_artikel.ast_id " & _
                            "IN ( SELECT tm_artikelvar_kd.ast_id FROM (DBA.tm_artikelvar_kd INNER JOIN DBA.tm_partei ON tm_artikelvar_kd.adk_id = tm_partei.pa_id )INNER JOIN  DBA.tm_kunde ON tm_artikelvar_kd.adk_id = tm_kunde.adk_id " & _
                            "LEFT OUTER JOIN DBA.xm_artikelvar_kd ON tm_artikelvar_kd.ast_id = xm_artikelvar_kd.ast_id AND tm_artikelvar_kd.avar_id = xm_artikelvar_kd.avar_id AND tm_artikelvar_kd.adk_id = xm_artikelvar_kd.adk_id " & _
                            "AND tm_artikelvar_kd.avkd_artnr_kd = xm_artikelvar_kd.avkd_artnr_kd  WHERE tm_kunde.ma_id = 2 AND tm_artikelvar_kd.avkd_artnr_kd LIKE '%###MaschNr###%') " & _
                            "AND tm_artikel.ast_id IN ( SELECT tm_stueckliste.ast_id FROM (DBA.tm_stueli_posvers INNER JOIN DBA.tm_stueckliste_pos ON tm_stueli_posvers.stlp_id = tm_stueckliste_pos.stlp_id " & _
                            "INNER JOIN DBA.tm_stueckliste ON tm_stueckliste.stl_id = tm_stueli_posvers.stl_id AND tm_stueckliste.stl_version = tm_stueli_posvers.stl_version) " & _
                            "LEFT OUTER JOIN DBA.tm_stueckliste_pos_memo ON tm_stueckliste_pos.stlp_id = tm_stueckliste_pos_memo.stlp_id  WHERE stlp_gueltig_ab <= {d '###Datum###'} AND stlp_gueltig_bis >= {d '###Datum###'} " & _
                            "AND tm_stueckliste_pos_memo.stlp_memo LIKE '%###BMK###%' ) AND tx_artikelstatus.artstat_archiv = 0 AND tm_artikel.ast_id <> 7688"


Const cs_sql_ListBG_BMK = "SELECT tm_artikel.*, tm_artikel_var.*, ean_code, ean_type, tm_artikel.ast_bezeichnung AS artikel_bezeichnung, tm_artikelvar_kd.avkd_artnr_kd, tm_artikel.ast_nummer AS ast_nummer_ueb, tx_artikelstatus.artstat_archiv, " & _
                            "tm_artikel.ast_id, tm_artikel_var.wanr_key, tm_artikel_var.avar_zolltarifnr, tm_artikel_var.avar_nummer AS avar_nummer_ueb, " & _
                            "tm_artikelvar_kd.adk_id AS adkid, tm_kunde.adk_nummer, tm_artvar_lief.avli_artnr_liefer, tm_sml_daten.sml_id, pa_org_name1, " & _
                            "xm_artikel.ast_lqadapter_1q, xm_artikel.ast_lqadapter_2q, xm_artikel.ast_lqadapter_1z, xm_artikel.ast_lqadapter_2z, xm_artikel.ast_ul_kennung, " & _
                            "tm_stueckliste.stuelistat_id, tm_stueckliste.stl_version, xm_artikel.fa_bemerkung, xm_artikel_var.avar_c14_import_kennung " & _
                            "FROM (((DBA.tm_artikel INNER JOIN DBA.tm_artikel_var ON tm_artikel.ast_id = tm_artikel_var.ast_id) " & _
                            "LEFT OUTER JOIN DBA.tm_artikel_var_ean ON tm_artikel_var.ast_id = tm_artikel_var_ean.ast_id AND tm_artikel_var.avar_id = tm_artikel_var_ean.avar_id) " & _
                            "LEFT OUTER JOIN DBA.tm_artikelvar_kd ON tm_artikelvar_kd.ast_id = tm_artikel.ast_id AND tm_artikelvar_kd.avar_id = tm_artikel_var.avar_id AND " & _
                            "tm_artikelvar_kd.adk_id = 73 LEFT OUTER JOIN DBA.tm_partei ON tm_partei.pa_id = tm_artikelvar_kd.adk_id " & _
                            "LEFT OUTER JOIN DBA.tm_kunde ON tm_artikelvar_kd.adk_id = tm_kunde.adk_id AND tm_kunde.ma_id = 2 " & _
                            "LEFT OUTER JOIN DBA.tm_sml_daten ON tm_artikel.ast_id = tm_sml_daten.ast_id AND tm_sml_daten.sprach_id = 1) " & _
                            "INNER JOIN DBA.tx_artikelstatus ON tm_artikel.artstat_id = tx_artikelstatus.artstat_id " & _
                            "LEFT OUTER JOIN DBA.tm_artvar_lief ON tm_artikel_var.ast_id = tm_artvar_lief.ast_id AND tm_artikel_var.avar_id = tm_artvar_lief.avar_id AND " & _
                            "tm_artvar_lief.avli_prioritaet = 1 LEFT OUTER JOIN DBA.xm_artikel ON xm_artikel.ast_id = tm_artikel.ast_id " & _
                            "LEFT OUTER JOIN DBA.xm_artikel_var ON tm_artikel_var.ast_id = xm_artikel_var.ast_id AND tm_artikel_var.avar_id = xm_artikel_var.avar_id " & _
                            "LEFT OUTER JOIN DBA.tm_stueckliste ON tm_artikel.ast_id = tm_stueckliste.ast_id " & _
                            "WHERE tm_artikel.ast_id IN ( SELECT tm_stueckliste.ast_id FROM (DBA.tm_stueli_posvers  " & _
                            "INNER JOIN DBA.tm_stueckliste_pos ON tm_stueli_posvers.stlp_id = tm_stueckliste_pos.stlp_id " & _
                            "INNER JOIN DBA.tm_stueckliste ON tm_stueckliste.stl_id = tm_stueli_posvers.stl_id AND " & _
                            "tm_stueckliste.stl_version = tm_stueli_posvers.stl_version) " & _
                            "LEFT OUTER JOIN DBA.tm_stueckliste_pos_memo ON tm_stueckliste_pos.stlp_id = tm_stueckliste_pos_memo.stlp_id " & _
                            "WHERE stlp_gueltig_ab <= {d '###Datum###'} AND stlp_gueltig_bis >= {d '###Datum###'} AND  " & _
                            "tm_stueckliste_pos_memo.stlp_memo LIKE '%###BMK###%' ) AND tx_artikelstatus.artstat_archiv = 0 AND " & _
                            "tm_artikel.ast_id <> 7688 ORDER BY tm_artikel.ast_nummer ASC, tm_artikel_var.avar_nummer ASC"

Const cs_sql_STLID = "SELECT tm_stueckliste.*, tm_artikel.ast_nummer, tm_artikel.ast_bezeichnung,  ast_matchcode, tm_artikel.ast_zeichnungsnr, " & _
                     "tm_artikel.ast_zeichnr_index " & _
                     "FROM DBA.tm_stueckliste INNER JOIN DBA.tm_artikel ON tm_stueckliste.ast_id = tm_artikel.ast_id " & _
                     "INNER JOIN DBA.tm_artikel_zuord ON tm_artikel_zuord.ast_id = tm_artikel.ast_id  " & _
                     "WHERE tm_stueckliste.ast_id = ###ASTID### ORDER BY ast_nummer"

Const cs_sql_STLBMK = "SELECT tm_stueli_posvers.*, tm_stueckliste.ast_id, tm_stueckliste_pos.*, ta_mengeneinheit.*,   " & _
                  "tm_artikel_var.avar_nummer AS varianten_nr, tm_artikel_var.avar_bezeichnung AS varianten_bez, artvar2.avar_nummer AS kopfart_var, " & _
                  "tm_stueckliste_pos.stlp_avar_id, art2.ast_nummer AS kopfart_nummer, art2.ast_bezeichnung AS kopfart_bez, tm_artikel.ast_nummer, " & _
                  "tm_artikel.ast_bezeichnung, arpl_nummer, arpl_bezeichnung1, tm_artikel.*, tm_artikel_var.avar_hersteller, " & _
                  "tm_artikel_var.avar_herst_nummer, '1234567890123456789012345678901234567890' AS pos_nummer, 'A234567890123456789012345678901234567890' AS pos_bezeichnung, " & _
                  "tm_artikel.artstat_id, tm_artikel.ast_nummer_alt, tm_artikel.artart_id, tm_artikel_ek.aste_lieferzeit, tm_artikel.ast_flag_4, tm_artikelvar_kd.avkd_artnr_kd, " & _
                  "tm_artikel_ek.aste_lieferzeit_hl, xm_stueckliste_pos.bmk_ok_pos_av, um_qz_parse_hist.cmuser_id , parse_hist_date,  tm_artikel.ast_stueli_kz AS isSTL, " & _
                  "xm_stueckliste_pos.stlp_ueberstand, xm_stueckliste_pos.stlp_ueberstand_str, tm_stueckliste_pos_memo.stlp_memo FROM " & _
                  "(((((((DBA.tm_stueli_posvers INNER JOIN DBA.tm_stueckliste_pos ON tm_stueli_posvers.stlp_id = tm_stueckliste_pos.stlp_id)     " & _
                  "INNER JOIN DBA.tm_stueckliste ON tm_stueckliste.stl_id = tm_stueli_posvers.stl_id AND tm_stueckliste.stl_version = tm_stueli_posvers.stl_version) " & _
                  "LEFT OUTER JOIN DBA.ta_mengeneinheit ON tm_stueckliste_pos.me_id = ta_mengeneinheit.me_id) LEFT OUTER JOIN DBA.tm_artikel ON " & _
                  "tm_stueckliste_pos.stlp_member_id = tm_artikel.ast_id AND  tm_stueckliste_pos.stlp_typ = 1) LEFT OUTER JOIN DBA.tm_artikel_var ON " & _
                  "tm_stueckliste_pos.stlp_member_id = tm_artikel_var.ast_id AND tm_stueckliste_pos.stlp_member_sub_id = tm_artikel_var.avar_id AND  " & _
                  "tm_stueckliste_pos.stlp_typ = 1)     LEFT OUTER JOIN DBA.tm_artikelvar_kd ON tm_artikel_var.ast_id = tm_artikelvar_kd.ast_id AND " & _
                  "tm_artikel_var.avar_id = tm_artikelvar_kd.avar_id AND tm_artikelvar_kd.adk_id = 73 AND tm_artikelvar_kd.avkd_default = 1  " & _
                  "LEFT OUTER JOIN DBA.tm_arbeitsplatz ON tm_stueckliste_pos.stlp_member_id = tm_arbeitsplatz.arpl_id AND tm_stueckliste_pos.stlp_typ = 3)  " & _
                  "LEFT OUTER JOIN DBA.tm_artikel_var AS artvar2 ON tm_stueckliste.ast_id = artvar2.ast_id AND tm_stueckliste_pos.stlp_avar_id = artvar2.avar_id) " & _
                  "LEFT OUTER JOIN DBA.tm_artikel AS art2 ON tm_stueckliste.ast_id = art2.ast_id LEFT OUTER JOIN DBA.tm_artikel_ek ON " & _
                  "tm_stueckliste_pos.stlp_member_id = tm_artikel_ek.ast_id AND tm_artikel_ek.avar_id = tm_stueckliste_pos.stlp_member_sub_id " & _
                  "LEFT OUTER JOIN DBA.tm_stueckliste_pos_memo ON tm_stueckliste_pos_memo.stlp_id = tm_stueckliste_pos.stlp_id  " & _
                  "LEFT OUTER JOIN DBA.xm_stueckliste_pos ON xm_stueckliste_pos.stlp_id = tm_stueckliste_pos.stlp_id " & _
                  "LEFT OUTER JOIN DBA.um_qz_parse_hist ON um_qz_parse_hist.stlp_id = tm_stueckliste_pos.stlp_id  " & _
                  "WHERE tm_stueckliste_pos.stlp_gueltig_bis >= '###DATUM###' AND tm_stueckliste_pos.stlp_gueltig_ab <= '###DATUM###' AND " & _
                  "tm_stueli_posvers.stl_id = ###STLID### AND tm_stueli_posvers.stl_version = '###STLVER###' AND stlp_memo LIKE '%###BMK###%' " & _
                  "ORDER BY stlp_nummer, stlp_alternative, kopfart_var, stlp_gueltig_ab"

Const cs_sql_STLArtikel = "SELECT tm_stueli_posvers.*, tm_stueckliste.ast_id, tm_stueckliste_pos.*, ta_mengeneinheit.*,   " & _
                  "tm_artikel_var.avar_nummer AS varianten_nr, tm_artikel_var.avar_bezeichnung AS varianten_bez, artvar2.avar_nummer AS kopfart_var, " & _
                  "tm_stueckliste_pos.stlp_avar_id, art2.ast_nummer AS kopfart_nummer, art2.ast_bezeichnung AS kopfart_bez, tm_artikel.ast_nummer, " & _
                  "tm_artikel.ast_bezeichnung, arpl_nummer, arpl_bezeichnung1, tm_artikel.*, tm_artikel_var.avar_hersteller, " & _
                  "tm_artikel_var.avar_herst_nummer, '1234567890123456789012345678901234567890' AS pos_nummer, 'A234567890123456789012345678901234567890' AS pos_bezeichnung, " & _
                  "tm_artikel.artstat_id, tm_artikel.ast_nummer_alt, tm_artikel.artart_id, tm_artikel_ek.aste_lieferzeit, tm_artikel.ast_flag_4, tm_artikelvar_kd.avkd_artnr_kd, " & _
                  "tm_artikel_ek.aste_lieferzeit_hl, xm_stueckliste_pos.bmk_ok_pos_av, um_qz_parse_hist.cmuser_id , parse_hist_date,  tm_artikel.ast_stueli_kz AS isSTL, " & _
                  "xm_stueckliste_pos.stlp_ueberstand, xm_stueckliste_pos.stlp_ueberstand_str, tm_stueckliste_pos_memo.stlp_memo FROM " & _
                  "(((((((DBA.tm_stueli_posvers INNER JOIN DBA.tm_stueckliste_pos ON tm_stueli_posvers.stlp_id = tm_stueckliste_pos.stlp_id)     " & _
                  "INNER JOIN DBA.tm_stueckliste ON tm_stueckliste.stl_id = tm_stueli_posvers.stl_id AND tm_stueckliste.stl_version = tm_stueli_posvers.stl_version) " & _
                  "LEFT OUTER JOIN DBA.ta_mengeneinheit ON tm_stueckliste_pos.me_id = ta_mengeneinheit.me_id) LEFT OUTER JOIN DBA.tm_artikel ON " & _
                  "tm_stueckliste_pos.stlp_member_id = tm_artikel.ast_id AND  tm_stueckliste_pos.stlp_typ = 1) LEFT OUTER JOIN DBA.tm_artikel_var ON " & _
                  "tm_stueckliste_pos.stlp_member_id = tm_artikel_var.ast_id AND tm_stueckliste_pos.stlp_member_sub_id = tm_artikel_var.avar_id AND  " & _
                  "tm_stueckliste_pos.stlp_typ = 1)     LEFT OUTER JOIN DBA.tm_artikelvar_kd ON tm_artikel_var.ast_id = tm_artikelvar_kd.ast_id AND " & _
                  "tm_artikel_var.avar_id = tm_artikelvar_kd.avar_id AND tm_artikelvar_kd.adk_id = 73 AND tm_artikelvar_kd.avkd_default = 1  " & _
                  "LEFT OUTER JOIN DBA.tm_arbeitsplatz ON tm_stueckliste_pos.stlp_member_id = tm_arbeitsplatz.arpl_id AND tm_stueckliste_pos.stlp_typ = 3)  " & _
                  "LEFT OUTER JOIN DBA.tm_artikel_var AS artvar2 ON tm_stueckliste.ast_id = artvar2.ast_id AND tm_stueckliste_pos.stlp_avar_id = artvar2.avar_id) " & _
                  "LEFT OUTER JOIN DBA.tm_artikel AS art2 ON tm_stueckliste.ast_id = art2.ast_id LEFT OUTER JOIN DBA.tm_artikel_ek ON " & _
                  "tm_stueckliste_pos.stlp_member_id = tm_artikel_ek.ast_id AND tm_artikel_ek.avar_id = tm_stueckliste_pos.stlp_member_sub_id " & _
                  "LEFT OUTER JOIN DBA.tm_stueckliste_pos_memo ON tm_stueckliste_pos_memo.stlp_id = tm_stueckliste_pos.stlp_id  " & _
                  "LEFT OUTER JOIN DBA.xm_stueckliste_pos ON xm_stueckliste_pos.stlp_id = tm_stueckliste_pos.stlp_id " & _
                  "LEFT OUTER JOIN DBA.um_qz_parse_hist ON um_qz_parse_hist.stlp_id = tm_stueckliste_pos.stlp_id  " & _
                  "WHERE tm_stueckliste_pos.stlp_gueltig_bis >= '###DATUM###' AND tm_stueckliste_pos.stlp_gueltig_ab <= '###DATUM###' AND " & _
                  "tm_stueli_posvers.stl_id = ###STLID### AND tm_stueli_posvers.stl_version = '###STLVER###' " & _
                  "ORDER BY stlp_nummer, stlp_alternative, kopfart_var, stlp_gueltig_ab"

Const cs_AddString = "Insert into tblTempBMK_BG_Pos_Uebersicht (BG_ArtNr,BG_ArtNr_AstID,BG_ArtNr_Bez,BG_ArtNr_Bez2,BG_ArtNr_M1,BG_ArtNr_M2,BG_ArtNr_M3,BG_ArtNr_M4," & _
                     "BG_ArtNr_M5,BG_ArtNr_M6,BMK_ArtNr,BMK_ArtNr_var,BMK_stlp_pos,BMK_stlp_id,BMK_stlp_menge," & _
                     "BMK_ArtNr_Bez,BMK_ArtNr_Memo,BMK_ArtNr_Ueberstand,BMK_ArtNr_Bez2,BMK_ArtNr_M1,BMK_ArtNr_M2,BMK_ArtNr_M3,BMK_ArtNr_M4,BMK_ArtNr_M5,BMK_ArtNr_M6) " & _
                     "values (###VALUES###);"


Const cs_ArtikelDataArtikelNr = "SELECT tm_artikel.*, tm_artikel_var.*, ean_code, ean_type, tm_artikel.ast_bezeichnung AS artikel_bezeichnung, tm_artikelvar_kd.avkd_artnr_kd, " & _
                                "tm_artikel.ast_nummer AS ast_nummer_ueb, tx_artikelstatus.artstat_archiv, " & _
                                "tm_artikel.ast_id, tm_artikel_var.wanr_key, tm_artikel_var.avar_zolltarifnr, " & _
                                "tm_artikel_var.avar_nummer AS avar_nummer_ueb, tm_artikelvar_kd.adk_id AS adkid, tm_kunde.adk_nummer, tm_artvar_lief.avli_artnr_liefer, " & _
                                "tm_sml_daten.sml_id, pa_org_name1, xm_artikel.ast_lqadapter_1q, xm_artikel.ast_lqadapter_2q, xm_artikel.ast_lqadapter_1z, " & _
                                "xm_artikel.ast_lqadapter_2z, xm_artikel.ast_ul_kennung, tm_stueckliste.stuelistat_id, tm_stueckliste.stl_version, xm_artikel.fa_bemerkung, " & _
                                "xm_artikel_var.avar_c14_import_kennung " & _
                                "FROM (((DBA.tm_artikel INNER JOIN DBA.tm_artikel_var ON tm_artikel.ast_id = tm_artikel_var.ast_id) " & _
                                "LEFT OUTER JOIN DBA.tm_artikel_var_ean ON tm_artikel_var.ast_id = tm_artikel_var_ean.ast_id AND " & _
                                "tm_artikel_var.avar_id = tm_artikel_var_ean.avar_id) LEFT OUTER JOIN DBA.tm_artikelvar_kd ON tm_artikelvar_kd.ast_id = tm_artikel.ast_id AND " & _
                                "tm_artikelvar_kd.avar_id = tm_artikel_var.avar_id AND tm_artikelvar_kd.adk_id = 73 " & _
                                "LEFT OUTER JOIN DBA.tm_partei ON tm_partei.pa_id = tm_artikelvar_kd.adk_id " & _
                                "LEFT OUTER JOIN DBA.tm_kunde ON tm_artikelvar_kd.adk_id = tm_kunde.adk_id AND tm_kunde.ma_id = 2 " & _
                                "LEFT OUTER JOIN DBA.tm_sml_daten ON tm_artikel.ast_id = tm_sml_daten.ast_id AND tm_sml_daten.sprach_id = 1) " & _
                                "INNER JOIN DBA.tx_artikelstatus ON tm_artikel.artstat_id = tx_artikelstatus.artstat_id " & _
                                "LEFT OUTER JOIN DBA.tm_artvar_lief ON tm_artikel_var.ast_id = tm_artvar_lief.ast_id AND tm_artikel_var.avar_id = tm_artvar_lief.avar_id AND " & _
                                "tm_artvar_lief.avli_prioritaet = 1 LEFT OUTER JOIN DBA.xm_artikel ON xm_artikel.ast_id = tm_artikel.ast_id " & _
                                "LEFT OUTER JOIN DBA.xm_artikel_var ON tm_artikel_var.ast_id = xm_artikel_var.ast_id AND tm_artikel_var.avar_id = xm_artikel_var.avar_id " & _
                                "LEFT OUTER JOIN DBA.tm_stueckliste ON tm_artikel.ast_id = tm_stueckliste.ast_id  " & _
                                "WHERE tx_artikelstatus.artstat_archiv = 0 AND tm_artikel.ast_nummer = '###ARTNR###' AND tm_artikel_var.avar_nummer = '###VAR###' AND " & _
                                "tm_artikel.ast_id <> 7688 " & _
                                "ORDER BY tm_artikel.ast_nummer ASC, tm_artikel_var.avar_nummer ASC"
Dim cnn As ADODB.Connection

Private Sub class_Initialize()

    Set cnn = New ADODB.Connection

'    cnn.ConnectionString = "DSN=avista;UID=dba;PWD=phoenix;"
    cnn.ConnectionString = getAvistaProvider
    cnn.Open

End Sub

Private Sub class_Terminate()

    If Not cnn Is Nothing Then
        cnn.Close
        Set cnn = Nothing
    End If

End Sub

Public Function GetAvistaConnection() As ADODB.Connection

    If cnn Is Nothing Then
        Set cnn = New ADODB.Connection
'        cnn.ConnectionString = "DSN=avista;UID=dba;PWD=phoenix;"
        cnn.ConnectionString = getAvistaProvider
        cnn.Open
    End If
    Set GetAvistaConnection = cnn

End Function


Public Function closeAvistaConnection() As Boolean

    If Not cnn Is Nothing Then
        closeAvistaConnection = False
        cnn.Close
        Set cnn = Nothing
        closeAvistaConnection = True
    End If

End Function

Public Function getRecordsetAvista(f_sql As String) As ADODB.Recordset
On Error GoTo err_SetBG

    Dim rs As ADODB.Recordset
    If cnn Is Nothing Then
        cnn = GetAvistaConnection()
    End If
    Set rs = New ADODB.Recordset
    With rs
        Set .ActiveConnection = cnn
        '.Source = AvistaArtikelSQL_Select & NewData & "' " & AvistaArtikelSQL_Order
        .source = f_sql
        .LockType = adLockOptimistic
        .CursorType = adOpenKeyset
        .Open
    End With        'Set the form's Recordset property to the ADO recordset

    'Set rs = cnn.Execute(f_sql)
    Set getRecordsetAvista = rs
    Set rs = Nothing

close_db:
   Exit Function

err_SetBG:
    MsgBox "error: " & Err & " - " & Error
    Resume Next

End Function
Public Function CommandExecuteAvista(f_sql As String) As ADODB.Recordset
On Error GoTo err_SetBG

    Dim rs As ADODB.Recordset
    Dim com As ADODB.Command
'    Dim x As ADODB.Command
    If cnn Is Nothing Then
        cnn = GetAvistaConnection()
    End If

    Set com = New ADODB.Command
    com.ActiveConnection = cnn

    com.CommandText = f_sql
    com.Execute

    Set com = Nothing

close_db:
   Exit Function

err_SetBG:
    If Err = -2147217900 Then
        'MsgBox "error: " & Err & " - " & Error
        Resume Next
    End If

End Function

Public Function getAvistaProvider() '20180316-ROS

    Dim AvistaProvider As String

    Select Case getCurrLanguage
        Case "German"
            AvistaProvider = "DSN=avista;UID=dba;PWD=phoenix;"
        Case "English"
            AvistaProvider = "DSN=avista_us;UID=dba;PWD=phoenix;"
        Case "Chinese"
            AvistaProvider = "DSN=avista_cn;UID=dba;PWD=phoenix;"
    End Select

    getAvistaProvider = AvistaProvider

End Function

Public Function genTempTableBGUbersichtBMK(fs_BMK As String, fs_Table As String, fs_MachNr As String) As String
On Error GoTo err_genTempTableBGUbersichtBMK

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Dim ls_currArtikelBMK As String
    Dim rsAvista As ADODB.Recordset
    Dim rsAvistaPos As ADODB.Recordset
    Dim ls_sql As String
    Dim ls_sql_STLID As String
    Dim ls_sql_STL As String
    Dim ListArray(2, 1, 1) As String
    Dim ls_Values As String
    Dim ls_ValuesBG As String
    Dim ls_ValuesBGUe As String
    Dim fld As Field
    Dim lr_FktRange As Range
    Dim lr_BMKRange As Range
    Dim lr_BMK As Range
    Dim ls_BMKAddress As String
    Dim i As Integer
    Dim lDatum As String

    genTempTableBGUbersichtBMK = ""

    lDatum = format(Date, "yyyy-mm-dd")

    '###################################################################################
    'Ermitteln der Baugruppen in Avista anhand des angezeigten BMK
    '###################################################################################
    If fs_BMK <> "" Then
        Set ldb = openAccDBConnection("openDB")
        'Temp tabelle leeren
        ldb.Execute ("delete * from " & fs_Table)
        Set rs = ldb.OpenRecordset("Select * from " & fs_Table)
        '###################################################################################
        'Ermitteln aller Baugruppen die im Memofeld das defnierte BMK eingetragen haben
        '###################################################################################
        If fs_MachNr = "" Then
            ls_sql = Replace(cs_sql_ListBG_BMK, "###BMK###", fs_BMK)
        Else
            ls_sql = Replace(cs_sql_ListBG_BMK_KN, "###BMK###", fs_BMK)
            ls_sql = Replace(ls_sql, "###MaschNr###", fs_MachNr)
        End If

        ls_sql = Replace(ls_sql, "###Datum###", lDatum)

        Set rsAvista = getRecordsetAvista(ls_sql)
        If rsAvista.RecordCount Then
'            Debug.Print "Anzahl gefundener Baugruppen: " & rsAvista.RecordCount
            i = 0
            '###################################################################################
            'Über alle ermittelten Baugruppen
            '###################################################################################
            While Not rsAvista.EOF
                'Baugruppenwerte für Artikel sichern
                ls_ValuesBG = "'" & rsAvista("ast_nummer") & "', "
                ls_ValuesBG = ls_ValuesBG & "" & rsAvista("ast_id") & ", "
                ls_ValuesBG = ls_ValuesBG & "'" & rsAvista("ast_bezeichnung") & "', "
                'Baugruppenwerte für BG-Übersicht sichern
                ls_ValuesBG = ls_ValuesBG & "'" & rsAvista("ast_bezeichnung2") & "', "
                ls_ValuesBG = ls_ValuesBG & "'" & rsAvista("ast_merkmal_1") & "', "
                ls_ValuesBG = ls_ValuesBG & "'" & rsAvista("ast_merkmal_2") & "', "
                ls_ValuesBG = ls_ValuesBG & "'" & rsAvista("ast_merkmal_3") & "', "
                ls_ValuesBG = ls_ValuesBG & "'" & rsAvista("ast_merkmal_4") & "', "
                ls_ValuesBG = ls_ValuesBG & "'" & rsAvista("ast_merkmal_5") & "', "
                ls_ValuesBG = ls_ValuesBG & "'" & rsAvista("ast_merkmal_6") & "', "

'                    Me.lstBGUebersicht.AddItem
'                    Me.lstBGUebersicht.List(i, 0) = rsAvista("ast_nummer")
'                    Me.lstBGUebersicht.List(i, 2) = rsAvista("ast_bezeichnung")
'                    Me.lstBGUebersicht.List(i, 1) = rsAvista("ast_id")

                'Position des BMK in der aktuellen Baugruppe ermitteln und ausgeben
                ls_sql_STLID = Replace(cs_sql_STLID, "###ASTID###", rsAvista("ast_id"))
                Set rsAvistaPos = getRecordsetAvista(ls_sql_STLID)
                If rsAvistaPos.RecordCount Then
                    ls_sql_STL = Replace(cs_sql_STLBMK, "###DATUM###", format(Date, "yyyy-mm-dd"))
                    ls_sql_STL = Replace(ls_sql_STL, "###STLID###", rsAvistaPos("stl_id"))
                    ls_sql_STL = Replace(ls_sql_STL, "###STLVER###", rsAvistaPos("stl_version"))
                    ls_sql_STL = Replace(ls_sql_STL, "###BMK###", fs_BMK)
                    Set rsAvistaPos = getRecordsetAvista(ls_sql_STL)
                    If rsAvistaPos.RecordCount Then
                        i = 0
                        'Me.lstSTLBG.Clear
'                        Debug.Print "Anzahl gefundener Positionen in Baugruppe: " & rsAvistaPos.RecordCount
                        While Not rsAvistaPos.EOF
                            ls_Values = ls_ValuesBG & "'" & rsAvistaPos("ast_nummer") & "', "
                            ls_Values = ls_Values & "'" & rsAvistaPos("varianten_nr") & "', "
                            ls_Values = ls_Values & "" & rsAvistaPos("stlp_nummer") & ", "
                            ls_Values = ls_Values & "" & rsAvistaPos("stlp_id") & ", "
                            'Wenn Menge Komma beinhaltet, diesen in einen Punkt umwandeln
                            If InStr(rsAvistaPos("stlp_menge"), ",") Then
                                ls_Values = ls_Values & "" & Replace(rsAvistaPos("stlp_menge"), ",", ".") & ", "
                            Else
                                ls_Values = ls_Values & "" & rsAvistaPos("stlp_menge") & ", "
                            End If
                            ls_Values = ls_Values & "'" & rsAvistaPos("ast_bezeichnung") & "', "
                            ls_Values = ls_Values & "'" & rsAvistaPos("stlp_memo") & "', "
                            If rsAvistaPos("stlp_ueberstand_str") = "" Or IsNull(rsAvistaPos("stlp_ueberstand_str")) Then
                                ls_Values = ls_Values & "Null, "
                            Else
                                ls_Values = ls_Values & "" & rsAvistaPos("stlp_ueberstand_str") & ", "
                            End If
                            ls_Values = ls_Values & "'" & rsAvistaPos("ast_bezeichnung2") & "', "
                            ls_Values = ls_Values & "'" & rsAvistaPos("ast_merkmal_1") & "', "
                            ls_Values = ls_Values & "'" & rsAvistaPos("ast_merkmal_2") & "', "
                            ls_Values = ls_Values & "'" & rsAvistaPos("ast_merkmal_3") & "', "
                            ls_Values = ls_Values & "'" & rsAvistaPos("ast_merkmal_4") & "', "
                            ls_Values = ls_Values & "'" & rsAvistaPos("ast_merkmal_5") & "', "
                            ls_Values = ls_Values & "'" & rsAvistaPos("ast_merkmal_6") & "'"

                            ldb.Execute (Replace(cs_AddString, "###VALUES###", ls_Values))
                            rsAvistaPos.MoveNext
                            i = i + 1
                        Wend
                        'lstSTLBG.ListIndex = 0

                    Else
                        MsgBox "Für STL_ID " & rsAvista("stl_id") & " gibt es keine Stücklisteneintraege!!"
                    End If


                Else
                    MsgBox "Für Artikel gibt es keine Stückliste!!"
                End If

                rsAvista.MoveNext
                i = i + 1
            Wend
'            Me.lstBGUebersicht.List = ListArray()
            'End With

        Else
            MsgBox "Für das BMK '" & ls_currArtikelBMK & "' wurde keine Baugruppe gefunden!"
        End If
    Else
        MsgBox "Kein Artikel ausgewählt!"
    End If

    Set rsAvista = Nothing
    Set rsAvistaPos = Nothing
    Set rs = Nothing
    Set db = Nothing
    Exit Function

err_genTempTableBGUbersichtBMK:

    If Err = 3075 Or Err = 3265 Then
        Resume Next
    End If

End Function
Public Function getLQMatNrBezeichnung(fsMatNr As String, Optional fsMatNrVar As String = "00000") As ADODB.Recordset

    Dim rsAvista As ADODB.Recordset
    Dim rsAvistaPos As ADODB.Recordset
    Dim ls_sql As String

    '###################################################################################
    'Ermitteln der Artikelbezeichnung in Avista anhand der Materialnummer oder
    'Ermitteln der AVISTA_Artikel_ID in Avista anhand der Materialnummer oder
    'Ermitteln der AVISTA_VAR_ID in Avista anhand der Materialnummer
    '###################################################################################
    If fsMatNr <> "" Then
        ls_sql = Replace(cs_ArtikelDataArtikelNr, "###ARTNR###", fsMatNr)
        ls_sql = Replace(ls_sql, "###VAR###", fsMatNrVar)
        Set rsAvista = getRecordsetAvista(ls_sql)
        If rsAvista.RecordCount >= 1 Then
            Set getLQMatNrBezeichnung = rsAvista
        Else
'            MsgBox "Funktion getLQMatNrBezeichnung: Keine Materialnummer übergeben oder kein Datensatz vorhanden!"
        End If
        Set rsAvista = Nothing
    End If

End Function

Public Function getLQMatNrSTL(fsMatNr As String, Optional fsMatNrVar As String = "00000") As ADODB.Recordset

    Dim rsAvista As ADODB.Recordset
    Dim rsAvistaPos As ADODB.Recordset
    Dim ls_sql As String

    '###################################################################################
    'Ermitteln der Artikelbezeichnung in Avista anhand der Materialnummer
    '###################################################################################
    If fsMatNr <> "" Then
    '
        Set rsAvista = getLQMatNrBezeichnung(fsMatNr, fsMatNrVar)
        If rsAvista.RecordCount Then
            Set rsAvista = getRecordsetAvista(Replace(cs_sql_STLID, "###ASTID###", rsAvista("ast_id")))
            If rsAvista.RecordCount Then
                ls_sql = Replace(cs_sql_STLArtikel, "###DATUM###", format(Date, "yyyy-mm-dd"))
                ls_sql = Replace(ls_sql, "###STLID###", rsAvista("stl_id"))
                ls_sql = Replace(ls_sql, "###STLVER###", rsAvista("stl_version"))
                Set rsAvista = getRecordsetAvista(ls_sql)
                If rsAvista.RecordCount >= 1 Then
                    Set getLQMatNrSTL = rsAvista
                Else
                    MsgBox "Funktion getLQMatNrSTL: Keine STL vorhanden!"
                End If
            End If
        End If
        Set rsAvista = Nothing
    End If

End Function
