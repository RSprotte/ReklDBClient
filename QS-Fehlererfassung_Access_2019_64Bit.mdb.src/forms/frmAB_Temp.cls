Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AB_DblClick(Cancel As Integer)

    Daten_AB_Avista Me.AB

End Sub

Private Sub cmbDrucken_Click()

    PrintForm Me

End Sub

Private Sub cmbDrucken_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbExcelExport_Click()
    Dim strFilter As String

    ' Filter nur übergeben, wenn er im Formular auch gerade aktiv angewendet wird
    If Me.FilterOn Then
        strFilter = Me.Filter
    Else
        strFilter = ""
    End If

    ' Aufruf der globalen Prozedur mit Übergabe der Werte
    ' Optional können Sie als dritten Parameter einen eigenen Dateinamen mitgeben
    Call ExportRecordSourceNachExcel(Me.RecordSource, strFilter, Me.Name)

End Sub

Private Sub cmbExcelExport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Form_Close()

    CurrentDb.Execute ("delete * from tblABTemp;")
    FormKoorSave Me

End Sub

Private Sub Form_DblClick(Cancel As Integer)

    Dim lsForm As String

    p_ReklNr = Me.Vorgang
        NewReklamationE "Bearbeiten", get_p_ReklNr()
        lsForm = "Fehlererfassung - Register"

End Sub

Private Sub Form_Load()

    FormOpenKoor Me

End Sub

Private Sub Form_Open(Cancel As Integer)

    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim sql As String
    Dim strVorgang As String

    'Abfrage dient zur Ermittlung der offenen AB`s 7940 = Sachbearbeiter ROS / 8032 = Sachbearbeiter RSP

'    sql = "SELECT tt_abtbkopf.adk_nummer,pa_matchcode, tt_abtbkopf.ab_nummer, tt_auftrag.af_best_ref,tt_auftrag.af_liefertermin_txt,  tt_abtbkopf_notiz.ab_text2 " & _
            "FROM DBA.tt_abtbkopf INNER JOIN DBA.tt_auftrag ON tt_abtbkopf.af_id = tt_auftrag.af_id " & _
            "INNER JOIN tt_abtbkopf_notiz ON tt_abtbkopf.ab_id = tt_abtbkopf_notiz.ab_id " & _
            "INNER JOIN DBA.tm_partei ON tt_auftrag.adk_id = tm_partei.pa_id " & _
            "WHERE tt_abtbkopf.ma_id = 2 AND ((tt_auftrag.af_sachbearb = 7940) or (tt_auftrag.af_sachbearb = 8032)) AND tt_auftrag.af_status = 10 order by af_liefertermin_txt DESC;"

    'Neue Abfrage ob im Auftragstext2 der AB ein Text eingetragen ist, der mit "Vor" anfängt, da wir vom Service immer Vorgang und die Vorgangsnummer eintragen z.B. "Vorgang 0815"
    'somit ist das unabhängig vom Sachbearbeiter

'     sql = "SELECT tt_abtbkopf.adk_nummer,pa_matchcode, tt_abtbkopf.ab_nummer, tt_auftrag.af_best_ref,tt_auftrag.af_liefertermin_txt,  tt_abtbkopf_notiz.ab_text2, tt_abtbpos_termin.ab_tpos_term_best, tt_abtbpos_termin.ab_tpos_term_versand " & _
'            "FROM DBA.tt_abtbkopf INNER JOIN DBA.tt_auftrag ON tt_abtbkopf.af_id = tt_auftrag.af_id " & _
'            "INNER JOIN tt_abtbkopf_notiz ON tt_abtbkopf.ab_id = tt_abtbkopf_notiz.ab_id " & _
'            "INNER JOIN DBA.tm_partei ON tt_auftrag.adk_id = tm_partei.pa_id " & _
'            "INNER JOIN DBA.tt_abtbpos_termin ON tt_abtbpos_termin.ab_id = tt_abtbkopf.ab_id " & _
'            "WHERE tt_abtbkopf.ma_id = 2 AND tt_abtbkopf_notiz.ab_text2 Like 'Vor%' AND tt_auftrag.af_status = 10 AND ab_pos_id = 1 order by tt_abtbpos_termin.ab_tpos_term_versand ASC;"

     sql = "SELECT tt_abtbkopf.adk_nummer,pa_matchcode, tt_abtbkopf.ab_nummer, tt_auftrag.af_best_ref,tt_auftrag.af_liefertermin_txt,  tt_abtbkopf_notiz.ab_text2, tt_abtbpos_termin.ab_tpos_term_best, tt_abtbpos_termin.ab_tpos_term_versand " & _
            "FROM DBA.tt_abtbkopf INNER JOIN DBA.tt_auftrag ON tt_abtbkopf.af_id = tt_auftrag.af_id " & _
            "INNER JOIN tt_abtbkopf_notiz ON tt_abtbkopf.ab_id = tt_abtbkopf_notiz.ab_id " & _
            "INNER JOIN DBA.tm_partei ON tt_auftrag.adk_id = tm_partei.pa_id " & _
            "INNER JOIN DBA.tt_abtbpos_termin ON tt_abtbpos_termin.ab_id = tt_abtbkopf.ab_id " & _
            "WHERE tt_abtbkopf.ma_id = 2 AND tt_abtbkopf_notiz.ab_text2 Like 'Vor%' AND tt_auftrag.af_status = 10 order by tt_abtbpos_termin.ab_tpos_term_versand ASC;"



    Set adbprov = New clsAvistaProvider
    Set rsAvista = New ADODB.Recordset
    Set rsAvista = adbprov.getRecordsetAvista(sql)

    'Daten werden in eine Temp-Tabelle eingetragen und bei öffnen des Formular frmAB_Temp wiedergegeben

    If rsAvista.RecordCount Then
        CurrentDb.Execute ("delete * from tblABTemp;")
        rsAvista.MoveFirst
        While Not rsAvista.EOF
'            strVorgang = Right(rsAvista("ab_text2"), Len(rsAvista("ab_text2")) - InStrRev(rsAvista("ab_text2"), " "))
            strVorgang = ZahlAusString(rsAvista("ab_text2"))
            CurrentDb.Execute ("Insert into tblABTemp ([KD_Nr],[KD],[AB],[Bestell_Nr],[LT],[Vorgang],[VT]) values ('" & rsAvista("adk_nummer") & "', '" & rsAvista("pa_matchcode") & "', '" & rsAvista("ab_nummer") & "', '" & rsAvista("af_best_ref") & "','" & rsAvista("ab_tpos_term_best") & "','" & strVorgang & "','" & rsAvista("ab_tpos_term_versand") & "');")
            rsAvista.MoveNext
        Wend
        Me.Requery
    Else
'        MsgBox ("Keine offene AB`s")
        MsgBox getTC_FCFromCode("TC4236"), vbInformation, getTC_FCFromCode("TC4237")
        DoCmd.Close acForm, "frmAB_Temp"
    End If

    Set rsAvista = Nothing
    Set adbprov = Nothing

End Sub
