Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub cboWareneingangs_Lager_AfterUpdate()
'Bestands Ermittung der Service-Lager beim öffnen des Formulars mit Ausgabe des Buchungs-Text

    Dim ldb As DAO.Database
    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
    Dim rsAvista As ADODB.Recordset
    Dim ls_sql As String
    Dim BuchText As String


    ls_sql = "SELECT ast_nummer + '-' + avar_nummer AS ArtNr, tm_artikel.ast_bezeichnung as Bez, tt_bestand_sum.besum_ist_menge as St, tt_bestand_sum.besum_zu_dat_last as Zugang ,tm_lager.lager_bez as Lager, " & _
            "(SELECT TOP(1) mbk_text1 as BuchText " & _
            "From tt_mb_beleg_kopf " & _
            "INNER JOIN tt_mb_beleg_pos ON tt_mb_beleg_kopf.mbk_id = tt_mb_beleg_pos.mbk_id " & _
            "WHERE tt_mb_beleg_pos.ast_id = tm_artikel_var.ast_id And tt_mb_beleg_pos.avar_id = tm_artikel_var.avar_id And tt_mb_beleg_pos.lager_id = tt_bestand_sum.lager_id And tt_mb_beleg_pos.mbp_art = 1 " & _
            "ORDER BY tt_mb_beleg_kopf.mbk_id DESC) " & _
            "From DBA.tt_bestand_sum " & _
            "INNER JOIN DBA.tm_artikel ON tt_bestand_sum.ast_id = tm_artikel.ast_id " & _
            "INNER JOIN DBA.tm_artikel_var ON tt_bestand_sum.ast_id = tm_artikel_var.ast_id  AND tt_bestand_sum.avar_id = tm_artikel_var.avar_id " & _
            "INNER JOIN DBA.tm_lager ON tt_bestand_sum.lager_id =  tm_lager.lager_id " & _
            "WHERE tt_bestand_sum.lager_id = ###Lager### And tt_bestand_sum.besum_ist_menge >= 1 " & _
            "ORDER BY Zugang ASC"

        ls_sql = Replace(ls_sql, "###Lager###", Me.cboWareneingangs_Lager)

        Set ldb = openAccDBConnection("openDB")
        ldb.Execute ("delete * from tblTempLager")
        Set adbprov = New clsAvistaProvider
        Set rsAvista = New ADODB.Recordset
        Set rsAvista = adbprov.getRecordsetAvista(ls_sql)
        If rsAvista.RecordCount Then
            Do While Not rsAvista.EOF
            'Funktion BuchText abrufen nach übergabe Zugangsdatum und Lager ID
                BuchText = Buchungstext(rsAvista("Zugang"), Me.cboWareneingangs_Lager, rsAvista("ArtNr"))
            '##################################################################
                ldb.Execute ("Insert into tblTempLager ([ArtNr], [Bez], [St],[Zugang],[Lager],[BuchText])" & _
                "values('" & rsAvista("ArtNr") & "', '" & rsAvista("Bez") & "', '" & rsAvista("St") & "',  '" & rsAvista("Zugang") & "', '" & rsAvista("Lager") & "', '" & BuchText & "');")
                rsAvista.MoveNext
            Loop
        End If
        Set rsAvista = Nothing
        Set adbprov = Nothing
        Me.Requery

End Sub

Private Sub Form_Close()
    Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")
    'Temp tabelle leeren
    ldb.Execute ("delete * from tblTempLager")

    FormKoorSave Me

End Sub

Private Sub Form_DblClick(Cancel As Integer)

    p_ReklNr = ZahlAusString(Me.BuchText)
    NewReklamationE "Bearbeiten", get_p_ReklNr()
    If IsLoaded("Fehlererfassung - Register") Then
        Forms("Fehlererfassung - Register").SetFocus
    End If

End Sub


'#############################################################
'Funktion in Allgemeine Funktionen abgelegt

'Function ZahlAusString(ByVal strGiven As String) As String
'    Dim intLength As Integer
'    Dim lngValue As Long
'
'        For intLength = 1 To Len(strGiven)
'          If IsNumeric(MID(strGiven, intLength, 1)) Then
'            lngValue = lngValue & MID(strGiven, intLength, 1)
'          End If
'        Next intLength
'        ZahlAusString = lngValue
'
'End Function
'#############################################################
Private Sub Form_Open(Cancel As Integer)

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub
