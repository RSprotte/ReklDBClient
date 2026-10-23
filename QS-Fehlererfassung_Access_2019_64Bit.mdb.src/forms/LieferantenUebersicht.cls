Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Befehl10_Click()

End Sub

Private Sub Befehl7_Click()

    FormKoorSave Me
    DoCmd.Close

End Sub

Private Sub Befehl8_Click()
Dim ldb As DAO.Database

 Set ldb = openAccDBConnection("openDB")

If MsgBox(getTC_FCFromCode("TC0472"), vbOKCancel) = vbOK Then
'If MsgBox("Lieferanten wirklich löschen?", vbOKCancel) = vbOK Then
    ldb.Execute ("delete * from Lieferanten where [Lieferant-ID] = " & Me.[Lieferant-ID] & ";")
    Me.Requery
    Me.Refresh
End If

 ' Set ldb = openAccDBConnection("closeDB")

End Sub

Private Sub Befehl9_Click()

'    Dim adbprov As clsAvistaProvider
'    Dim rsAvista As ADODB.Recordset
'    Dim dbAccess As DAO.Database
'    Dim rsAccess As DAO.Recordset
'    Dim strsql  As String
'    Dim SQL As String
'    Dim blnGesperrt As Boolean
'    Dim blnLQLief As Boolean
'    Dim intNeuLiefID As Long
'
'    strsql = "SELECT tm_lieferant.adl_nummer, tm_lieferant.adl_info1, tm_lieferant.adl_info4, ta_adress_gruppe.adgrp_kuerzel, tm_partei.pa_matchcode, tm_partei.pa_org_name1, " & _
'                "tm_adress_basis.adb_strasse , tm_adress_basis.adb_plz_str, tm_adress_basis.adb_ort, tm_adress_basis.adb_tel1, tm_adress_basis.adb_fax, tm_adress_basis.adb_email " & _
'                "From DBA.tm_partei " & _
'                "INNER JOIN DBA.ta_adress_gruppe ON tm_partei.adgrp_id = ta_adress_gruppe.adgrp_id " & _
'                "INNER JOIN DBA.tm_lieferant ON tm_partei.pa_id =  tm_lieferant.adl_id AND tm_lieferant.ma_id = 2 " & _
'                "INNER JOIN DBA.tm_adress_basis ON tm_partei.pa_id = tm_adress_basis.pa_id " & _
'                "WHERE tm_lieferant.ma_id = 2  AND tm_adress_basis.adt_id = 1 AND tm_partei.adgrp_id <> 1 AND tm_lieferant.adl_nummer = 'XXXLiefNrXXX' ORDER BY pa_matchcode, adl_nummer;"
'
'    blnGesperrt = False
'    blnLQLief = False
'    Set dbAccess = openAccDBConnection("openDB")
'    Set adbprov = New clsAvistaProvider
'    SQL = Replace(strsql, "XXXLiefNrXXX", Me.L_Nummer_AVISTA)
'    Set rsAvista = adbprov.getRecordsetAvista(SQL)
'
'    If rsAvista.RecordCount Then
'        If rsAvista("adgrp_kuerzel") = "Gesperrt" Then blnGesperrt = True
'        If rsAvista("adl_info4") = "LQ" Then blnLQLief = True
'        dbAccess.Execute ("Update Lieferanten set [Matchcode] = '" & rsAvista("pa_matchcode") & "', [Lieferantenname] = '" & rsAvista("pa_org_name1") & "', " & _
'        "[Telefaxnummer] = '" & rsAvista("adb_fax") & "', [Telefonnummer] = '" & rsAvista("adb_tel1") & "',[E-Mail] = '" & rsAvista("adb_email") & "', " & _
'        "[Adresse]= '" & rsAvista("adb_strasse") & "',[Postleitzahl] = " & rsAvista("adb_plz_str") & ",[Ort] = '" & rsAvista("adb_ort") & "', " & _
'        "[L-Kategorie] = '" & rsAvista("adgrp_kuerzel") & "',[VerantwortlichEinkauf] = '" & rsAvista("adl_info1") & "', " & _
'        "[GESPERRT] = " & CInt(blnGesperrt) & " ,[LieferantenverantwortungLQ] = " & CInt(blnLQLief) & " " & _
'        "where [L-Nummer/AVISTA] = " & Me.L_Nummer_AVISTA & ";")
'    End If

    If Not IsNull(Me.L_Nummer_AVISTA) Then
        Call LiefDatenAbrufen(Me.L_Nummer_AVISTA)
        Me.Filter = "[L-Nummer/AVISTA] = " & Me.L_Nummer_AVISTA
    Else
        Me.Filter = "[Lieferant-ID] = " & Me.[Lieferant-ID]
    End If

    Me.FilterOn = True
    Me.Refresh
    DoCmd.OpenForm "Lieferanten"
    With Forms("Lieferanten")
        .Filter = "[Lieferant-ID] = " & Me.[Lieferant-ID]
        .FilterOn = True '20190827-ROS_1
        .Requery
        .Refresh
    End With

End Sub

Private Sub cmbAktuallisieren_Click()
'    On Error Resume Next
'
'    Dim adbprov As clsAvistaProvider
'    Dim rsAvista As ADODB.Recordset
'    Dim dbAccess As DAO.Database
'    Dim rsAccess As DAO.Recordset
'    Dim strsql  As String
'    Dim SQL As String
'    Dim blnGesperrt As Boolean
'    Dim blnLQLief As Boolean
'
'    strsql = "SELECT tm_lieferant.adl_nummer, tm_lieferant.adl_info1, tm_lieferant.adl_info4, ta_adress_gruppe.adgrp_kuerzel, tm_partei.pa_matchcode, tm_partei.pa_org_name1, " & _
'                "tm_adress_basis.adb_strasse , tm_adress_basis.adb_plz_str, tm_adress_basis.adb_ort, tm_adress_basis.adb_tel1, tm_adress_basis.adb_fax, tm_adress_basis.adb_email " & _
'                "From DBA.tm_partei " & _
'                "INNER JOIN DBA.ta_adress_gruppe ON tm_partei.adgrp_id = ta_adress_gruppe.adgrp_id " & _
'                "INNER JOIN DBA.tm_lieferant ON tm_partei.pa_id =  tm_lieferant.adl_id AND tm_lieferant.ma_id = 2 " & _
'                "INNER JOIN DBA.tm_adress_basis ON tm_partei.pa_id = tm_adress_basis.pa_id " & _
'                "WHERE tm_lieferant.ma_id = 2  AND tm_adress_basis.adt_id = 1 AND tm_partei.adgrp_id <> 1 AND tm_lieferant.adl_nummer = 'XXXLiefNrXXX' ORDER BY pa_matchcode, adl_nummer;"
'
'    Set dbAccess = openAccDBConnection("openDB")
'    Set adbprov = New clsAvistaProvider
'    Set rsAccess = dbAccess.OpenRecordset("Select [L-Nummer/AVISTA] from Lieferanten;")
'    If rsAccess.RecordCount Then
'        rsAccess.MoveFirst
'        While Not rsAccess.EOF
'            SQL = Replace(strsql, "XXXLiefNrXXX", rsAccess("L-Nummer/AVISTA"))
'            Set rsAvista = adbprov.getRecordsetAvista(SQL)
'            If rsAvista.RecordCount Then
'
''            If rsAccess("L-Nummer/AVISTA") = 72470 Then
''                Debug.Print rsAccess("L-Nummer/AVISTA")
''            End If
'
'                blnGesperrt = False
'                blnLQLief = False
'                If rsAvista("adgrp_kuerzel") = "Gesperrt" Then blnGesperrt = True
'                If rsAvista("adl_info4") = "LQ" Then blnLQLief = True
'                dbAccess.Execute ("Update Lieferanten set [Matchcode] = '" & rsAvista("pa_matchcode") & "', [Lieferantenname] = '" & rsAvista("pa_org_name1") & "', " & _
'                "[Telefaxnummer] = '" & rsAvista("adb_fax") & "', [Telefonnummer] = '" & rsAvista("adb_tel1") & "',[E-Mail] = '" & rsAvista("adb_email") & "', " & _
'                "[Adresse]= '" & rsAvista("adb_strasse") & "',[Postleitzahl] = " & rsAvista("adb_plz_str") & ",[Ort] = '" & rsAvista("adb_ort") & "', " & _
'                "[L-Kategorie] = '" & rsAvista("adgrp_kuerzel") & "',[VerantwortlichEinkauf] = '" & rsAvista("adl_info1") & "', " & _
'                "[GESPERRT] = " & CInt(blnGesperrt) & " ,[LieferantenverantwortungLQ] = " & CInt(blnLQLief) & " " & _
'                "where [L-Nummer/AVISTA] = " & rsAccess("L-Nummer/AVISTA") & ";")
'            End If
'            rsAccess.MoveNext
'        Wend
'    End If
'    Me.Refresh

    Call LiefDatenAbrufen

End Sub

Private Sub FilterLieferant_Change()        '20190827-ROS_1
On Error Resume Next

    Dim strFilter As String
    Dim intStart As Integer

    intStart = Me.FilterLieferant.SelStart

    If Not Len(Me.FilterLieferant.Text) = 0 Then
        strFilter = "[Lieferantenname] like '" & Me.FilterLieferant.Text & "*'"
        Me.Filter = strFilter
        Me.FilterOn = True
        Me.FilterLieferant.SelStart = intStart
        Me.FilterLieferant.SetFocus
    Else
        Me.FilterLieferant = ""
        Me.FilterOn = False
        Me.FilterLieferant.SetFocus
    End If

End Sub

Private Sub FilterLöschen_Click()

Dim sfilter As String

    sfilter = ""
    Me.Filter = ""
    Me.FilterOn = True
    Me.FilterESachbearbeiter = ""
    Me.FilterLiefVerantwortung = 3
    Me.FilterLKategorie = ""
    Me.FilterLieferant = ""
    Me.FilterGesperrt = Null
    Me.FilterAnzeigen = Null
    sfilter = ""
    Me.Filter = sfilter
    Me.Requery
    Me.Refresh


End Sub

Private Sub FilterSetzen_Click()

Dim lFilter As Boolean
Dim sfilter As String
Dim lFilterAnzeigen As Boolean

sfilter = ""
lFilter = False
If Me.FilterLKategorie <> "" Then
    sfilter = "[L-Kategorie] = '" & Me.FilterLKategorie & "'"
    lFilter = True
End If
If Me.FilterLiefVerantwortung < 3 Then
    If lFilter Then
        sfilter = sfilter & " and [LieferantenverantwortungLQ] = "
        If Me.FilterLiefVerantwortung = 1 Then
            sfilter = sfilter & "True"
        Else
            sfilter = sfilter & "False"
        End If
    Else
        sfilter = "[LieferantenverantwortungLQ] = "
        If Me.FilterLiefVerantwortung = 1 Then
            sfilter = sfilter & "True"
        Else
            sfilter = sfilter & "False"
        End If
    End If
    lFilter = True
End If
If Me.FilterLieferant <> "" Then
    If lFilter Then
        sfilter = sfilter & " and [Lieferantenname] like '" & Me.FilterLieferant & "*'"
    Else
        sfilter = "[Lieferantenname] like '" & Me.FilterLieferant & "*'"
    End If
    lFilter = True
End If

If Me.FilterESachbearbeiter <> "" Then
    If lFilter Then
        sfilter = sfilter & " and [VerantwortlichEinkauf] = '" & Me.FilterESachbearbeiter & "'"
    Else
        sfilter = "[VerantwortlichEinkauf] like '" & Me.FilterESachbearbeiter & "'"
        lFilter = True
    End If
End If

'Setzen Filter Gesperrt
If Not IsNull(Me.FilterAnzeigen) Then
    If lFilter Then
        sfilter = sfilter & " and [Anzeigen] = " & Me.FilterAnzeigen
    Else
        sfilter = "[Anzeigen] = " & Me.FilterAnzeigen
        lFilter = True
    End If
End If

If Not IsNull(Me.FilterGesperrt) Then
    If lFilter Then
        sfilter = sfilter & " and [GESPERRT] = " & Me.FilterGesperrt
    Else
        sfilter = "[GESPERRT] = " & Me.FilterGesperrt
    End If
End If

Me.Filter = sfilter
Me.FilterOn = True
Me.Requery
Me.Refresh

'MsgBox ("Anzahl Datensätze:" & Me.Count)

End Sub



Private Sub Form_DblClick(Cancel As Integer)

    Befehl9_Click

End Sub

Private Sub Form_Load()

    Dim objMouseOverForm As clsMouseoverForm
    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With


'    Dim varKoord As Variant
'    varKoord = getFktKoordinatenRelHF(Me)
'    If varKoord(3) > 0 And varKoord(2) > 0 Then
'        'Fenster setzen relativ zum Hauptfenster TB:autostart
'        DoCmd.MoveSize varKoord(0), varKoord(1), varKoord(2), varKoord(3)
'    End If
'    DoCmd.Restore

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub

Private Sub Form_Open(Cancel As Integer)
    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)
End Sub

Private Sub Neu_Click()

    Dim adbprov As clsAvistaProvider
    Dim rsAvista As ADODB.Recordset
    Dim dbAccess As DAO.Database
    Dim rsAccess As DAO.Recordset
    Dim strSQL  As String
    Dim sql As String
    Dim blnGesperrt As Boolean
    Dim blnLQLief As Boolean
    Dim intNeuLiefID As Long
    Dim intAvistaLiefNr As String
    Dim strAvistaLiefNr As String
    Dim strLiefname As String
    Dim rs As DAO.Recordset
    Dim ingLiefID As Integer

    '########20221124-ROS#########
    ' Code angepasst zum anlegen von Internen Liefreanten

    Set dbAccess = openAccDBConnection("openDB")
    Set rs = dbAccess.OpenRecordset("Select * from [Lieferanten];")

'    strAvistaLiefNr = InputBox("Avista Lieferanten Nummer eingeben")
'    strAvistaLiefNr = getValueFromInputBox("Avista Lieferanten Nummer eingeben", "Avista Lieferanten-Nummer?")
    strAvistaLiefNr = getValueFromInputBox(getTC_FCFromCode("IC4287"), getTC_FCFromCode("IC4288"))
    '########20221124-ROS#########
    'Lieferant anlegen ohne AVISTA Kundennummer
    If strAvistaLiefNr = "" Then
'        strLiefname = getValueFromInputBox("Lieferanten Namen eingeben", "Lieferanten-Name?")
        strLiefname = getValueFromInputBox(getTC_FCFromCode("IC4289"), getTC_FCFromCode("IC4290"))
        If strLiefname = "0" Then
            Set rs = Nothing
            Exit Sub
        End If
        rs.MoveFirst
        rs.MoveLast
        ingLiefID = rs("Lieferant-ID") + 1
        dbAccess.Execute ("Insert into Lieferanten([Lieferant-ID],[L-Kategorie],[Lieferantenname], [Matchcode])values('" & ingLiefID & "','Lieferant A','" & strLiefname & "','" & strLiefname & "' );")
        Me.Filter = "[Lieferant-ID] = " & ingLiefID
        Set rs = Nothing
    Else
        If strAvistaLiefNr = 0 Then
            Set rs = Nothing
            Exit Sub
        Else
        '########20221124-ROS#########
        'Lieferant anlegen mit AVISTA Kundennummer
            intAvistaLiefNr = strAvistaLiefNr
                'mit diesem String werden die Daten aus AVISTA abgerufen.
            strSQL = "SELECT tm_lieferant.adl_nummer, tm_lieferant.adl_info1, tm_lieferant.adl_info4, ta_adress_gruppe.adgrp_kuerzel, tm_partei.pa_matchcode, tm_partei.pa_org_name1, " & _
                "tm_adress_basis.adb_strasse , tm_adress_basis.adb_plz_str, tm_adress_basis.adb_ort, tm_adress_basis.adb_tel1, tm_adress_basis.adb_fax, tm_adress_basis.adb_email " & _
                "From DBA.tm_partei " & _
                "INNER JOIN DBA.ta_adress_gruppe ON tm_partei.adgrp_id = ta_adress_gruppe.adgrp_id " & _
                "INNER JOIN DBA.tm_lieferant ON tm_partei.pa_id =  tm_lieferant.adl_id AND tm_lieferant.ma_id = 2 " & _
                "INNER JOIN DBA.tm_adress_basis ON tm_partei.pa_id = tm_adress_basis.pa_id " & _
                "WHERE tm_lieferant.ma_id = 2  AND tm_adress_basis.adt_id = 1 AND tm_partei.adgrp_id <> 1 AND tm_lieferant.adl_nummer = 'XXXLiefNrXXX' ORDER BY pa_matchcode, adl_nummer;"

            blnGesperrt = False
            blnLQLief = False
            Set adbprov = New clsAvistaProvider
            sql = Replace(strSQL, "XXXLiefNrXXX", intAvistaLiefNr)
            Set rsAvista = adbprov.getRecordsetAvista(sql)
            Set rsAccess = dbAccess.OpenRecordset("Select [L-Nummer/AVISTA] from Lieferanten where [L-Nummer/AVISTA] = " & intAvistaLiefNr & " ;")
            'Lieferant ist in der Access-Datenbank schon angelegt dann werden die Daten nur aktuallisiert
            If rsAccess.RecordCount Then
'                MsgBox ("Lieferant schon angelegt, Daten werden aktuallisiert")
                MsgBox getTC_FCFromCode("TC4238"), vbInformation, getTC_FCFromCode("TC4197")
                If rsAvista.RecordCount Then
                    If rsAvista("adgrp_kuerzel") = "Gesperrt" Then blnGesperrt = True
                    If rsAvista("adl_info4") = "LQ" Then blnLQLief = True
                    dbAccess.Execute ("Update Lieferanten set [Matchcode] = '" & rsAvista("pa_matchcode") & "', [Lieferantenname] = '" & rsAvista("pa_org_name1") & "', " & _
                                    "[Telefaxnummer] = '" & rsAvista("adb_fax") & "', [Telefonnummer] = '" & rsAvista("adb_tel1") & "',[E-Mail] = '" & rsAvista("adb_email") & "', " & _
                                    "[Adresse]= '" & rsAvista("adb_strasse") & "',[Postleitzahl] = " & rsAvista("adb_plz_str") & ",[Ort] = '" & rsAvista("adb_ort") & "', " & _
                                    "[L-Kategorie] = '" & rsAvista("adgrp_kuerzel") & "',[VerantwortlichEinkauf] = '" & rsAvista("adl_info1") & "', " & _
                                    "[GESPERRT] = " & CInt(blnGesperrt) & " ,[LieferantenverantwortungLQ] = " & CInt(blnLQLief) & " " & _
                                    "where [L-Nummer/AVISTA] = " & rsAccess("L-Nummer/AVISTA") & ";")
            End If
            Else

                'Lieferant ist in der Access-Datenbank noch nicht angelegt dann wir ein neuer Datensatz angelegt
                'Lieferant-ID wird um 1 hochgezählt
                intNeuLiefID = DMax("[Lieferant-ID]", "Lieferanten") + 1
                If rsAvista("adgrp_kuerzel") = "Gesperrt" Then blnGesperrt = True
                If rsAvista("adl_info4") = "LQ" Then blnLQLief = True
                dbAccess.Execute ("Insert into Lieferanten ([Lieferant-ID],[L-Nummer/AVISTA],[Matchcode],[Lieferantenname],[Telefaxnummer],[Telefonnummer],[E-Mail],[Adresse],[Postleitzahl],[Ort],[L-Kategorie],[VerantwortlichEinkauf]," & _
                                "[GESPERRT],[LieferantenverantwortungLQ])values " & _
                                "(" & intNeuLiefID & "," & intAvistaLiefNr & ",'" & rsAvista("pa_matchcode") & "', '" & rsAvista("pa_org_name1") & "','" & rsAvista("adb_fax") & "','" & rsAvista("adb_tel1") & "','" & rsAvista("adb_email") & "', " & _
                                "'" & rsAvista("adb_strasse") & "'," & rsAvista("adb_plz_str") & ",'" & rsAvista("adb_ort") & "','" & rsAvista("adgrp_kuerzel") & "','" & rsAvista("adl_info1") & "', " & _
                                "" & CInt(blnGesperrt) & "," & CInt(blnLQLief) & ");")
            End If
        End If
        Me.Filter = "[L-Nummer/AVISTA] = " & intAvistaLiefNr
    End If

    Me.FilterOn = True

    Set rs = Nothing
    Set rsAvista = Nothing
    Set rsAccess = Nothing
    Set adbprov = Nothing

End Sub

Private Sub LiefDatenAbrufen(Optional sLiefNrAvista As Long)    'Optional sLiefNrAvista für das Updaten von nur einen Lieferanten sonst werden alle Lieferanten aktuallisiert
    On Error Resume Next

    Dim adbprov As clsAvistaProvider
    Dim rsAvista As ADODB.Recordset
    Dim dbAccess As DAO.Database
    Dim rsAccess As DAO.Recordset
    Dim strSQL  As String
    Dim sql As String
    Dim blnGesperrt As Boolean
    Dim blnLQLief As Boolean

    'mit diesem String werden die Daten aus AVISTA abgerufen.
    strSQL = "SELECT tm_lieferant.adl_nummer, tm_lieferant.adl_info1, tm_lieferant.adl_info4, ta_adress_gruppe.adgrp_kuerzel, tm_partei.pa_matchcode, tm_partei.pa_org_name1, " & _
                "tm_adress_basis.adb_strasse , tm_adress_basis.adb_plz_str, tm_adress_basis.adb_ort, tm_adress_basis.adb_tel1, tm_adress_basis.adb_fax, tm_adress_basis.adb_email " & _
                "From DBA.tm_partei " & _
                "INNER JOIN DBA.ta_adress_gruppe ON tm_partei.adgrp_id = ta_adress_gruppe.adgrp_id " & _
                "INNER JOIN DBA.tm_lieferant ON tm_partei.pa_id =  tm_lieferant.adl_id AND tm_lieferant.ma_id = 2 " & _
                "INNER JOIN DBA.tm_adress_basis ON tm_partei.pa_id = tm_adress_basis.pa_id " & _
                "WHERE tm_lieferant.ma_id = 2  AND tm_adress_basis.adt_id = 1 AND tm_partei.adgrp_id <> 1 AND tm_lieferant.adl_nummer = 'XXXLiefNrXXX' ORDER BY pa_matchcode, adl_nummer;"

    Set dbAccess = openAccDBConnection("openDB")
    Set adbprov = New clsAvistaProvider
    Set rsAccess = dbAccess.OpenRecordset("Select [L-Nummer/AVISTA] from Lieferanten;")
    If sLiefNrAvista = 0 Then
        DoCmd.Hourglass True
        If rsAccess.RecordCount Then
            rsAccess.MoveFirst
            'Schleife über alle Liferanten die in der Access-Datenbank vorhanden sind
            While Not rsAccess.EOF
                'Lieferanten-Nummer(XXXLiefNrXXX) wird im SQL ersetzt duch die Access L-Nummer/AVISTA damit nur die Werte eines Lieferanten in AVISTA abgerufen werden
                sql = Replace(strSQL, "XXXLiefNrXXX", rsAccess("L-Nummer/AVISTA"))
                Set rsAvista = adbprov.getRecordsetAvista(sql)
                If rsAvista.RecordCount Then
    '            If rsAccess("L-Nummer/AVISTA") = 72470 Then
    '                Debug.Print rsAccess("L-Nummer/AVISTA")
    '            End If
                    blnGesperrt = False
                    blnLQLief = False
                    If rsAvista("adgrp_kuerzel") = "Gesperrt" Then blnGesperrt = True
                    If rsAvista("adl_info4") = "LQ" Then blnLQLief = True
                    dbAccess.Execute ("Update Lieferanten set [Matchcode] = '" & rsAvista("pa_matchcode") & "', [Lieferantenname] = '" & rsAvista("pa_org_name1") & "', " & _
                                    "[Telefaxnummer] = '" & rsAvista("adb_fax") & "', [Telefonnummer] = '" & rsAvista("adb_tel1") & "',[E-Mail] = '" & rsAvista("adb_email") & "', " & _
                                    "[Adresse]= '" & rsAvista("adb_strasse") & "',[Postleitzahl] = " & rsAvista("adb_plz_str") & ",[Ort] = '" & rsAvista("adb_ort") & "', " & _
                                    "[L-Kategorie] = '" & rsAvista("adgrp_kuerzel") & "',[VerantwortlichEinkauf] = '" & rsAvista("adl_info1") & "', " & _
                                    "[GESPERRT] = " & CInt(blnGesperrt) & " ,[LieferantenverantwortungLQ] = " & CInt(blnLQLief) & " " & _
                                    "where [L-Nummer/AVISTA] = " & rsAccess("L-Nummer/AVISTA") & ";")
                End If
                rsAccess.MoveNext
            Wend
        End If
        DoCmd.Hourglass False
    Else
        'Update eines bestimmten Lieferanten
        'Lieferanten-Nummer(XXXLiefNrXXX) wird im SQL ersetzt duch die Access sLiefNrAvista damit nur diese Werte des Lieferanten in AVISTA abgerufen werden
        sql = Replace(strSQL, "XXXLiefNrXXX", sLiefNrAvista)
        Set rsAvista = adbprov.getRecordsetAvista(sql)
        If rsAvista.RecordCount Then
            blnGesperrt = False
            blnLQLief = False
            If rsAvista("adgrp_kuerzel") = "Gesperrt" Then blnGesperrt = True
            If rsAvista("adl_info4") = "LQ" Then blnLQLief = True
            dbAccess.Execute ("Update Lieferanten set [Matchcode] = '" & rsAvista("pa_matchcode") & "', [Lieferantenname] = '" & rsAvista("pa_org_name1") & "', " & _
                            "[Telefaxnummer] = '" & rsAvista("adb_fax") & "', [Telefonnummer] = '" & rsAvista("adb_tel1") & "',[E-Mail] = '" & rsAvista("adb_email") & "', " & _
                            "[Adresse]= '" & rsAvista("adb_strasse") & "',[Postleitzahl] = " & rsAvista("adb_plz_str") & ",[Ort] = '" & rsAvista("adb_ort") & "', " & _
                            "[L-Kategorie] = '" & rsAvista("adgrp_kuerzel") & "',[VerantwortlichEinkauf] = '" & rsAvista("adl_info1") & "', " & _
                            "[GESPERRT] = " & CInt(blnGesperrt) & " ,[LieferantenverantwortungLQ] = " & CInt(blnLQLief) & " " & _
                            "where [L-Nummer/AVISTA] = " & sLiefNrAvista & ";")
        End If
    End If

    Set rsAccess = Nothing
    Set adbprov = Nothing

    Me.Refresh

End Sub
