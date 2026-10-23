Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Betrag_AfterUpdate()

    Me.Aufwand = Me.StdSatz * Me.Zeit

End Sub

Private Sub cmbChangeStdSatz_Click()

    If get_ServiceZuordnung = True Then
        DoCmd.OpenForm "frmAuswahlStundensatz", acNormal, , , , , "MID=" & Me.VorgangID
    End If

End Sub

Private Sub cmbChangeStdSatz_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub cmbEmailSendTo_Click()

    UpdatetFormTab Forms("Fehlererfassung - Register"), "[ID]=" & Forms("Fehlererfassung - Register").Controls("Mastertab#ID")
    DoCmd.OpenForm ("frmEmailService")

End Sub

Private Sub cmbEmailSendTo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub

Private Sub Form_AfterUpdate()

    Dim ldb As DAO.Database
    Dim lZeitaufwand As Double

    Set ldb = openAccDBConnection("openDB")

    lZeitaufwand = DSum("[Zeit]", "AufwandVorgangPositionen", "[MID] = " & Forms("Fehlererfassung - Register").Controls("Mastertab#ID"))

    ldb.Execute ("Update Mastertab set [Zeitbedarf] = '" & lZeitaufwand & "' where ID = " & Forms("Fehlererfassung - Register").Controls("Mastertab#ID") & ";")

End Sub


Private Sub Form_Load()

    Me.cmbEmailSendTo.Visible = False

    If get_ServiceZuordnung = True Then
        Me.cmbEmailSendTo.Visible = True
    End If


End Sub

Private Sub txtInfoPos_DblClick(Cancel As Integer)

    If Forms("Fehlererfassung - Register").Controls("Mastertab#Kategorie-Status") = "Abgeschlossen" Then
        DoCmd.OpenForm ("frmInfoPos"), , , "[VorgangID] = " & Me.VorgangID, acFormReadOnly
    Else
        DoCmd.OpenForm ("frmInfoPos"), , , "[VorgangID] = " & Me.VorgangID, acFormEdit
    End If

End Sub

Private Sub Vorgang_AfterUpdate()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim strAbteilung As String
    Dim strVorgang As String
    Dim lStdsatzVorgangID As Integer  'Stundensatz Vorgang-Text ID
    Dim lStdsatzVorgang As Double  'Stundensatz Vorgang


    lStdsatzVorgangID = Nz(DLookup("StdSatzID", "Vorgang-Text", "[Vorgang-ID] = " & Me.Vorgang & ""))
    lStdsatzVorgang = Nz(DLookup("Betrag", "tblAufwandStdSaetze", "[StdSatzID] = " & lStdsatzVorgangID & ""))


    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("SELECT AufwandVorgangPositionen.VorgangID, AufwandVorgangPositionen.Datum  " & _
                            "From AufwandVorgangPositionen " & _
                            "WHERE [AufwandVorgangPositionen.MID] = " & Me.ReklaID & " AND [AufwandVorgangPositionen.Vorgang]= 78;")


'Bei Auswahl des Reiters "Rücklieferung versandfertig(Lieferant / Hersteller)" ID26
'Abfrage ob Frachkosten eingetragen werden müssen über MsgBox
'wenn Ja Frachtkosten automatisch in Kostenrechnung eintragen; Pauschal 7€; damit Kosten überhaupt erfasst werden
    If Me.Vorgang = 26 Then
        If MsgBox(getTC_FCFromCode("TC4192"), vbYesNo, getTC_FCFromCode("TC4193")) = vbYes Then
            Call Forms("Fehlererfassung - Register").PosKostenAnlegen("Frachtkosten")
        Else
            Exit Sub
        End If
    End If
'Datum in Ergebnis-Daten eintragen wenn Stellungnahme versendet wurde, in Feld "Rückmeldung_Kunde_Ist" 20140118-ROS
    If Me.Vorgang = 78 Then
        strAbteilung = Nz(DLookup("[Bürostandort]", "Bearbeiter", "[Kurzzeichen] = '" & Me.Kürzel & "'"))
        If strAbteilung = "Service" Then
            ldb.Execute ("UPDATE [Mastertab] set [Rückmeldung_Kunde_Ist]=  '" & Me.Datum & "' where [Mastertab.ID] = " & Me.ReklaID & " ;")
        End If
    Else
        If rs.RecordCount Then
            rs.MoveFirst
            rs.MoveLast
            If rs.RecordCount > 1 Then
                rs.MovePrevious
                ldb.Execute ("UPDATE [Mastertab] set [Rückmeldung_Kunde_Ist]=  '" & rs("Datum") & "' where [Mastertab.ID] = " & Me.ReklaID & " ;")
            End If
        End If
    End If

    If lStdsatzVorgangID <> Me.StdSatzID Then
'        Debug.Print
        Me.StdSatzID = lStdsatzVorgangID
        Me.DatumStdSatz = Date
        Me.StdSatz = lStdsatzVorgang
        Me.Aufwand = Me.Zeit * Me.StdSatz
    End If
    Me.Refresh
    Set rs = Nothing

End Sub

Public Sub Zeit_AfterUpdate()

    Dim lPos As Integer


    lPos = Me.Pos

    Me.Requery
    Forms("Fehlererfassung - Register").Controls("Mastertab#Zeitbedarf") = DSum("[Zeit]", "AufwandVorgangPositionen", "[MID] = " & Me.ReklaID)
    Forms("Fehlererfassung - Register").Refresh
    Me.RecordsetClone.FindFirst "[Pos] = " & lPos & ""
    Me.Bookmark = Me.RecordsetClone.Bookmark
    Me.Controls("Vorgang").SetFocus
    Me.Aufwand = Me.StdSatz * Me.Zeit

End Sub
