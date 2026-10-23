Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Const csKundenAuswahl = "SELECT Kunden.Firma " _
                        & "FROM Kunden INNER JOIN Mastertab ON Kunden.[Kunden-ID] = Mastertab.[Kunden-ID] " _
                        & "WHERE (((Mastertab.[Kategorie-Status]) <> 'Abgeschlossen')) "
Const csGroupByKunde = "GROUP BY Kunden.Firma "
Const csSortByDate = "ORDER BY Mastertab.Rückmeldung_Kunde_Soll;"

Const FilterAlleRechnungen = 8
Const FilterKunde = 4
Const FilterRueckmeldung = 2
Const FilterMeldenummer = 1

Public RekFilter As Byte





Private Sub AlleKunden_Click()

    Me.Kundenauswahl = Null
    RekFilter = RekFilter And (255 - 4) '-> 1111 1011

    SetFilterOffeneRekl
    Me.Requery
    Me.Refresh

End Sub



Private Sub AlleReklamationen2_AfterUpdate()

    If Me.AlleReklamationen2 Then
        'Set Bit
        RekFilter = RekFilter Or FilterAlleRechnungen
    Else
        RekFilter = RekFilter And (255 - 8) '-> 1111 1000
    End If

    SetFilterOffeneRekl
    Me.Requery
    Me.Refresh

End Sub

Private Sub Form_DblClick(Cancel As Integer)
'[ID]=[ID]![Fehlererfassung - Register]

    DoCmd.OpenForm "Fehlererfassung - Register", acNormal, , "[ID]=" & Me.[ID]
    'Forms("Fehlererfassung - Register").Filter = "[ID]=" & Me.[ID]
    Forms("Fehlererfassung - Register").AllowAdditions = False


End Sub

Private Sub Form_Load()

    'Initialisierung auf alle offenen Reklamtionen ohne Einschränkungen
    '-> Filtereintrag löschen
    RekFilter = 0
    Me.AlleReklamationen2.Value = 1
    RekFilter = RekFilter Or FilterAlleRechnungen
    Me.Kundenauswahl = ""
    Me.Filter = ""
    Me.RueckmeldedatumAnKunde.Value = 0
    'RekFilter = RekFilter Or 2
    Me.MeldenummerVorhanden.Value = 0
    'RekFilter = RekFilter Or 1
    'MsgBox ("Rekfilter: " + RekFilter)
    SetFilterOffeneRekl
    'DoCmd.Maximize


End Sub


Private Sub Form_Open(Cancel As Integer)
    'Me.AlleReklamationen.ForeColor = 65280
    'DoCmd.Maximize
End Sub

Private Sub Kundenauswahl_AfterUpdate()

    If Me.Kundenauswahl <> "" Then
        RekFilter = RekFilter Or FilterKunde
        SetFilterOffeneRekl

        Me.Requery
        Me.Refresh
    End If

End Sub
'(((Mastertab.[Kategorie-Status])<>"Abgeschlossen"))
Private Sub SetFilterOffeneRekl()
    bFilterOn = False
    sfilter = ""
    sKndAuswQuery = csKundenAuswahl

    If RekFilter And FilterAlleRechnungen Then
        sfilter = sfilter + "((Mastertab.[Kategorie-Status] <> 'Abgeschlossen')"
        bFilterOn = True
    End If

    If RekFilter And FilterKunde Then
        If Not bFilterOn Then
            sfilter = sfilter + "((Kunden.[Firma] = '" + Me.Kundenauswahl + "')"
            bFilterOn = True
        Else
            sfilter = sfilter + " And (Kunden.[Firma] = '" + Me.Kundenauswahl + "')"
        End If
    End If

    If RekFilter And FilterRueckmeldung Then
        If Not bFilterOn Then
            sfilter = sfilter + "((Mastertab.[Rückmeldung_Kunde_Soll] Is not Null)"
            bFilterOn = True
        Else
            sfilter = sfilter + " And (Mastertab.[Rückmeldung_Kunde_Soll] Is not Null)"
        End If
    End If

    If RekFilter And FilterMeldenummer Then
        If Not bFilterOn Then
            sfilter = sfilter + "((Mastertab.[Meldungsnummer Kunde] Is not Null)"
            bFilterOn = True
        Else
            sfilter = sfilter + " And (Mastertab.[Meldungsnummer Kunde] Is not Null)"
        End If
    End If



    If bFilterOn Then
        sfilter = sfilter + ")"
        Me.Filter = sfilter
        Me.FilterOn = True
    Else
        Me.Filter = ""
        Me.FilterOn = False
    End If
    Me.FilterView = Me.Filter

End Sub
Private Sub SetFilterKundenauswahl()
    sKndAuswQuery = csKundenAuswahl & csGroupByKunde & ";"
    'Me.Filter = sFilter
    Me.Kundenauswahl.RowSource = sKndAuswQuery

End Sub

Private Sub MeldenummerVorhanden_AfterUpdate()

    If Me.MeldenummerVorhanden Then
        'Set Bit
        RekFilter = RekFilter Or FilterMeldenummer
    Else
        RekFilter = RekFilter And (255 - 1) '-> 1111 1110
    End If

    SetFilterOffeneRekl
    Me.Requery
    Me.Refresh

End Sub

Private Sub RueckmeldedatumAnKunde_AfterUpdate()

    If Me.RueckmeldedatumAnKunde Then
        'Set Bit
        RekFilter = RekFilter Or FilterRueckmeldung
    Else
        RekFilter = RekFilter And (255 - 2) '-> 1111 1101
    End If

    SetFilterOffeneRekl
    Me.Requery
    Me.Refresh

End Sub
