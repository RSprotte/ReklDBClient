Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub FilterSetzen_Click()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim bFilterOn As Boolean
    Dim sfilter As String
    Dim iMehfachausw As Integer

    bFilterOn = False
    sfilter = ""
    sfilter = "("
    iMehfachausw = 0

    If Not IsNull(Me.txtMaßnahmenID) Then
        If bFilterOn Then
            sfilter = sfilter + " And "
        End If

        sfilter = sfilter + "([MaßnahmeID] =   " & Me.txtMaßnahmenID & " )"
        bFilterOn = True

    End If

    If Not IsNull(Me.ListMaßnahmenHerkunft) Then
        If bFilterOn Then
            sfilter = sfilter + " And "
        End If

        sfilter = sfilter + "([MaßnahmenHerkunftID] = " & Me.ListMaßnahmenHerkunft & ")"
        bFilterOn = True

    End If


    If Not IsNull(Me.ListVerantwortlich) Then
        If bFilterOn Then
            sfilter = sfilter + " And "
        End If

        sfilter = sfilter + "([Verantwortlich] = " & Me.ListVerantwortlich & ")"
        bFilterOn = True

    End If
    If Not IsNull(Me.ListArtMaßnahme) Then
        If bFilterOn Then
            sfilter = sfilter + " And "
        End If

        sfilter = sfilter + "([ArtDerMaßnahme] = " & Me.ListArtMaßnahme & ")" '20160318-ROS '" & Me.ListArtMaßnahme & "'
        bFilterOn = True

    End If
    If Not IsNull(Me.ListLQBereich) Then
        If bFilterOn Then
            sfilter = sfilter + " And "
        End If

        sfilter = sfilter + "(MaßnahmenFehlerortGruppe.[LQBereichID] = " & Me.ListLQBereich & ")"
        bFilterOn = True
        iMehfachausw = iMehfachausw + 2
    End If

    If Not IsNull(Me.ListLQEinheit) Then
        If bFilterOn Then
            sfilter = sfilter + " And "
        End If

        sfilter = sfilter + "(LQBereich.[ID-LQEinheit] = " & Me.ListLQEinheit & ")"
        bFilterOn = True
        iMehfachausw = iMehfachausw + 1
    End If

    If Not IsNull(Me.ListTätigkeit) Then
        If bFilterOn Then
            sfilter = sfilter + " And "
        End If

        sfilter = sfilter + "(MaßnahmenFehlerarten.[Fehlerklasse] = '" & Me.ListTätigkeit & "' )"
        bFilterOn = True
        iMehfachausw = iMehfachausw + 4
    End If

    If Me.CtrlInBearbeitung Then
        If bFilterOn Then
            sfilter = sfilter + " And "
        End If

        sfilter = sfilter + "([UmgesetztAm]Is Null)"
        bFilterOn = True
    End If

    If Me.CtrlAbgeschlossen Then
        If bFilterOn Then
            sfilter = sfilter + " And "
        End If

        sfilter = sfilter + "([UmgesetztAm]Is Not Null)"
        bFilterOn = True
    End If

    If Not bFilterOn Then
        sfilter = ""
    Else
        sfilter = sfilter + ")"
    End If

    'Filter in Übersicht setzen und aktualisieren
    Select Case iMehfachausw

        Case 0
            Forms("Maßnahmen").RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige " & _
                        "FROM ((Maßnahmen LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) LEFT JOIN MaßnahmenHerkunft ON Maßnahmen.MaßnahmenHerkunftID = MaßnahmenHerkunft.ID) LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme] " & _
                        "GROUP BY Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige " & _
                        "ORDER BY Maßnahmen.MaßnahmeID;"
        Case 1
            Forms("Maßnahmen").RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige, LQBereich.[ID-LQEinheit] " & _
                        "FROM ((((Maßnahmen LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) LEFT JOIN MaßnahmenHerkunft ON Maßnahmen.MaßnahmenHerkunftID = MaßnahmenHerkunft.ID) LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]) INNER JOIN MaßnahmenFehlerortGruppe ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerortGruppe.MaßnahmeID) INNER JOIN LQBereich ON MaßnahmenFehlerortGruppe.LQBereichID = LQBereich.ID " & _
                        "GROUP BY Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige, LQBereich.[ID-LQEinheit] " & _
                        "ORDER BY Maßnahmen.MaßnahmeID;"

        Case 2
            Forms("Maßnahmen").RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige, MaßnahmenFehlerortGruppe.LQBereichID " & _
                        "FROM (((Maßnahmen LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) LEFT JOIN MaßnahmenHerkunft ON Maßnahmen.MaßnahmenHerkunftID = MaßnahmenHerkunft.ID) LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]) LEFT JOIN MaßnahmenFehlerortGruppe ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerortGruppe.MaßnahmeID " & _
                        "GROUP BY Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige, MaßnahmenFehlerortGruppe.LQBereichID " & _
                        "ORDER BY Maßnahmen.MaßnahmeID;"

        Case 3
            Forms("Maßnahmen").RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige, MaßnahmenFehlerortGruppe.LQBereichID, LQBereich.[ID-LQEinheit] " & _
                        "FROM ((((Maßnahmen LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) LEFT JOIN MaßnahmenHerkunft ON Maßnahmen.MaßnahmenHerkunftID = MaßnahmenHerkunft.ID) LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]) INNER JOIN MaßnahmenFehlerortGruppe ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerortGruppe.MaßnahmeID) INNER JOIN LQBereich ON MaßnahmenFehlerortGruppe.LQBereichID = LQBereich.ID " & _
                        "GROUP BY Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige, MaßnahmenFehlerortGruppe.LQBereichID, LQBereich.[ID-LQEinheit] " & _
                        "ORDER BY Maßnahmen.MaßnahmeID;"
        Case 4
            Forms("Maßnahmen").RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige, MaßnahmenFehlerarten.Fehlerklasse " & _
                        "FROM (((Maßnahmen LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) LEFT JOIN MaßnahmenHerkunft ON Maßnahmen.MaßnahmenHerkunftID = MaßnahmenHerkunft.ID) LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]) LEFT JOIN MaßnahmenFehlerarten ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerarten.MaßnahmeID " & _
                        "GROUP BY Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige, MaßnahmenFehlerarten.Fehlerklasse " & _
                        "ORDER BY Maßnahmen.MaßnahmeID;"
        Case 5

            Forms("Maßnahmen").RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige, LQBereich.[ID-LQEinheit], MaßnahmenFehlerarten.Fehlerklasse " & _
                        "FROM (((((Maßnahmen LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) LEFT JOIN MaßnahmenHerkunft ON Maßnahmen.MaßnahmenHerkunftID = MaßnahmenHerkunft.ID) LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]) INNER JOIN MaßnahmenFehlerortGruppe ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerortGruppe.MaßnahmeID) INNER JOIN LQBereich ON MaßnahmenFehlerortGruppe.LQBereichID = LQBereich.ID) INNER JOIN MaßnahmenFehlerarten ON MaßnahmenFehlerortGruppe.MaßnahmeID = MaßnahmenFehlerarten.MaßnahmeID " & _
                        "GROUP BY Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige, LQBereich.[ID-LQEinheit], MaßnahmenFehlerarten.Fehlerklasse " & _
                        "ORDER BY Maßnahmen.MaßnahmeID;"
        Case 6

            Forms("Maßnahmen").RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige, MaßnahmenFehlerortGruppe.LQBereichID, MaßnahmenFehlerarten.Fehlerklasse " & _
                        "FROM ((((Maßnahmen LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) LEFT JOIN MaßnahmenHerkunft ON Maßnahmen.MaßnahmenHerkunftID = MaßnahmenHerkunft.ID) LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]) INNER JOIN MaßnahmenFehlerortGruppe ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerortGruppe.MaßnahmeID) INNER JOIN MaßnahmenFehlerarten ON MaßnahmenFehlerortGruppe.MaßnahmeID = MaßnahmenFehlerarten.MaßnahmeID " & _
                        "GROUP BY Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige, MaßnahmenFehlerortGruppe.LQBereichID, MaßnahmenFehlerarten.Fehlerklasse " & _
                        "ORDER BY Maßnahmen.MaßnahmeID;"

        Case 7
            Forms("Maßnahmen").RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], " & _
                        "Maßnahmen.StartUmsetzung , Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, " & _
                        "Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, LQBereich.[ID-LQEinheit], MaßnahmenFehlerortGruppe.LQBereichID, MaßnahmenFehlerarten.Fehlerklasse, " & _
                        "Maßnahmen.ArtDerMaßnahme , MaßnahmeUmsetzungFrist.Anzeige " & _
                        "FROM (((((Maßnahmen LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID])" & _
                        "LEFT JOIN MaßnahmenHerkunft ON Maßnahmen.MaßnahmenHerkunftID = MaßnahmenHerkunft.ID)" & _
                        "INNER JOIN MaßnahmenFehlerortGruppe ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerortGruppe.MaßnahmeID)" & _
                        "INNER JOIN LQBereich ON MaßnahmenFehlerortGruppe.LQBereichID = LQBereich.ID)" & _
                        "INNER JOIN MaßnahmenFehlerarten ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerarten.MaßnahmeID)" & _
                        "INNER JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]" & _
                        "ORDER BY Maßnahmen.MaßnahmeID;"

    End Select


'    Select Case Case iMehfachausw
'        Case 0
'            Forms("Maßnahmen").RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.ArtDerMaßnahme, Maßnahmen.StartUmsetzung, " & _
'                         "Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Verantwortlich, MaßnahmeUmsetzungFrist.Anzeige, " & _
'                         "Sachbearbeiter.Kurzzeichen, Maßnahmen.Fortschritt " & _
'                         "FROM (Maßnahmen LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]) " & _
'                         "LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID] " & _
'                         "GROUP BY Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.ArtDerMaßnahme, Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, " & _
'                         "Maßnahmen.UmgesetztAm, Maßnahmen.Verantwortlich, MaßnahmeUmsetzungFrist.Anzeige, Sachbearbeiter.Kurzzeichen, Maßnahmen.Fortschritt " & _
'                         "ORDER BY Maßnahmen.MaßnahmeID;"
'
'        Case 1
'            Forms("Maßnahmen").RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.ArtDerMaßnahme, Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Verantwortlich, MaßnahmeUmsetzungFrist.Anzeige, Sachbearbeiter.Kurzzeichen, LQBereich.[ID-LQEinheit], Maßnahmen.Fortschritt " & _
'                        "FROM (((Maßnahmen LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]) LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) INNER JOIN MaßnahmenFehlerortGruppe ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerortGruppe.MaßnahmeID) INNER JOIN (LQBereich INNER JOIN LQEinheit ON LQBereich.[ID-LQEinheit] = LQEinheit.ID) ON MaßnahmenFehlerortGruppe.LQBereichID = LQBereich.ID " & _
'                        "GROUP BY Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.ArtDerMaßnahme, Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Verantwortlich, MaßnahmeUmsetzungFrist.Anzeige, Sachbearbeiter.Kurzzeichen, LQBereich.[ID-LQEinheit], Maßnahmen.Fortschritt " & _
'                        "ORDER BY Maßnahmen.MaßnahmeID;"
'
'        Case 2
'            Forms("Maßnahmen").RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.ArtDerMaßnahme, Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Verantwortlich, MaßnahmeUmsetzungFrist.Anzeige, Sachbearbeiter.Kurzzeichen, MaßnahmenFehlerortGruppe.LQBereichID, LQBereich.[ID-LQEinheit], Maßnahmen.Fortschritt " & _
'                        "FROM (((Maßnahmen LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]) LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) INNER JOIN MaßnahmenFehlerortGruppe ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerortGruppe.MaßnahmeID) INNER JOIN LQBereich ON MaßnahmenFehlerortGruppe.LQBereichID = LQBereich.ID " & _
'                        "GROUP BY Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.ArtDerMaßnahme, Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Verantwortlich, MaßnahmeUmsetzungFrist.Anzeige, Sachbearbeiter.Kurzzeichen, MaßnahmenFehlerortGruppe.LQBereichID, LQBereich.[ID-LQEinheit], Maßnahmen.Fortschritt " & _
'                        "ORDER BY Maßnahmen.MaßnahmeID;"
'        Case 3
'            Forms("Maßnahmen").RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.ArtDerMaßnahme, Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Verantwortlich, MaßnahmeUmsetzungFrist.Anzeige, Sachbearbeiter.Kurzzeichen, MaßnahmenFehlerortGruppe.LQBereichID, LQBereich.[ID-LQEinheit], Maßnahmen.Fortschritt " & _
'                        "FROM (((Maßnahmen LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]) LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) INNER JOIN MaßnahmenFehlerortGruppe ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerortGruppe.MaßnahmeID) INNER JOIN (LQBereich INNER JOIN LQEinheit ON LQBereich.[ID-LQEinheit] = LQEinheit.ID) ON MaßnahmenFehlerortGruppe.LQBereichID = LQBereich.ID " & _
'                        "GROUP BY Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.ArtDerMaßnahme, Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Verantwortlich, MaßnahmeUmsetzungFrist.Anzeige, Sachbearbeiter.Kurzzeichen, MaßnahmenFehlerortGruppe.LQBereichID, LQBereich.[ID-LQEinheit], Maßnahmen.Fortschritt " & _
'                        "ORDER BY Maßnahmen.MaßnahmeID;"
'
'        Case 4
'            Forms("Maßnahmen").RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.ArtDerMaßnahme, Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Verantwortlich, MaßnahmeUmsetzungFrist.Anzeige, Sachbearbeiter.Kurzzeichen, MaßnahmenFehlerarten.Fehlerklasse, Maßnahmen.Fortschritt " & _
'                        "FROM ((Maßnahmen LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]) LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) LEFT JOIN MaßnahmenFehlerarten ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerarten.MaßnahmeID " & _
'                        "GROUP BY Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.ArtDerMaßnahme, Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Verantwortlich, MaßnahmeUmsetzungFrist.Anzeige, Sachbearbeiter.Kurzzeichen, MaßnahmenFehlerarten.Fehlerklasse, Maßnahmen.Fortschritt " & _
'                        "ORDER BY Maßnahmen.MaßnahmeID;"
'
'        Case 5
'            Forms("Maßnahmen").RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.ArtDerMaßnahme, Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Verantwortlich, MaßnahmeUmsetzungFrist.Anzeige, " & _
'                        "Sachbearbeiter.Kurzzeichen, LQBereich.[ID-LQEinheit], MaßnahmenFehlerarten.Fehlerklasse, Maßnahmen.Fortschritt " & _
'                        "FROM ((((Maßnahmen LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]) " & _
'                        "LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) " & _
'                        "INNER JOIN MaßnahmenFehlerortGruppe ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerortGruppe.MaßnahmeID) " & _
'                        "INNER JOIN (LQBereich INNER JOIN LQEinheit ON LQBereich.[ID-LQEinheit] = LQEinheit.ID) ON MaßnahmenFehlerortGruppe.LQBereichID = LQBereich.ID) " & _
'                        "LEFT JOIN MaßnahmenFehlerarten ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerarten.MaßnahmeID " & _
'                        "GROUP BY Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.ArtDerMaßnahme, Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Verantwortlich, MaßnahmeUmsetzungFrist.Anzeige, Sachbearbeiter.Kurzzeichen, " & _
'                        "MaßnahmenFehlerortGruppe.LQBereichID, LQBereich.[ID-LQEinheit], MaßnahmenFehlerarten.Fehlerklasse, Maßnahmen.Fortschritt " & _
'                        "ORDER BY Maßnahmen.MaßnahmeID;"
'        Case 6
'            Forms("Maßnahmen").RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.ArtDerMaßnahme, Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Verantwortlich, MaßnahmeUmsetzungFrist.Anzeige, Sachbearbeiter.Kurzzeichen, MaßnahmenFehlerortGruppe.LQBereichID, LQBereich.[ID-LQEinheit], MaßnahmenFehlerarten.Fehlerklasse, Maßnahmen.Fortschritt " & _
'                        "FROM ((((Maßnahmen LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]) LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) INNER JOIN MaßnahmenFehlerortGruppe ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerortGruppe.MaßnahmeID) INNER JOIN LQBereich ON MaßnahmenFehlerortGruppe.LQBereichID = LQBereich.ID) LEFT JOIN MaßnahmenFehlerarten ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerarten.MaßnahmeID " & _
'                        "GROUP BY Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.ArtDerMaßnahme, Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Verantwortlich, MaßnahmeUmsetzungFrist.Anzeige, Sachbearbeiter.Kurzzeichen, MaßnahmenFehlerortGruppe.LQBereichID, LQBereich.[ID-LQEinheit], MaßnahmenFehlerarten.Fehlerklasse, Maßnahmen.Fortschritt " & _
'                        "ORDER BY Maßnahmen.MaßnahmeID;"
'
'        Case 7
'            Forms("Maßnahmen").RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.ArtDerMaßnahme, Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Verantwortlich, MaßnahmeUmsetzungFrist.Anzeige, " & _
'                        "Sachbearbeiter.Kurzzeichen, MaßnahmenFehlerortGruppe.LQBereichID, LQBereich.[ID-LQEinheit], MaßnahmenFehlerarten.Fehlerklasse, Maßnahmen.Fortschritt " & _
'                        "FROM ((((Maßnahmen LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]) " & _
'                        "LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) " & _
'                        "INNER JOIN MaßnahmenFehlerortGruppe ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerortGruppe.MaßnahmeID) " & _
'                        "INNER JOIN (LQBereich INNER JOIN LQEinheit ON LQBereich.[ID-LQEinheit] = LQEinheit.ID) ON MaßnahmenFehlerortGruppe.LQBereichID = LQBereich.ID) " & _
'                        "LEFT JOIN MaßnahmenFehlerarten ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerarten.MaßnahmeID " & _
'                        "GROUP BY Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.ArtDerMaßnahme, Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Verantwortlich, MaßnahmeUmsetzungFrist.Anzeige, Sachbearbeiter.Kurzzeichen, " & _
'                        "MaßnahmenFehlerortGruppe.LQBereichID, LQBereich.[ID-LQEinheit], MaßnahmenFehlerarten.Fehlerklasse, Maßnahmen.Fortschritt " & _
'                        "ORDER BY Maßnahmen.MaßnahmeID;"
'
'    End Select



    Forms("Maßnahmen").Controls("FilterView") = sfilter
    Forms("Maßnahmen").Filter = sfilter
    Forms("Maßnahmen").FilterOn = True
    Forms("Maßnahmen").Requery
    Forms("Maßnahmen").Refresh
    DoCmd.Close acForm, "MaßnahmeFilter_Dialog"
    Forms("Maßnahmen").SetFocus



End Sub

Private Sub Form_Load()

Dim objMouseOverForm As clsMouseoverForm
    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With

End Sub

Private Sub Form_Open(Cancel As Integer)

    Dim XPos As Long
    Dim YPos As Long
    Dim lgPosition As Long

    lgPosition = getPointerPos(XPos, YPos)
    FormOpenKoorMousPointer Me, XPos, YPos - 200, 450, 400

    Me.CtrlInBearbeitung = False
    Me.CtrlAbgeschlossen = False


End Sub
