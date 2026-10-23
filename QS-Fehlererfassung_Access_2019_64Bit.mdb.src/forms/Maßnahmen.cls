Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub AuswTaetigkeit_AfterUpdate()
On Error GoTo err_AuswFehlerklasse_AfterUpdate
'RSP - 20131110 - Änderung
'    Me.RecordSource = "SELECT Maßnahmen.MaßnahmeID, MaßnahmenFehlerarten.Fehlerklasse, Sachbearbeiter.Kurzzeichen, Maßnahmen.Maßnahme, " & _
'                    "Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, " & _
'                    "Maßnahmen.StatusUmsetzung, Maßnahmen.ArtDerMaßnahme, Maßnahmen.LinkDokument, Maßnahmen.Verantwortlich, Maßnahmen.Fortschritt, Maßnahmen.Ursprung, " & _
'                    "Maßnahmen.BearbeitungsOrt " & _
'                    "FROM (Maßnahmen LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) " & _
'                    "INNER JOIN MaßnahmenFehlerarten ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerarten.MaßnahmeID " & _
'                    "ORDER BY Maßnahmen.MaßnahmeID;"

    Me.RecordSource = "SELECT Maßnahmen.UmgesetztAm, Maßnahmen.MaßnahmeID, MaßnahmenFehlerarten.Fehlerklasse, Sachbearbeiter.Kurzzeichen, Maßnahmen.Maßnahme, " & _
                    "Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, " & _
                    "Maßnahmen.StatusUmsetzung, Maßnahmen.ArtDerMaßnahme, Maßnahmen.LinkDokument, Maßnahmen.Verantwortlich, Maßnahmen.Fortschritt, Maßnahmen.Ursprung, " & _
                    "Maßnahmen.BearbeitungsOrt, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige " & _
                    "FROM ((Maßnahmen LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) INNER JOIN MaßnahmenFehlerarten ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerarten.MaßnahmeID) " & _
                    "LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme] " & _
                    "ORDER BY Maßnahmen.MaßnahmeID;"


    Me.Filter = "[Fehlerklasse] = '" & Me.AuswTaetigkeit & "'"
    Me.FilterOn = True
    Me.Requery
    Me.Refresh
    Exit Sub

err_AuswFehlerklasse_AfterUpdate:
    MsgBox ("Fehler (" & Err & ") - " & Error)
End Sub


Private Sub cmbEcxelExport_Click()

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

Private Sub DatensatzÜbernehmen_Click()
    Dim ldb As DAO.Database
    Dim rs  As DAO.Recordset
    'Maßnahme der Fehlerart zuordnen
    s_Fehlerklasse = Forms("Fehlererfassung - Register").Controls("Kategorie-Fehleranalyse")
     Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from MaßnahmenFehlerarten where [Fehlerklasse] = '" & s_Fehlerklasse & "' and MaßnahmeID = " & Me.MID & ";")
    If rs.RecordCount = 0 Then
        ldb.Execute ("Insert into MaßnahmenFehlerarten (Fehlerklasse, MaßnahmeID) Values ('" & s_Fehlerklasse & "', " & Me.MID & ");")
    Else
'        MsgBox "Maßnahme ist der Fehlerklasse: " & s_Fehlerklasse & " schon zugeordnet"  'TC0231
        MsgBox getTC_FCFromCode("TC0231", ldb) & " " & s_Fehlerklasse & " " & getTC_FCFromCode("TC0396", ldb)
    End If

    Set rs = Nothing
     ' Set ldb = openAccDBConnection("closeDB")

    DoCmd.Close
    Forms("Fehlererfassung - Register").Refresh

End Sub


Private Sub FilterLoeschen_Click()
'RSP - 20131110 - Änderung
'    Me.RecordSource = "SELECT Maßnahmen.MaßnahmeID, Sachbearbeiter.Kurzzeichen, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], " & _
'                    "Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.StatusUmsetzung, Maßnahmen.ArtDerMaßnahme, " & _
'                    "Maßnahmen.LinkDokument, Maßnahmen.Verantwortlich, Maßnahmen.Fortschritt, Maßnahmen.Ursprung, Maßnahmen.BearbeitungsOrt " & _
'                    "FROM Maßnahmen LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID] " & _
'                    "ORDER BY Maßnahmen.MaßnahmeID;"


'        Me.RecordSource = "SELECT Maßnahmen.UmgesetztAm, Maßnahmen.MaßnahmeID, Sachbearbeiter.Kurzzeichen, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], " & _
'                "Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.StatusUmsetzung, Maßnahmen.ArtDerMaßnahme, Maßnahmen.LinkDokument, " & _
'                "Maßnahmen.Fortschritt, Maßnahmen.Ursprung, Maßnahmen.BearbeitungsOrt, MaßnahmeUmsetzungFrist.Anzeige, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme] " & _
'                "FROM (Maßnahmen LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]) " & _
'                "LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID] " & _
'                "ORDER BY Maßnahmen.MaßnahmeID;"

'     Me.RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], " & _
'                        "Maßnahmen.StartUmsetzung , Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, " & _
'                        "Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, LQBereich.[ID-LQEinheit], MaßnahmenFehlerortGruppe.LQBereichID, MaßnahmenFehlerarten.Fehlerklasse, " & _
'                        "Maßnahmen.ArtDerMaßnahme , MaßnahmeUmsetzungFrist.Anzeige " & _
'                        "FROM (((((Maßnahmen LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID])" & _
'                        "LEFT JOIN MaßnahmenHerkunft ON Maßnahmen.MaßnahmenHerkunftID = MaßnahmenHerkunft.ID)" & _
'                        "INNER JOIN MaßnahmenFehlerortGruppe ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerortGruppe.MaßnahmeID)" & _
'                        "INNER JOIN LQBereich ON MaßnahmenFehlerortGruppe.LQBereichID = LQBereich.ID)" & _
'                        "INNER JOIN MaßnahmenFehlerarten ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerarten.MaßnahmeID)" & _
'                        "INNER JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]" & _
'                        "ORDER BY Maßnahmen.MaßnahmeID;"


        Me.RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige " & _
                        "FROM ((Maßnahmen LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) LEFT JOIN MaßnahmenHerkunft ON Maßnahmen.MaßnahmenHerkunftID = MaßnahmenHerkunft.ID) LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme] " & _
                        "GROUP BY Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige " & _
                        "ORDER BY Maßnahmen.MaßnahmeID;"



    Me.Filter = ""
    Me.FilterOn = True
    Me.AuswTaetigkeit = ""
    Me.FilterView = ""
    Me.Requery
    Me.Refresh

End Sub

Private Sub FilterSetzen_Click()

DoCmd.OpenForm "MaßnahmeFilter_Dialog"

End Sub

Private Sub Form_Close()

On Error GoTo err_Form_Close

    DoCmd.Close acForm, "MaßnahmeFilter_Dialog"

err_Form_Close:
    Exit Sub



End Sub

Private Sub Form_DblClick(Cancel As Integer)
On Error GoTo Err_Form_DblClick


    DoCmd.OpenForm "MaßnahmenDialog", acNormal, , , acFormAdd
    Forms("MaßnahmenDialog").Filter = "[Maßnahmen].[MaßnahmeID] = " & Me.MID
    Forms("MaßnahmenDialog").FilterOn = True
    Forms("MaßnahmenDialog").Requery
    Forms("MaßnahmenDialog").Refresh
    Forms("MaßnahmenDialog").AllowAdditions = False

Exit_Form_DblClick:
    Exit Sub

Err_Form_DblClick:
    MsgBox Err.Description
    Resume Exit_Form_DblClick

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

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub

Private Sub Form_Open(Cancel As Integer)

'    Me.RecordSource = "SELECT Maßnahmen.UmgesetztAm, Maßnahmen.MaßnahmeID, Sachbearbeiter.Kurzzeichen, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], " & _
'                        "Maßnahmen.StartUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.EndeUmsetzung, Maßnahmen.StatusUmsetzung, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.Anzeige, Maßnahmen.Priorität, MaßnahmenHerkunft.ID, MaßnahmenHerkunft.Anzeige,  Maßnahmen.Verantwortlich, Sachbearbeiter.[SB-ID] " & _
'                        "FROM ((Maßnahmen LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]) " & _
'                        "LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) LEFT JOIN MaßnahmenHerkunft ON Maßnahmen.MaßnahmenHerkunftID = MaßnahmenHerkunft.ID " & _
'                        "ORDER BY Maßnahmen.MaßnahmeID;"

'    Me.RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], " & _
'                        "Maßnahmen.StartUmsetzung , Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, " & _
'                        "Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, LQBereich.[ID-LQEinheit], MaßnahmenFehlerortGruppe.LQBereichID, MaßnahmenFehlerarten.Fehlerklasse, " & _
'                        "Maßnahmen.ArtDerMaßnahme , MaßnahmeUmsetzungFrist.Anzeige " & _
'                        "FROM (((((Maßnahmen LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID])" & _
'                        "LEFT JOIN MaßnahmenHerkunft ON Maßnahmen.MaßnahmenHerkunftID = MaßnahmenHerkunft.ID)" & _
'                        "INNER JOIN MaßnahmenFehlerortGruppe ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerortGruppe.MaßnahmeID)" & _
'                        "INNER JOIN LQBereich ON MaßnahmenFehlerortGruppe.LQBereichID = LQBereich.ID)" & _
'                        "INNER JOIN MaßnahmenFehlerarten ON Maßnahmen.MaßnahmeID = MaßnahmenFehlerarten.MaßnahmeID)" & _
'                        "INNER JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme]" & _
'                        "ORDER BY Maßnahmen.MaßnahmeID;"

    Me.RecordSource = "SELECT Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige " & _
                        "FROM ((Maßnahmen LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]) LEFT JOIN MaßnahmenHerkunft ON Maßnahmen.MaßnahmenHerkunftID = MaßnahmenHerkunft.ID) LEFT JOIN MaßnahmeUmsetzungFrist ON Maßnahmen.ArtDerMaßnahme = MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme] " & _
                        "GROUP BY Maßnahmen.MaßnahmeID, Maßnahmen.Maßnahme, Maßnahmen.Beschreibung, Maßnahmen.[Ziel der Maßnahme], Maßnahmen.[Ergebnis Maßnahme], Maßnahmen.StartUmsetzung, Maßnahmen.EndeUmsetzung, Maßnahmen.UmgesetztAm, Maßnahmen.Priorität, Maßnahmen.MaßnahmenHerkunftID, MaßnahmenHerkunft.Anzeige, Maßnahmen.Verantwortlich, Sachbearbeiter.Kurzzeichen, Maßnahmen.ArtDerMaßnahme, MaßnahmeUmsetzungFrist.[ID-UmsetzungMaßnahme], MaßnahmeUmsetzungFrist.Anzeige " & _
                        "ORDER BY Maßnahmen.MaßnahmeID DESC;"



    Me.Filter = ""
    Me.FilterOn = True
    Me.AuswTaetigkeit = ""
    Me.Requery
    Me.Refresh

End Sub

Private Sub NewDataset_Click()
On Error GoTo Err_NewDataset_Click

    Dim ldb As DAO.Database
    Dim rs  As DAO.Recordset
    Dim lIDM As Long

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Maßnahmen", dbOpenDynaset)

    rs.AddNew
    rs("Maßnahme") = "###"
    rs.Update
    rs.MoveFirst
    rs.MoveLast
    lIDM = rs("MaßnahmeID")


    DoCmd.OpenForm "MaßnahmenDialog", acNormal, , , acFormAdd
'    DoCmd.GoToRecord , , acNewRec
    '##############################################################
    Forms("MaßnahmenDialog").Filter = "[Maßnahmen].[MaßnahmeID] = " & lIDM
    Forms("MaßnahmenDialog").FilterOn = True
    Forms("MaßnahmenDialog").Controls("StartUmsetzung").Value = Date
    Forms("MaßnahmenDialog").Controls("EndeUmsetzung").Value = Date
    '##############################################################
    Forms("MaßnahmenDialog").Controls("cboPriorität").Value = 1
    Forms("MaßnahmenDialog").Controls("cboMaßnahmenHerkunft").Value = 5
'    DoCmd.RunCommand acCmdSaveRecord

    Set rs = Nothing


Exit_NewDataset_Click:
    Exit Sub

Err_NewDataset_Click:
    MsgBox Err.Description
    Resume Exit_NewDataset_Click

End Sub
Private Sub DatensatzBearbeiten_Click()
On Error GoTo Err_DatensatzBearbeiten_Click


    DoCmd.OpenForm "MaßnahmenDialog", acNormal, , , acFormAdd
    Forms("MaßnahmenDialog").Filter = "[Maßnahmen].[MaßnahmeID] = " & Me.MID
    Forms("MaßnahmenDialog").FilterOn = True
    Forms("MaßnahmenDialog").Requery
    Forms("MaßnahmenDialog").Refresh
    Forms("MaßnahmenDialog").AllowAdditions = False





Exit_DatensatzBearbeiten_Click:
    Exit Sub

Err_DatensatzBearbeiten_Click:
    MsgBox Err.Description
    Resume Exit_DatensatzBearbeiten_Click

End Sub

Private Sub ObjektSchließen_Click()

    FormKoorSave Me
    DoCmd.Close

End Sub
Private Sub Befehl15_Click()
On Error GoTo Err_Befehl15_Click


    Screen.PreviousControl.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_Befehl15_Click:
    Exit Sub

Err_Befehl15_Click:
    MsgBox Err.Description
    Resume Exit_Befehl15_Click

End Sub
Private Sub BerichtMassnahmenDrucken_Click()
On Error GoTo Err_BerichtMassnahmenDrucken_Click

    Dim stDocName As String

    stDocName = "MaßnahmenBericht"
    DoCmd.OpenReport stDocName, acViewDesign, , , acHidden
    'DoCmd.OpenReport stDocName, acViewPreview, , , acWindowNormal

    If IsNull(Me.AuswTaetigkeit) Or Me.AuswTaetigkeit = "" Then
        Reports(stDocName).RecordSource = MassnahmeReport
    Else
        Reports(stDocName).RecordSource = MassnahmeMitKlasseReport
    End If
    Reports(stDocName).Filter = Me.Filter
    Reports(stDocName).FilterOn = True
    'Reports(stDocName).ServerFilter
    DoCmd.Close acReport, stDocName, acSaveYes
    'Reports(stDocName).Requery
    DoCmd.OpenReport stDocName, acViewPreview, , , acWindowNormal
    DoCmd.Maximize


Exit_BerichtMassnahmenDrucken_Click:
    Exit Sub

Err_BerichtMassnahmenDrucken_Click:
    MsgBox Err.Description
    Resume Exit_BerichtMassnahmenDrucken_Click

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub
