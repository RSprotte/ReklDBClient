Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Befehl31_Click()

    Dim ldb As DAO.Database
    Dim rs As Recordset
    Dim MID As Integer



    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select MaßnahmenTempDruck.MaßnahmeID from MaßnahmenTempDruck where MaßnahmenTempDruck.Druck = True;")
        Do While Not rs.EOF
            MID = rs![MaßnahmeID]
            ldb.Execute ("Insert into MIDZuReklID ([MID],[ReklID]) values (" & MID & ", " & Forms("Fehlererfassung - Register").Controls("Mastertab#ID") & ");")
            rs.MoveNext
        Loop

        Set rs = Nothing

    DoCmd.Save

    OhneXDReport 'Füllt das Control mit allen zugeordneten Maßnahmen der Reklamation
    Set rs = getCtrlDataset(Forms("Fehlererfassung - Register"), Forms("Fehlererfassung - Register").Controls("Mastertab#langfristige Maßnahmen"))

    DoCmd.Close acForm, Me.Name

'    Select Case psMassnahmeBearbeitung
'        Case "XD-Report"
'            XDReport
'            OhneXDReport
'        Case "Ohne XD-Report"
'            OhneXDReport
'    End Select
    Set rs = Nothing



End Sub

Sub XDReport()
On Error GoTo err_XDReport

    Dim ldb As DAO.Database
    Dim rs, rsKunde As DAO.Recordset
    Dim fs As Object
    Dim wb As Excel.Workbook
    Dim WS As Excel.Worksheet
    Dim lexcelprov As clsExcelProvider
    Dim sPfadDatei As String

    Dim XDReportPfad As String  '20160304-ROS
    XDReportPfad = getTC_FCFromCode("DP0001")

    Dim XDReportTemplate As String  '20160304-ROS
    XDReportTemplate = getTC_FCFromCode("DP0002")

    Dim XDReportTemplateLieferant As String  '20160304-ROS
    XDReportTemplateLieferant = getTC_FCFromCode("DP0003")



    With Forms("Fehlererfassung - Register")
        'Erzeugen einer Excel Instanz
        Set lexcelprov = New clsExcelProvider
        'Prüfen ob XD-Report schon existiert
        Set fs = CreateObject("Scripting.FileSystemObject")
        'ist Verzeichnis vorhanden
        sPfadDatei = getReklVerzeichnis(.Controls("Mastertab#ID"), format(.Controls("Mastertab#Auftragsdatum#3"), "yyyy"), .Controls("Mastertab#Meldungsnummer Kunde#2"))
        If sPfadDatei <> "" Then
            If fs.FileExists(sPfadDatei & "\" & .ID & "_QS-XD-Report.xlsx") = False Then
                'Report generieren
                '1. Kopie von Template erstellen
                Set wb = lexcelprov.openExcelInstance(XDReportPfad & XDReportTemplate, sPfadDatei & "\" & .ID & "_QS-XD-Report.xlsx")
                '2. XD-Report generieren
                Set WS = wb.Worksheets("Report")
                Set ldb = openAccDBConnection("openDB")
                Set rs = ldb.OpenRecordset("Select * from RSP_Ausw_Basisdaten_1 where [Mastertab].[ID] = " & .ID & ";")
                If rs.RecordCount > 0 Then
                    WS.Range("Datum") = format(rs("Auftragsdatum"))
                    WS.Range("Meldungsnummer") = rs("Meldungsnummer Kunde")
                    WS.Range("Bestellnummer") = rs("Bestellnummer Kunde")

                    WS.Range("Kunde") = rs("Firma")
                    Set rsKunde = ldb.OpenRecordset("Select * from Kunden where [Kunden-ID] = " & .Kunden_ID & ";")
                    If rsKunde.RecordCount > 0 Then
                        WS.Range("AnsprechpartnerKunde") = rsKunde("Kontaktperson")
                        WS.Range("AnsprechpartnerKundeMail") = rsKunde("EmailAdresse")
                        WS.Range("AnsprechpartnerKundeTel") = rsKunde("Telefonnummer")
                    Else
                    End If
                    Set rsKunde = Nothing
                    WS.Range("Maschinennummer") = rs("Maschinennummer")
                    WS.Range("TRUMPF_Baugruppe") = rs("BG-TRUMPF")
                    WS.Range("TRUMPF_Artikelnummer") = rs("ArtNr-TRUMPF")
                    WS.Range("LQ_Baugruppe") = rs("BG")
                    WS.Range("LQ_Artikelnummer") = rs("ArtNr")
                    WS.Range("Fehlermenge") = rs("Menge betroffen - Kunde")
                    'WS.Range("LQ_SBEmail") = rs("Email-Name")
                    'WS.Range("LQ_SBTelefon") = rs("Telefon/beruflich")
                    'WS.Range("Fehlermenge") = rs("Menge betroffen - Kunde")
                    WS.Range("ReklNr") = rs("ID")
                    WS.Range("Bearbeitungsdatum") = rs("Rückmeldung_Kunde_Verschiebung")
        '            WS.Range("Datum") = Date

                    Select Case rs("Verursacher")
                        Case "LQ"
                            WS.Range("FehlerLQ") = "X"
                        Case "LQUS"
                            WS.Range("Fehler_LQUS") = "X"
                        Case "LQCN"
                            WS.Range("Fehler_LQCN") = "X"
                        Case "Lieferant"
                            WS.Range("FehlerLieferant") = "X"
                        Case "Kunde"
                    End Select
                    WS.Range("Fehlerbeschreibung") = rs("Welcher Fehler tritt auf")
                    WS.Rows(WS.Range("Fehlerbeschreibung").Row & ":" & WS.Range("Fehlerbeschreibung").Row).Select
                    Selection.Rows.AutoFit
                    WS.Range("Fehlerursache") = rs("Wie Warum passierte der Fehler")
                    WS.Rows(WS.Range("Fehlerursache").Row & ":" & WS.Range("Fehlerursache").Row).Select
                    Selection.Rows.AutoFit


                    'me.UFMaßnahmen.Form.RecordsetClone.MoveLast
                    If Me.Form.Recordset.RecordCount > 0 Then
                        Me.Form.Recordset.MoveFirst
                    'If !UFMaßnahmen.Form.Recordset.RecordCount > 0 Then
                        Dim A, B, C, i, i2 As Integer
                    '    !UFMaßnahmen.Form.Recordset.MoveFirst
                    '    For i = 1 To !UFMaßnahmen.Form.Recordset.RecordCount
                        i = 1
                        While Not Me.Form.Recordset.EOF
                        'For i = 1 To Me.Form.Recordset.RecordCount
                            If Me.[Druck] Then
                                A = 24
                                C = 24
                                If i > 1 Then
                                        WS.Rows(C & ":" & C).Select
                                        Selection.Copy
                                        WS.Rows(C + (i - 1) & ":" & C + (i - 1)).Select
                                        Selection.Insert Shift:=xlDown
                                        'wb.Names.Add Name:="Massnahme" & i,
                                        wb.Names.Add Name:="Massnahme" & i, RefersTo:="=Report!$B$" & C + (i - 1)
                                        wb.Names.Add Name:="Verantwortlich" & i, RefersTo:="=Report!$G$" & C + (i - 1)
                                        wb.Names.Add Name:="MGeplant" & i, RefersTo:="=Report!$I$" & C + (i - 1)
                                        wb.Names.Add Name:="MErledigt" & i, RefersTo:="=Report!$J$" & C + (i - 1)
                                End If

                                WS.Range("Massnahme" & LTrim(Str(i))) = "Maßnahme:" & Chr(13) & Chr(10) & _
                                                                        delCRLRandMinus(Me.Controls("Massnahme")) & Chr(13) & Chr(10) & _
                                                                        "Beschreibung:" & Chr(13) & Chr(10) & _
                                                                        delCRLRandMinus(Me.Controls("Beschreibung"))
                                Dim rsSB As DAO.Recordset
                                If Me.Controls("Verantwortlich") <> 0 Then
                                    Set rsSB = ldb.OpenRecordset("Select * from Sachbearbeiter where [SB-ID] = " & Me.Controls("Verantwortlich") & ";", dbOpenSnapshot)
                                    WS.Range("Verantwortlich" & LTrim(Str(i))) = rsSB("Name")
                                    Set rsSB = Nothing
                                Else
                                    WS.Range("Verantwortlich" & LTrim(Str(i))) = "NN"
                                End If

                                WS.Range("MGeplant" & LTrim(Str(i))) = Me.Controls("StartUmsetzung")
                                WS.Range("MErledigt" & LTrim(Str(i))) = Me.Controls("EndeUmsetzung")
                                WS.Range("Massnahme" & LTrim(Str(i))).Select
                                WS.Rows(C + (i - 1) & ":" & C + (i - 1)).Select
'                                WS.Range("Massnahme" & LTrim(Str(i))) = "Maßnahme:" & Chr(13) & Chr(10) & _
'                                                                        delCRLRandMinus(!UFMaßnahmen.Controls("Massnahme")) & Chr(13) & Chr(10) & _
'                                                                        "Beschreibung:" & Chr(13) & Chr(10) & _
'                                                                        delCRLRandMinus(!UFMaßnahmen.Controls("Beschreibung"))
'                                WS.Range("Verantwortlich" & LTrim(Str(i))) = !UFMaßnahmen.Controls("Verantwortlich")
'                                WS.Range("MGeplant" & LTrim(Str(i))) = !UFMaßnahmen.Controls("StartUmsetzung")
'                                WS.Range("MErledigt" & LTrim(Str(i))) = !UFMaßnahmen.Controls("EndeUmsetzung")
'                                WS.Range("Massnahme" & LTrim(Str(i))).Select
'                                WS.Rows(c + (i - 1) & ":" & c + (i - 1)).Select
                                Selection.Rows.AutoFit

                                i = i + 1
                            End If
                            'Next i
                            Me.Form.Recordset.MoveNext
                        Wend
                        'Next i

                        'For i2 = 29 To (29 - me.UFMaßnahmen.Form.Recordset.RecordCount) + 1 Step -1
'                        For i2 = 29 To (29 - (6 - i)) Step -1
'                            WS.Rows(i2 & ":" & i2).Select
'                            Selection.Delete Shift:=xlUp
'                        Next i2
                    End If
                    'Eintrag QSV Sachbearbeiter
                    Set rs = Nothing
                    Set rs = ldb.OpenRecordset("Select * from Sachbearbeiter where [Name] = '" & .[Kzz_Sachbearbeiter] & "';")
                    If rs.RecordCount > 0 Then
                        WS.Range("QSVName") = rs("Name")
                        WS.Range("QSVMail") = rs("EmailName")
                        WS.Range("QSVTel") = rs("TelefonBeruflich")
                        WS.Range("QSVDatum") = Date
                    Else

                    End If
                    Set WS = Nothing
                    Set rs = Nothing
                     ' Set ldb = openAccDBConnection("closeDB")
                End If

                '...............
            End If
            'open XD-Report
        Else
'            MsgBox ("Verzeichnis Vorgang " & .ID & " nicht vorhanden!")      'TC0312
            MsgBox (getTC_FCFromCode("TC0312") & " " & .ID & " " & getTC_FCFromCode("TC0405"))
            'Instanze des bestehenden Report erzeugen
            'Set wb = lexcelprov.openExcelInstance(lPfad & "\" & Me.ID & "_QS-XD-Report.xlsx")
        End If
        'wb.Activate
        Set fs = Nothing
        lexcelprov.closeExcelInstance
    End With

Exit Sub

err_XDReport:

        MsgBox (getTC_FCFromCode("TC0313") & Error)    'TC0313


End Sub

Sub OhneXDReport()
'Prozedur füllt alle zugeordnete Maßnahme als Klartext in das Control des Fehlererfassung - Register
On Error GoTo Err_OhneXDReport

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Set ldb = openAccDBConnection("openDB")
    'Set rs = ldb.OpenRecordset("Select * from MaßnahmenTempDruck where [Druck] = True;")
    Set rs = ldb.OpenRecordset("SELECT MaßnahmenTempDruck.* FROM MaßnahmenTempDruck WHERE (((MaßnahmenTempDruck.Druck)=True));")
    Forms("Fehlererfassung - Register").Controls("Mastertab#langfristige Maßnahmen").Value = " "
    If rs.RecordCount Then
        rs.MoveFirst
        While Not rs.EOF
            'MsgBox (delCRLRandMinus(rs("Beschreibung").Value))
            With Forms("Fehlererfassung - Register")
                .Controls("Mastertab#langfristige Maßnahmen") = .Controls("Mastertab#langfristige Maßnahmen") & "MID-" & CStr(rs("MaßnahmeID").Value) & Chr(13) & Chr(10)
                .Controls("Mastertab#langfristige Maßnahmen") = .Controls("Mastertab#langfristige Maßnahmen") & "Maßnahme:" & Chr(13) & Chr(10)
                .Controls("Mastertab#langfristige Maßnahmen") = .Controls("Mastertab#langfristige Maßnahmen") & "-----------------" & Chr(13) & Chr(10)
                .Controls("Mastertab#langfristige Maßnahmen") = .Controls("Mastertab#langfristige Maßnahmen") & delCRLRandMinus(rs("Maßnahme").Value) & Chr(13) & Chr(10)
                .Controls("Mastertab#langfristige Maßnahmen") = .Controls("Mastertab#langfristige Maßnahmen") & "Beschreibung:" & Chr(13) & Chr(10)
                .Controls("Mastertab#langfristige Maßnahmen") = .Controls("Mastertab#langfristige Maßnahmen") & "---------------------" & Chr(13) & Chr(10)
                .Controls("Mastertab#langfristige Maßnahmen") = .Controls("Mastertab#langfristige Maßnahmen") & delCRLRandMinus(rs("Beschreibung").Value) & Chr(13) & Chr(10)
                .Controls("Mastertab#langfristige Maßnahmen") = .Controls("Mastertab#langfristige Maßnahmen") & "Start Umsetzung:" & rs("StartUmsetzung").Value
                .Controls("Mastertab#langfristige Maßnahmen") = .Controls("Mastertab#langfristige Maßnahmen") & "      Ende Umsetzung:" & rs("EndeUmsetzung").Value & Chr(13) & Chr(10)
                .Controls("Mastertab#langfristige Maßnahmen") = .Controls("Mastertab#langfristige Maßnahmen") & "----------------------------------------------------------------------------------------------------------" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
            End With
            rs.MoveNext
        Wend
    End If
    Set rs = Nothing
     ' Set ldb = openAccDBConnection("closeDB")

    Exit Sub
Exit_OhneXDReport:
    Set rs = Nothing
     ' Set ldb = openAccDBConnection("closeDB")

    DoCmd.Close
    Exit Sub

Err_OhneXDReport:
    MsgBox Err.Description
    Resume Exit_OhneXDReport



End Sub

Private Sub cmbSetDruck_Click()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Set ldb = openAccDBConnection("openDB")

    Set rs = ldb.OpenRecordset("Select * from MaßnahmenTempDruck where MaßnahmeID = " & Me.MaßnahmeID & ";")
    If rs.RecordCount Then
        rs.Edit
        If rs("Druck") = True Then
            rs("Druck") = False
        Else
            rs("Druck") = True
        End If
        rs.Update
    End If
    Set rs = Nothing

    Me.Requery
    Me.Refresh


End Sub

Private Sub Form_Close()

    FormKoorSave Me

End Sub

Private Sub Form_DblClick(Cancel As Integer)

    DoCmd.OpenForm "MaßnahmenDialog", acNormal
    Forms("MaßnahmenDialog").Filter = "[Maßnahmen].[MaßnahmeID] = " & Me.MaßnahmeID
    Forms("MaßnahmenDialog").FilterOn = True
    Forms("MaßnahmenDialog").Requery
    Forms("MaßnahmenDialog").Refresh
    Forms("MaßnahmenDialog").AllowAdditions = False

End Sub

Private Sub Form_Load()

    FormOpenKoor Me

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
Private Sub Befehl32_Click()
On Error GoTo Err_Befehl32_Click


    Screen.PreviousControl.SetFocus
    DoCmd.RunCommand acCmdFind

Exit_Befehl32_Click:
    Exit Sub

Err_Befehl32_Click:
    MsgBox Err.Description
    Resume Exit_Befehl32_Click

End Sub
