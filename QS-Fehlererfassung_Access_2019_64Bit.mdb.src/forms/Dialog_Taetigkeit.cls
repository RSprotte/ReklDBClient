Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database

Private Sub Abbrechen_Click()

DoCmd.Close

End Sub

Private Sub Form_Close()

    FormKoorSave Me

End Sub

Private Sub Form_Load()

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

Private Sub OK_Click()
On Error GoTo err_OK_Click

    Dim ldb As DAO.Database
    Dim rsTaetigkeit As DAO.Recordset
    Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider

    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider
    'Daten aus Access auslesen

    Set ldb = openAccDBConnection("openDB")
    If Me.Taetigkeit <> "" And Not IsNull(Me.Taetigkeit) Then
        If DialogTaetigkeit = "Neu" Then
                'Neuanlage der Tätigkeit, da Recordcount bei Abfrage in Tabelle = 0
                ldb.Execute ("Insert into Taetigkeiten (Taetigkeit, Beschreibung, Zuordnung_ID, Anzeige) " & _
                            "values ('" & Me.Taetigkeit & "', '" & Me.Beschreibung & "', " & Me.CboZuordnungID & ", '" & Me.Taetigkeit & "');")
                Me.Requery

            ' Neuer Eintrag in Avista
            Set rsTaetigkeit = ldb.OpenRecordset("Select * from Taetigkeiten ;")
                rsTaetigkeit.MoveLast
                    If rsTaetigkeit("Zuordnung_ID") = 3 Or rsTaetigkeit("Zuordnung_ID") = 5 Then
                        adbprov.CommandExecuteAvista ("Insert into zlqpp_fe_taetigkeit ([fet_id],[fet_name]) Values (" & rsTaetigkeit("ID") & ", '" & rsTaetigkeit("Taetigkeit") & "');")
                    End If

        Else
            Set rsTaetigkeit = ldb.OpenRecordset("Select * from Taetigkeiten where ID=" & TaetigkeitID & ";")
            If Not rsTaetigkeit.RecordCount = 0 Then
                With Forms("Taetigkeiten")
                    'Sollte eigentlich nicht passieren, da Recordcount = 0 -> Tätigkeit noch nicht vorhanden, also auch kein Update notwendig
                    ldb.Execute ("UPDATE Taetigkeiten set Taetigkeit='" & Me.Taetigkeit & "', Beschreibung='" & Me.Beschreibung & "', Zuordnung_ID=" & Me.CboZuordnungID & ", Anzeige='" & Me.Taetigkeit & "'" & _
                                " WHERE ID =" & TaetigkeitID & ";")

                    adbprov.CommandExecuteAvista ("Update zlqpp_fe_taetigkeit set [fet_name] = '" & Me.Taetigkeit & "' where [fet_id] = " & TaetigkeitID & ";")

                    .Requery
                    .Refresh
                End With
            End If

        End If
            DoCmd.Close acForm, "Dialog_Taetigkeit"
            Forms("Taetigkeiten").SetFocus
    Else
        MsgBox getTC_FCFromCode("TC0335", ldb)  'MsgBox ("Daten nicht vollständig!")   'TC0336
    End If

    Set adbprov = Nothing
    Set rsTaetigkeit = Nothing

    Me.Refresh

    Exit Sub
err_OK_Click:

    If Err = xxxx Then
        MsgBox getTC_FCFromCode("TC0334", ldb)  'MsgBox ("Tätigkeit schon vorhanden!")  'TC0334
        Set db = Nothing
    End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub
