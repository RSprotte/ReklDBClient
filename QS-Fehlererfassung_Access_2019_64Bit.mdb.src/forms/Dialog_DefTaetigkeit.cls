Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database

Private Sub Abbrechen_Click()

    FormKoorSave Me
    DoCmd.Close

End Sub

Private Sub Form_Load()
'20250710-ROS

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

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

Private Sub OK_Click()
On Error GoTo err_OK_Click

Dim ldb As DAO.Database
Dim rsNewFehlerbildID As DAO.Recordset
Dim rs As DAO.Recordset
Dim rs2 As DAO.Recordset
Dim l_FehlerbildID As Long

 Set ldb = openAccDBConnection("openDB")
 If Me.DefTaetigkeit <> "" Then
    With Forms("Fehlerbilder")
        If DialogDefTaetigkeit = "Neu" Then
            'Neues Fehlerbild zur Tätigkeit anlegen
            ldb.Execute ("Insert into ZuordnungDefTaetigkeit (TaetigkeitID, Beschreibung, DefTaetigkeit, Anzeige) " & _
                        "values (" & TaetigkeitID & ", '" & Me.Beschreibung & "', '" & Me.DefTaetigkeit & "', '" & Me.DefTaetigkeit & "');")
            'Ermitteln ID des neuen Fehlerbildes
            Set rsNewFehlerbildID = ldb.OpenRecordset("SELECT Max(ZuordnungDefTaetigkeit.ID) AS MaxvonID FROM ZuordnungDefTaetigkeit;")
            'ID des neu angelegten Fehlerbild sichern
            l_FehlerbildID = rsNewFehlerbildID("MaxvonID")


            Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider

            'Avista DB Provider Instanz erzeugen
            Set adbprov = New clsAvistaProvider

            adbprov.CommandExecuteAvista ("Insert into zlqpp_fe_fehler_info([fef_id],[fet_id],[fef_text]) Values (" & l_FehlerbildID & ", " & TaetigkeitID & ",'" & Me.DefTaetigkeit & "');")

            Set adbprov = Nothing

            '######## Hier werden alle Fehlerbilder den WS zugeordnet die als Verursacher oder Melder die Tätigkeit zugeordnet haben  ################
            'Ermittelung aller WS die die übergeordnete Tätigkeit als Verursacher zugeordnet haben
            Set rs = ldb.OpenRecordset("SELECT ZuordnungWsEinzeltaetigkeit.WsID " & _
                                      "FROM ZuordnungDefTaetigkeit INNER JOIN ZuordnungWsEinzeltaetigkeit ON ZuordnungDefTaetigkeit.ID = ZuordnungWsEinzeltaetigkeit.EinzelTaetigkeitID " & _
                                      "GROUP BY ZuordnungWsEinzeltaetigkeit.WsID, ZuordnungDefTaetigkeit.TaetigkeitID " & _
                                      "HAVING (((ZuordnungDefTaetigkeit.TaetigkeitID)=" & TaetigkeitID & "));")
            rs.MoveFirst
            'Über alle WS
            While (Not rs.EOF)
                'Neues Fehlerbild einfügen und aktivieren
                ldb.Execute ("INSERT INTO ZuordnungWsEinzeltaetigkeit ( WsID, EinzelTaetigkeitID, Status ) " & _
                           "values (" & rs("WsID") & ", " & l_FehlerbildID & ", True);")
                rs.MoveNext
            Wend
            'ermitteln aller WS die die übergeordnete Tätigkeit als Melder zugeordnet haben
            Set rs = ldb.OpenRecordset("SELECT ZuordnungWsEinzeltaetigkeit.WsID " & _
                                      "FROM ZuordnungDefTaetigkeit INNER JOIN ZuordnungWsEinzeltaetigkeit ON ZuordnungDefTaetigkeit.ID = ZuordnungWsEinzeltaetigkeit.EinzelTaetigkeitMelderID " & _
                                      "GROUP BY ZuordnungWsEinzeltaetigkeit.WsID, ZuordnungDefTaetigkeit.TaetigkeitID " & _
                                      "HAVING (((ZuordnungDefTaetigkeit.TaetigkeitID)=" & TaetigkeitID & "));")
            rs.MoveFirst
            'Über alle WS
            While (Not rs.EOF)
                'Wenn der WS schon die Tätigkeit als Verursacher zugeordnet wurde, kann das Fehlerbild als Melder nur im bestehend DS ergänzt werden
                Set rs2 = ldb.OpenRecordset("SELECT ZuordnungWsEinzeltaetigkeit.StatusMelder, ZuordnungWsEinzeltaetigkeit.EinzelTaetigkeitMelderID From ZuordnungWsEinzeltaetigkeit " & _
                                           "Where (((ZuordnungWsEinzeltaetigkeit.EinzelTaetigkeitID)=" & l_FehlerbildID & ") AND ((ZuordnungWsEinzeltaetigkeit.WsID)=" & rs("WsID") & "));")
                If (rs2.RecordCount) Then
                    rs2.MoveFirst
                    'Datensatz vorhanden: ID von Fehlerbild eintragen und Status Melder setzen
                    rs2.Edit
                    rs2("StatusMelder").Value = True
                    rs2("EinzelTaetigkeitMelderID").Value = rsNewFehlerbildID("MaxvonID")
                    rs2.Update
                Else
                    'Datensatz nicht vorhanden: Neuen datensatz anlegen
                    'Neues Fehlerbild einfügen und aktivieren
                    ldb.Execute ("INSERT INTO ZuordnungWsEinzeltaetigkeit ( WsID, EinzelTaetigkeitMelderID, StatusMelder ) " & _
                               "values (" & rs("WsID") & ", " & l_FehlerbildID & ", True);")
                End If
                rs.MoveNext
            Wend

        Else
            ldb.Execute ("UPDATE ZuordnungDefTaetigkeit set DefTaetigkeit='" & Me.DefTaetigkeit & "', Beschreibung='" & Me.Beschreibung & "'" & _
                        " WHERE ID =" & DefTaetigkeitID & ";")

        End If
        .Requery
        .Refresh
    End With
    FormKoorSave Me
    DoCmd.Close acForm, "Dialog_DefTaetigkeit"
    Forms("Fehlerbilder").SetFocus
Else
    MsgBox getTC_FCFromCode("TC0335", ldb)  'MsgBox ("Daten nicht vollständig!")   'TC0336
End If
end_OK_Click:
Set rs = Nothing
Set rs2 = Nothing
Set rsNewFehlerbildID = Nothing
Set adbprov = Nothing
 ' Set ldb = openAccDBConnection("closeDB")
Exit Sub

err_OK_Click:
    If Err = 3021 Then
        Resume Next
    Else
        GoTo end_OK_Click
    End If

End Sub
