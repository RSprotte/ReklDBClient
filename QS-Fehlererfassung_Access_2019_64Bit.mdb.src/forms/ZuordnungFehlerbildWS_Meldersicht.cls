Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub EinzelTaetigkeitLoeschen_Click()

Dim ldb As DAO.Database
Dim varElem As Variant

 Set ldb = openAccDBConnection("openDB")

If IsNull(Me.ListLQWorkstation) Then
'    MsgBox "Keine Workstation ausgewählt!"     'TC0317
    MsgBox getTC_FCFromCode("TC0307", ldb)
Else
    If Me.ListEinzelTaetigkeitenWSAktiv.ItemsSelected.Count Then
        For Each varElem In Me.ListEinzelTaetigkeitenWSAktiv.ItemsSelected
            ldb.Execute ("Update ZuordnungWsEinzeltaetigkeit set StatusMelder = No where WsID = " & Me.ListLQWorkstation & " AND " & _
                        "EinzelTaetigkeitMelderID = " & Me.ListEinzelTaetigkeitenWSAktiv.ItemData(varElem) & ";")
            'Hier Tabelle lqpp_fe_ws_taetigkeit (DBA) den Status der WS und des Fehlbild auf inaktiv setzen 0> gleiches SQL Statement.

            If Me.ListLQWorkstation.Column(2) <> "" Then
                'Declaration Avista DB provider
                Dim adbprov As clsAvistaProvider
                'Avista DB Provider Instanz erzeugen
                Set adbprov = New clsAvistaProvider
                    adbprov.CommandExecuteAvista ("Update zlqpp_fe_ws_fehler_info set [fewf_aktiv] = 0 where [lqpp_ws_id] = " & Me.ListLQWorkstation.Column(2) & " AND " & _
                            "[fef_id] = " & Me.ListEinzelTaetigkeitenWSAktiv.ItemData(varElem) & ";")
                Set adbprov = Nothing
            End If

        Next varElem
    End If
    Me.Requery
    Me.Refresh
End If

 ' Set ldb = openAccDBConnection("closeDB")

End Sub

Private Sub EinzelTaetigkeitZuordnen_Click()

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim varElem As Variant

 Set ldb = openAccDBConnection("openDB")
If IsNull(Me.ListLQWorkstation) Then
'    MsgBox "Keine Workstation ausgewählt!"     'TC0317
    MsgBox getTC_FCFromCode("TC0307", ldb)
Else
    'Prüfung ob das Fehlerbild der Workstation zugeordnet ist.
    '(Die Zuordnung der Fehlerbilder wird über Zurodnung der Tätigkeit gesammelt durchgeführt. Sobald im Nachhinein ein Fehlerbild der Tätigkeit ergänzt wird fehlt die Zuordnung an der WS)

    If Me.ListEinzelTaetigkeitenWS.ItemsSelected.Count Then
        For Each varElem In Me.ListEinzelTaetigkeitenWS.ItemsSelected
            'Bearbeitung Access
            Set rs = ldb.OpenRecordset("Select * from ZuordnungWsEinzeltaetigkeit where (WsID = " & Me.ListLQWorkstation & _
                        " AND (EinzelTaetigkeitMelderID = " & Me.ListEinzelTaetigkeitenWS.ItemData(varElem) & " OR " & _
                        "EinzelTaetigkeitID = " & Me.ListEinzelTaetigkeitenWS.ItemData(varElem) & "));")
            If rs.RecordCount Then
                'Fehlerbild aktivieren
                rs.Edit
                rs("EinzelTaetigkeitMelderID").Value = Me.ListEinzelTaetigkeitenWS.ItemData(varElem)
                rs("StatusMelder").Value = True
                rs.Update
'                ldb.Execute ("Update ZuordnungWsEinzeltaetigkeit set StatusMelder = Yes,  EinzelTaetigkeitMelderID = where WsID = " & Me.ListLQWorkstation & " AND " & _
'                            "EinzelTaetigkeitMelderID = " & Me.ListEinzelTaetigkeitenWS.ItemData(varElem) & ";")
            Else
                'Fehlerbild-Zuordnung anlegen und aktivieren
                ldb.Execute ("INSERT INTO ZuordnungWsEinzeltaetigkeit ( WsID, EinzelTaetigkeitMelderID, StatusMelder ) " & _
                        "values (" & Me.ListLQWorkstation & ", " & Me.ListEinzelTaetigkeitenWS.ItemData(varElem) & ", True);")
            End If
            'Bearbeitung Avista

            If Me.ListLQWorkstation.Column(2) <> "" Then
                'Abfrage ist Datenstaz in Avista vorhanden
                ' Ja: Update und Aktiv schalten
                Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider
                Dim rsAvista As ADODB.Recordset
                Set rsAvista = New ADODB.Recordset
                Set adbprov = New clsAvistaProvider
                ' Ja: Update und Aktiv schalten
                Set rsAvista = adbprov.getRecordsetAvista("SELECT * FROM [zlqpp_fe_ws_fehler_info] where([lqpp_ws_id] = " & Me.ListLQWorkstation.Column(2) & _
                            " AND ([fef_id] = " & Me.ListEinzelTaetigkeitenWS.ItemData(varElem) & " OR " & _
                            "[fef_id] = " & Me.ListEinzelTaetigkeitenWS.ItemData(varElem) & "));")
                If rsAvista.RecordCount Then
                    adbprov.CommandExecuteAvista ("Update zlqpp_fe_ws_fehler_info set [fewf_aktiv] = 1 where [lqpp_ws_id] = " & Me.ListLQWorkstation.Column(2) & " AND " & _
                            "[fef_id] = " & Me.ListEinzelTaetigkeitenWS.ItemData(varElem) & ";")
                Else
                    ' Nein: DS anlegen und Aktiv schalten
                    adbprov.CommandExecuteAvista ("Insert into zlqpp_fe_ws_fehler_info ([lqpp_ws_id],[fef_id],[fewf_aktiv]) Values (" & Me.ListLQWorkstation.Column(2) & ", " & Me.ListEinzelTaetigkeitenWS.ItemData(varElem) & ", 1);")
                End If
                Set rsAvista = Nothing
                Set adbprov = Nothing
            End If

        Next varElem
    End If
    Me.Requery
    Me.Refresh
End If

    Set rs = Nothing

 ' Set ldb = openAccDBConnection("closeDB")


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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub

Private Sub ListLQBereich_AfterUpdate()
Me.Requery
Me.Refresh
End Sub

Private Sub ListLQEinheit_AfterUpdate()
Me.Requery
Me.Refresh

End Sub

Private Sub ListLQWorkstation_AfterUpdate()

Me.ListTaetigkeitenWS = Null
Me.Requery
Me.Refresh


End Sub

Private Sub ListTaetigkeitenWS_AfterUpdate()
Me.Requery
Me.Refresh

End Sub

Private Sub Schließen_Click()
On Error GoTo Err_Schließen_Click


    FormKoorSave Me
    DoCmd.Close

Exit_Schließen_Click:
    Exit Sub

Err_Schließen_Click:
    MsgBox Err.Description
    Resume Exit_Schließen_Click

End Sub

Private Sub TaetigkeitLoeschen_Click()

Dim ldb As DAO.Database

 Set ldb = openAccDBConnection("openDB")
If IsNull(Me.ListLQWorkstation) Then
    MsgBox getTC_FCFromCode("TC048")
Else

    ldb.Execute ("UPDATE (Taetigkeiten INNER JOIN ZuordnungDefTaetigkeit ON Taetigkeiten.ID = ZuordnungDefTaetigkeit.TaetigkeitID) " & _
                "INNER JOIN ZuordnungWsEinzeltaetigkeit ON ZuordnungDefTaetigkeit.ID = ZuordnungWsEinzeltaetigkeit.EinzelTaetigkeitMelderID SET ZuordnungWsEinzeltaetigkeit.StatusMelder = No " & _
                "WHERE (((Taetigkeiten.ID)=" & Me.ListTaetigkeitenWS & ") AND ((ZuordnungWsEinzeltaetigkeit.WsID)=" & Me.ListLQWorkstation & "));")
    Me.Requery
    Me.Refresh
End If
 ' Set ldb = openAccDBConnection("closeDB")


End Sub

Private Sub TaetigkeitZuordnen_Click()

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim sSQL As String
Dim varElem As Variant

 Set ldb = openAccDBConnection("openDB")

If Me.ListLQWorkstation.ItemsSelected.Count = 0 Then
'    MsgBox "Keine Workstation ausgewählt!"     'TC0317
    MsgBox getTC_FCFromCode("TC0307", ldb)
Else
'Prüfung ob der Workstation schon genrell die Tätigkeit zugeordnet wurde
    For Each varElem In Me.ListLQWorkstation.ItemsSelected

        Set rs = ldb.OpenRecordset("SELECT ZuordnungWsEinzeltaetigkeit.WsID, ZuordnungDefTaetigkeit.TaetigkeitID " & _
                                  "FROM ZuordnungWsEinzeltaetigkeit LEFT JOIN ZuordnungDefTaetigkeit ON ZuordnungWsEinzeltaetigkeit.EinzelTaetigkeitMelderID = ZuordnungDefTaetigkeit.ID " & _
                                  "GROUP BY ZuordnungWsEinzeltaetigkeit.WsID, ZuordnungDefTaetigkeit.TaetigkeitID " & _
                                  "HAVING (((ZuordnungWsEinzeltaetigkeit.WsID)=" & Me.ListLQWorkstation.ItemData(varElem) & ") AND ((ZuordnungDefTaetigkeit.TaetigkeitID)=" & Me.ListTaetigkeiten & "));")

        If rs.RecordCount = 1 Then
            'Keine weiter Aktion, wenn die Einzeltätikeiten inaktiv sind müssen diese manuell über den Dialog aktiviert werden
            MsgBox getTC_FCFromCode("TC0311", ldb)  'MsgBox "Taetigkeit ist WS schon zugeordnet"   'TC0318
        Else
            sSQL = "INSERT INTO ZuordnungWsEinzeltaetigkeit ( WsID, EinzelTaetigkeitMelderID, StatusMelder ) SELECT " & Me.ListLQWorkstation.ItemData(varElem) & " AS [WS-ID], ZuordnungDefTaetigkeit.ID, Yes AS Status " & _
                        "FROM Taetigkeiten LEFT JOIN ZuordnungDefTaetigkeit ON Taetigkeiten.ID = ZuordnungDefTaetigkeit.TaetigkeitID " & _
                        "WHERE (((Taetigkeiten.ID) = " & Me.ListTaetigkeiten & ")) GROUP BY ZuordnungDefTaetigkeit.ID;"
            ldb.Execute (sSQL)
            'Hier Tabelle lqpp_fe_ws_taetigkeit (DBA) mit dem selben SQL Statement beschreiben.
            If Me.ListLQWorkstation.Column(2) <> "" Then
                'Declaration Avista DB provider
                Dim adbprov As clsAvistaProvider
                'Avista DB Provider Instanz erzeugen
                Set adbprov = New clsAvistaProvider
                    adbprov.CommandExecuteAvista ("INSERT INTO zlqpp_fe_ws_fehler_info ( [lqpp_ws_id], [fef_id], [fewf_aktiv] ) SELECT " & Me.ListLQWorkstation.Column(2) & " AS [lqpp_ws_id], zlqpp_fe_fehler_info.fef_id, 1" & _
                            "FROM zlqpp_fe_taetigkeit LEFT JOIN zlqpp_fe_fehler_info ON zlqpp_fe_taetigkeit.fet_id = zlqpp_fe_fehler_info.fet_id " & _
                            "WHERE (((zlqpp_fe_taetigkeit.fet_id) = " & Me.ListTaetigkeiten & ")) GROUP BY zlqpp_fe_fehler_info.fef_id;")
                Set adbprov = Nothing
            End If
        End If
        Set rs = Nothing
    Next varElem
    Me.Requery
    Me.Refresh

End If
 ' Set ldb = openAccDBConnection("closeDB")



End Sub

Private Sub Umschaltfläche31_Click()

Dim LQE As String
Dim LQB As String

LQE = Me.ListLQEinheit
LQB = Me.ListLQBereich

Application.Echo False


    If Me.Umschaltfläche31 Then


    '    If Me.ListLQWorkstation.MultiSelect = 2 Then
            DoCmd.OpenForm "ZuordnungFehlerbildWS_Meldersicht", acViewDesign, , , , acHidden

            Forms("ZuordnungFehlerbildWS_Meldersicht").Controls("ListLQWorkstation").MultiSelect = 2
            'DoCmd.Close acForm, "ZuordnungFehlerbildWS_Meldersicht", acSaveYes
            DoCmd.OpenForm "ZuordnungFehlerbildWS_Meldersicht", acNormal
            Forms("ZuordnungFehlerbildWS_Meldersicht").Controls("Umschaltfläche31") = True
            Forms("ZuordnungFehlerbildWS_Meldersicht").Controls("ListLQEinheit") = LQE
            Forms("ZuordnungFehlerbildWS_Meldersicht").Controls("ListLQBereich") = LQB
            Forms("ZuordnungFehlerbildWS_Meldersicht").Refresh
    Else
            DoCmd.OpenForm "ZuordnungFehlerbildWS_Meldersicht", acViewDesign, , , , acHidden
            Forms("ZuordnungFehlerbildWS_Meldersicht").Controls("ListLQWorkstation").MultiSelect = 0
            'DoCmd.Close acForm, "ZuordnungFehlerbildWS_Meldersicht", acSaveYes
            DoCmd.OpenForm "ZuordnungFehlerbildWS_Meldersicht", acNormal
            Forms("ZuordnungFehlerbildWS_Meldersicht").Controls("Umschaltfläche31") = False
            Forms("ZuordnungFehlerbildWS_Meldersicht").Controls("ListLQEinheit") = LQE
            Forms("ZuordnungFehlerbildWS_Meldersicht").Controls("ListLQBereich") = LQB
            Forms("ZuordnungFehlerbildWS_Meldersicht").Refresh
    End If
Application.Echo True

End Sub

Private Sub Umschaltfläche31_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call CursorWechseln(IDC_HAND)

End Sub
