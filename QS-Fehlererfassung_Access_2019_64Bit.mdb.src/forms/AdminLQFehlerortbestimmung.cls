Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AddLQBereich_Click()
'Änderung:
'RSP-20150410 - Funktion um die Möglichkeit den seltierten Wert in der Liste umzubennen erweitert. Sobald ein Eintrag in der Liste selektiert ist, wird der User die Möglichkeit eröffnet
'               den selektierten Wert umzubennen oder neu anzulegen

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim bNew As Boolean
Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider

    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider
    Set ldb = openAccDBConnection("openDB")
    bNew = False

        If Not IsNull(Me.AddNewLQBereich) Then
            If Me.ListLQBereich.ItemsSelected.Count = 1 Then
                If MsgBox(getTC_FCFromCode("TC0478"), vbYesNo + vbDefaultButton2) = vbYes Then
                'If MsgBox("Soll der selektierte LQ Bereich umbenannt (Ja) oder ein neuer LQ Bereich erzeugt (Nein) werden?", vbYesNo + vbDefaultButton2) = vbYes Then
                    ldb.Execute ("Update LQBereich Set [LQBereich] = '" & Me.AddNewLQBereich & "' where [ID] = " & Me.ListLQBereich & ";")
                    adbprov.CommandExecuteAvista ("Update zlqpp_fe_bereich set [feb_name] = '" & Me.AddNewLQBereich & "' where [feb_id] = " & Me.ListLQBereich & ";")
                Else
                    bNew = True
                End If
            Else
                bNew = True
            End If
            If bNew Then
                Set rs = ldb.OpenRecordset("Select * from LQBereich where [ID-LQEinheit] = " & Me.ListLQEinheit & " and [LQBereich] = '" & Me.AddNewLQBereich & "';")
                If rs.RecordCount = 0 Then
                    ldb.Execute ("insert into LQBereich ([ID-LQEinheit], [LQBereich]) values (" & Me.ListLQEinheit & ", '" & Me.AddNewLQBereich & "');")
                    Me.AddNewLQBereich = Null
                    rs.Requery
                    rs.MoveLast
                    adbprov.CommandExecuteAvista ("Insert into zlqpp_fe_bereich ([feb_id],[fee_id],[feb_name],[feb_aktiv])" & _
                                                "Values (" & rs("ID") & "," & rs("ID-LQEinheit") & ", '" & rs("LQBereich") & "', 1);")
                Else
        '           MsgBox ("Eintrag schon vorhanden! Es sind keine doppelten Einträge erlaubt!") 'TC0214
                    MsgBox getTC_FCFromCode("TC0214", ldb)
                End If
            End If
        End If

    Set rs = Nothing
    Set adbprov = Nothing

    Me.Requery
    Me.Refresh

End Sub

Private Sub AddLQEinheit_Click()
'Änderung:
'RSP-20150410 - Funktion um die Möglichkeit den seltierten Wert in der Liste umzubennen erweitert. Sobald ein Eintrag in der Liste selektiert ist, wird der User die Möglichkeit eröffnet
'               den selektierten Wert umzubennen oder neu anzulegen

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim bNew As Boolean
Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider

    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider

    Set ldb = openAccDBConnection("openDB")
    bNew = False

    If Not IsNull(Me.AddNewLQEinheit) Then
        If Me.ListLQEinheit.ItemsSelected.Count = 1 Then
            If MsgBox(getTC_FCFromCode("TC0479"), vbYesNo + vbDefaultButton2) = vbYes Then
            'If MsgBox("Soll die selektierte LQ Einheit umbenannt (Ja) oder eine neue LQ Einheit erzeugt (Nein) werden?", vbYesNo + vbDefaultButton2) = vbYes Then
                ldb.Execute ("Update LQEinheit Set [LQEinheit] = '" & Me.AddNewLQEinheit & "' where [ID] = " & Me.ListLQEinheit & ";")
                adbprov.CommandExecuteAvista ("Update zlqpp_fe_einheit set [fee_name] = '" & Me.AddNewLQEinheit & "' where [fee_id] = " & Me.ListLQEinheit & ";")
            Else
                bNew = True
            End If
        Else
            bNew = True
        End If
        If bNew Then
            Set rs = ldb.OpenRecordset("Select * from LQEinheit where [LQEinheit] = '" & Me.AddNewLQEinheit & "';")
            If rs.RecordCount = 0 Then
                ldb.Execute ("insert into LQEinheit ([LQEinheit]) values ('" & Me.AddNewLQEinheit & "');")
                Me.AddNewLQEinheit = Null
                rs.Requery
                rs.MoveLast
                adbprov.CommandExecuteAvista ("Insert into zlqpp_fe_einheit ([fee_id],[fee_name],[fee_aktiv]) Values (" & rs("ID") & ", '" & rs("LQEinheit") & "',1);")
            Else
    '            MsgBox ("Eintrag schon vorhanden! Es sind keine doppelten Einträge erlaubt!")
                MsgBox getTC_FCFromCode("TC0214", ldb)
            End If
        End If
    End If

    Set rs = Nothing
    Set adbprov = Nothing

Me.Requery
Me.Refresh

End Sub




Private Sub AddLQFunktion_Click()
'Änderung:
'RSP-20150410 - Abfrage Löschen Funktion auf Update umgestellt. Bereich wird auf inaktiv gesetzt

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim bNew As Boolean

 Set ldb = openAccDBConnection("openDB")
bNew = False


If Not IsNull(Me.AddNewLQFunktion) Then
    If Me.ListLQFunktion.ItemsSelected.Count = 1 Then
        If MsgBox(getTC_FCFromCode("TC0480"), vbYesNo + vbDefaultButton2) = vbYes Then
        'If MsgBox("Soll die selektierte LQ Funktion umbenannt (Ja) oder eine neue LQ Funktion erzeugt (Nein) werden?", vbYesNo + vbDefaultButton2) = vbYes Then
            ldb.Execute ("Update LQFunktion Set [LQFunktion] = '" & Me.AddNewLQEinheit & "' where [ID] = " & Me.ListLQFunktion & ";")
        Else
            bNew = True
        End If
    Else
        bNew = True
    End If
    If bNew Then
        If Me.ListLQBereich.ListIndex >= 0 Then 'Ist ein Listenelement ausgewählt
            Set rs = ldb.OpenRecordset("Select * from LQFunktion where [ID-LQBereich] = " & Me.ListLQBereich & " and [LQFunktion] = '" & Me.AddNewLQFunktion & "';")
            If rs.RecordCount = 0 Then
                ldb.Execute ("insert into LQFunktion ([ID-LQBereich], [LQFunktion]) values (" & Me.ListLQBereich & ", '" & Me.AddNewLQFunktion & "');")
                Me.AddNewLQFunktion = Null
            Else
    '            MsgBox ("Eintrag schon vorhanden! Es sind keine doppelten Einträge erlaubt!")
                MsgBox getTC_FCFromCode("TC0214", ldb)

            End If
        Else
'            MsgBox ("Es kein Bereich einer LQ-Einheit ausgewählt!!")
            MsgBox getTC_FCFromCode("TC0217", ldb)
        End If
    End If
End If

Set rs = Nothing
'db.Close
 ' Set ldb = openAccDBConnection("closeDB")

Me.Requery
Me.Refresh

End Sub

Private Sub AddLQWorkstation_Click()
'Änderung:
'RSP-20150410 - Funktion um die Möglichkeit den seltierten Wert in der Liste umzubennen erweitert. Sobald ein Eintrag in der Liste selektiert ist, wird der User die Möglichkeit eröffnet
'               den selektierten Wert umzubennen oder neu anzulegen

On Error GoTo err_AddLQWorkstation_Click

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim bNew As Boolean

Set ldb = openAccDBConnection("openDB")
bNew = False

'Debug.Print ("In Proz: err_AddLQWorkstation_Click")
If Not IsNull(Me.AddNewLQWorkstation) Then
    If Me.ListLQEinheit.ItemsSelected.Count = 1 Then
        If MsgBox(getTC_FCFromCode("TC4215"), vbYesNo + vbDefaultButton2) = vbYes Then
            ldb.Execute ("Update LQWorkstation Set [LQWorkstation] = '" & Me.AddNewLQWorkstation & "' where [ID] = " & Me.ListLQWorkstation.ItemData(Me.ListLQWorkstation.ItemsSelected(0)) & ";")
        Else
            bNew = True
        End If
    Else
        bNew = True
    End If
    If bNew Then
        If Me.ListLQBereich.ListIndex >= 0 Then 'Ist ein Listenelement ausgewählt
            Set rs = ldb.OpenRecordset("Select * from LQWorkstation where [ID-LQBereich] = " & Me.ListLQBereich & " and [LQWorkstation] = '" & Me.AddNewLQWorkstation & "';")
            If rs.RecordCount = 0 Then
                ldb.Execute ("insert into LQWorkstation ([ID-LQBereich], [ID-LQFunktion], [LQWorkstation]) values (" & Me.ListLQBereich & ", 7, '" & Me.AddNewLQWorkstation & "');")
                Me.AddNewLQWorkstation = Null
            Else
'               MsgBox ("Eintrag schon vorhanden! Es sind keine doppelten Einträge erlaubt!")  'TC0218
                MsgBox getTC_FCFromCode("TC0214", ldb)
            End If
        Else
'            MsgBox ("Es wurde kein Bereich einer LQ-Funktion ausgewählt!!")
            MsgBox getTC_FCFromCode("TC0219", ldb)  'MsgBox ("Es kein Bereich einer LQ-Funktion ausgewählt!!")  'TC0219
        End If
    End If

End If

Set rs = Nothing
'db.Close
 ' Set ldb = openAccDBConnection("closeDB")

Me.Requery
Me.Refresh


Exit Sub
err_AddLQWorkstation_Click:
    MsgBox ("Error: " & Err & " - " & Error)
    Resume Next

End Sub

Private Sub DelLQBereich_Click()
'Änderung:
'RSP-20150410 - Abfrage Löschen Bereich auf Update umgestellt. Bereich wird auf inaktiv gesetzt
Dim ldb As DAO.Database
Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider

    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider


    Set ldb = openAccDBConnection("openDB")

        'If (MsgBox("Wollen Sie wirklich den LQ-Bereich '" & Me.ListLQBereich.Column(1) & "' löschen? Dabei werden alle abhängigen Funktionen und Workstations gelöscht!!!", vbYesNo, "Löschen LQ-Bereich") = vbYes) Then
        '    ldb.Execute ("Delete from LQBereich where [ID] = " & Me.ListLQBereich & ";")  'TC0220
        'End If
        If (MsgBox(getTC_FCFromCode("TC0220", ldb) & "'" & Me.ListLQBereich.Column(1) & "'" & getTC_FCFromCode("TC0338", ldb), vbYesNo, getTC_FCFromCode("TC0389", ldb)) = vbYes) Then
            ldb.Execute ("Update LQBereich set [Aktiv] = False where [ID] = " & Me.ListLQBereich & ";")
            adbprov.CommandExecuteAvista ("Update zlqpp_fe_bereich set [feb_aktiv] = 0 where [feb_id] = " & Me.ListLQBereich & ";")
        End If

    Set adbprov = Nothing

    Me.AddNewLQBereich = ""

    Me.Requery
    Me.Refresh

End Sub

Private Sub DelLQEinheit_Click()
'Änderung:
'RSP-20150410 - Abfrage Löschen Einheit auf Update umgestellt. Einheit wird auf inaktiv gesetzt

Dim ldb As DAO.Database
Dim adbprov As clsAvistaProvider 'Declaration Avista DB provider

    'Avista DB Provider Instanz erzeugen
    Set adbprov = New clsAvistaProvider

    Set ldb = openAccDBConnection("openDB")

        'If (MsgBox("Wollen Sie wirklich die LQ-Einheit '" & Me.ListLQEinheit.Column(1) & "' löschen? Dabei werden alle abhängigen Bereiche, Funktionen und Workstations gelöscht!!!", vbYesNo, "Löschen LQ-Einheit") = vbYes) Then
        '    ldb.Execute ("Delete from LQEinheit where [ID] = " & Me.ListLQEinheit & ";") 'TC0221
        'End If
        If (MsgBox(getTC_FCFromCode("TC0221", ldb) & "'" & Me.ListLQEinheit.Column(1) & "'" & getTC_FCFromCode("TC0338", ldb), vbYesNo, getTC_FCFromCode("TC0390", ldb)) = vbYes) Then
            ldb.Execute ("Update LQEinheit set [Aktiv] = False where [ID] = " & Me.ListLQEinheit & ";")
            adbprov.CommandExecuteAvista ("Update zlqpp_fe_einheit set [fee_aktiv] = 0 where [fee_id] = " & Me.ListLQEinheit & ";")
        End If

    Set adbprov = Nothing

    Me.AddNewLQEinheit = ""


Me.Requery
Me.Refresh


End Sub

Private Sub DelLQFunktion_Click()

Dim ldb As DAO.Database

 Set ldb = openAccDBConnection("openDB")

If (MsgBox(getTC_FCFromCode("TC0222") & " " & "'" & Me.ListLQFunktion.Column(1) & "'" & " " & getTC_FCFromCode("TC0388"), vbYesNo, getTC_FCFromCode("TC0391")) = vbYes) Then
    ldb.Execute ("Update LQFunktion set [Aktiv] = False where [ID] = " & Me.ListLQFunktion & ";")
End If


'db.Close
 ' Set ldb = openAccDBConnection("closeDB")

Me.Requery
Me.Refresh

End Sub

Private Sub DelLQWorkstation_Click()
'Änderung:
'RSP-20150410 - Abfrage Löschen Workstation auf Update umgestellt. WS wird auf inaktiv gesetzt
Dim ldb As DAO.Database

 Set ldb = openAccDBConnection("openDB")

'If (MsgBox("Wollen Sie wirklich die selektierten LQ-Workstations löschen?", vbYesNo, "Löschen LQ-Workstation") = vbYes) Then  'TC0223
If (MsgBox(getTC_FCFromCode("TC0223", ldb), vbYesNo, getTC_FCFromCode("TC0292", ldb)) = vbYes) Then
    If Me.ListLQWorkstation.ItemsSelected.Count Then
        For Each varElem In Me.ListLQWorkstation.ItemsSelected
            ldb.Execute ("Update LQWorkstation set [Aktiv] = False where [ID] = " & Me.ListLQWorkstation.ItemData(varElem) & ";")
        Next varElem
    End If
End If


'db.Close
 ' Set ldb = openAccDBConnection("closeDB")

Me.Requery
Me.Refresh

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

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

Private Sub InListWSFunk_Click() '20160314-ROS
On Error GoTo err_InListWSFunk_Click

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
 Set ldb = openAccDBConnection("openDB")

    'Ist ein Listenelement Funktion sowie Workstation ausgewählt
    If Me.ListLQFunktion.ListIndex >= 0 And Me.ListLQWorkstation.ItemsSelected.Count Then
        For Each varElem In Me.ListLQWorkstation.ItemsSelected
            Set rs = ldb.OpenRecordset("Select * from LQWSzuGrp where [ID-WS] = " & Me.ListLQWorkstation.ItemData(varElem) & " and [ID-LQFunktion] = " & Me.ListLQFunktion.ItemData(Me.ListLQFunktion.ListIndex) & ";")
            If rs.RecordCount = 0 Then
                ldb.Execute ("Insert into LQWSzuGrp ([ID-WS], [ID-LQFunktion]) values (" & Me.ListLQWorkstation.ItemData(varElem) & ", " & Me.ListLQFunktion.ItemData(Me.ListLQFunktion.ListIndex) & ");")
            End If
        Next varElem
        Me.ListLQWorkstation = Null
    Else
'        MsgBox ("Es ist keine Funktion bzw. Workstation ausgewählt!!")
        MsgBox getTC_FCFromCode("TC0224", ldb)  'MsgBox ("Es ist keine Funktion bzw. Workstation ausgewählt!!")  'TC0224
    End If


Set rs = Nothing
'db.Close
 ' Set ldb = openAccDBConnection("closeDB")

Me.Requery
Me.Refresh

Exit Sub
err_InListWSFunk_Click:
    MsgBox ("Error: " & Err & " - " & Error)
    Resume Next


End Sub

Private Sub ListLQBereich_AfterUpdate()
'Me.ListLQFunktion = -1
Me.ListLQWorkstation = -1
Me.AddNewLQBereich = Me.ListLQBereich.Column(1)
Me.AddNewLQWorkstation = ""
Me.AddNewLQFunktion = ""

Me.Requery
Me.Refresh

End Sub

Private Sub ListLQEinheit_AfterUpdate()
'MsgBox (Me.ListLQEinheit.ListIndex)
Me.ListLQBereich = -1
'Me.ListLQFunktion = -1
Me.ListLQWorkstation = -1
Me.AddNewLQEinheit = Me.ListLQEinheit.Column(1)
Me.AddNewLQBereich = ""
Me.AddNewLQWorkstation = ""
Me.AddNewLQFunktion = ""

Me.Requery
Me.Refresh
End Sub


Private Sub ListLQFunktion_AfterUpdate()
Me.ListLQWorkstation = -1
Me.AddNewLQFunktion = Me.ListLQFunktion.Column(1)
Me.Requery
Me.Refresh

End Sub

Private Sub ListLQWorkstation_AfterUpdate()
'Debug.Print Me.ListLQWorkstation.Column(0)
Me.AddNewLQWorkstation = Me.ListLQWorkstation.Column(1)
Me.Liste33 = Me.ListLQWorkstation.Column(0)
Me.Requery
Me.Refresh



End Sub

Private Sub MailZuordnenLQBereich_Click()

    DoCmd.OpenForm "LQBereichMail_Dialog"
    With Forms("LQBereichMail_Dialog")
        .Filter = "[ID] = " & Me.ListLQBereich
        .FilterOn = True
        .AllowAdditions = False
        .Controls("SelLQBereich") = Me.ListLQBereich
        .Requery
        .Refresh
    End With


End Sub

Private Sub MailZuordnenLQEinheit_Click()

DoCmd.OpenForm "LQEinheitMail_Dialog"
With Forms("LQEinheitMail_Dialog")
    .Filter = "[ID] = " & Me.ListLQEinheit
    .FilterOn = True
    .AllowAdditions = False
    .Controls("SelLQEinheit") = Me.ListLQEinheit
    .Requery
    .Refresh
End With

End Sub


Private Sub MailZuordnenLQWS_Click()

DoCmd.OpenForm "LQWSMail_Dialog"
With Forms("LQWSMail_Dialog")
    .Filter = "[ID] = " & Me.Liste33
    .FilterOn = True
    .AllowAdditions = False
    .Controls("SelLQWS") = Me.Liste33
    .Refresh
End With

End Sub

Private Sub OutListWSFunk_Click()   '20160314-ROS
On Error GoTo err_OutListWSFunk_Click

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
 Set ldb = openAccDBConnection("openDB")

    'Ist ein Listenelement Funktion sowie Workstation ausgewählt
    If Me.ListLQWorkstationFunkt.ItemsSelected.Count Then
        For Each varElem In Me.ListLQWorkstationFunkt.ItemsSelected
            ldb.Execute ("delete * from LQWSzuGrp where [ID-WS] = " & Me.ListLQWorkstationFunkt.ItemData(varElem) & " and [ID-LQFunktion] = " & Me.ListLQFunktion.ItemData(Me.ListLQFunktion.ListIndex) & ";")
        Next varElem
        Me.ListLQWorkstationFunkt = Null
    Else
        MsgBox getTC_FCFromCode("TC0225", ldb)  'MsgBox ("Es ist keine Workstation ausgewählt!!")  'TC0225
    End If

    Set rs = Nothing
'    ldb.Close
     ' Set ldb = openAccDBConnection("closeDB")

    Me.Requery
    Me.Refresh

Exit Sub
err_OutListWSFunk_Click:
    MsgBox ("Error: " & Err & " - " & Error)
    Resume Next

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
