Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Bearbeiten_Click()

    If Not IsNull(Me.ID) Then
        DoCmd.OpenForm "Dialog_Taetigkeit"
        'globale Variable setzen
        DialogTaetigkeit = "Bearbeiten"
        'globale Variable setzen
        TaetigkeitID = Me.ID
        'globale Variable setzen
        PTaetigkeit = Me.Taetigkeit

        With Forms("Dialog_Taetigkeit")
            .Controls("Taetigkeit") = PTaetigkeit
            .Controls("Beschreibung") = Me.UFo2.Controls("Beschreibung")
            .Controls("CboZuordnungID") = Me.[Zuordnung_ID]
            .Refresh
        End With
    Else
        MsgBox getTC_FCFromCode("TC0201")  'MsgBox "Nullwert. Keine Bearbeitung möglich!"   'TC0337
    End If

End Sub

Private Sub DatensatzSpeichern_Click()

    DatensatzÜbernehmen_Click

End Sub

Private Sub DatensatzÜbernehmen_Click()
'    Forms("Fehlererfassung - Register").Controls("Fehlerart") = Me.Fehlerart
'    Forms("Fehlererfassung - Register").Refresh
'    Me.SetFocus
'    DoCmd.Close

End Sub

Private Sub DefTaetigkeit_Click()

    DoCmd.OpenForm "Fehlerbilder", acNormal
    With Forms("Fehlerbilder")
        .Controls("ListTaetigkeiten") = Me.ID
        .Controls("ListTaetigkeiten").Requery
        .Requery
        .Refresh
        .Controls("ListTaetigkeiten").Locked = True
        .Controls("DefTaetigkeit").Locked = True
        .UF01.Controls("Beschreibung").Locked = True
        .Controls("Übernehmen").Enabled = False
        .AllowAdditions = False
    End With


End Sub

Private Sub Form_DblClick(Cancel As Integer)

    Bearbeiten_Click

End Sub

Private Sub Form_GotFocus()
    DoCmd.GoToRecord , , acPrevious, 5

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

Private Sub Löschen_Click()
' !!!!!!!!!!!!!!!!!!!!!!!! NEU - OFFEN !!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
On Error GoTo err_Löschen_Click

    Dim ldb As DAO.Database
    Dim rsDefTaet As DAO.Recordset

     Set ldb = openAccDBConnection("openDB")

    Set rsDefTaet = ldb.OpenRecordset("Select * from ZuordnungDefTaetigkeit where TaetigkeitID=" & Me.ID & ";")
    rsDefTaet.MoveFirst
    While Not rsDefTaet.EOF
        ldb.Execute ("UPDATE Mastertab set DefTatigkeitID=NULL" & _
                  " WHERE DefTatigkeitID = " & rsDefTaet("ID") & ";")
        rsDefTaet.MoveNext
    Wend

    ldb.Execute ("UPDATE Mastertab set TaetigkeitID=NULL" & _
                        " WHERE TaetigkeitID = " & Me.ID & ";")

    ldb.Execute ("Delete from Taetigkeiten where ID=" & Me.ID & ";")

    Set rsDefTaet = Nothing
     ' Set ldb = openAccDBConnection("closeDB")

    Me.Requery
    Me.Refresh

err_Löschen_Click:
    If Err = 3021 Then ' Kein aktueller Datensatz - Es gibt zur Fehlerbeschreibung Grob keine Fehlerbeschr. Fein
        Resume Next
    End If
End Sub

Private Sub Neu_Click()
    DoCmd.OpenForm "Dialog_Taetigkeit"
    DialogTaetigkeit = "Neu"
End Sub

Private Sub Schließen_Click()

    FormKoorSave Me
    DoCmd.Close
     If CurrentProject.AllForms("Fehlererfassung - Register").IsLoaded Then
        Forms("Fehlererfassung - Register").Requery
        Forms("Fehlererfassung - Register").Refresh
     End If


End Sub


Private Sub Übernehmen_Click()
'
'    Forms("Fehlererfassung - Register").Controls("Fehlerart") = Me.Fehlerart
'    Forms("Fehlererfassung - Register").Controls("AuswFehlerart") = Me.ID
'    Forms("Fehlererfassung - Register").Controls("FehlerartFein") = ""
'    Forms("Fehlererfassung - Register").Controls("AuswFehlerartFein") = Null
'    Forms("Fehlererfassung - Register").Refresh
'    Me.SetFocus
'    DoCmd.Close

End Sub
Sub Form_AfterOpen()

    'Dim ldb As DAO.Database
    'Dim rs As DAO.Recordset

    'Me.RecordsetClone.MoveLast
    'MsgBox Me.RecordsetClone.RecordCount

    'DoCmd.GoToRecord , , acLast
    DoCmd.GoToRecord , , acPrevious, 5

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub
