Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Dim gi_ShutdownInterval As Integer
Dim gi_ShutdownMinuten As Long

Const ReklAuswahlIntExt = "Extern"


Private Sub Datenbankfenster_Anzeigen_Click()
' Dieser Code wurde teilweise vom Befehlsschaltflächen-Assistenten erstellt.
On Error GoTo Err_Datenbankfenster_Anzeigen_Click

    Dim strDocName As String

    strDocName = "Lieferanten"

    ' Formular "Hauptübersicht" schließen.
    DoCmd.Close

    ' Focus auf Datenbankfenster setzen; Tabelle "Lieferanten" auswählen (erster
    ' Eintrag in der Liste).
    DoCmd.SelectObject acTable, strDocName, True

Exit_Datenbankfenster_Anzeigen_Click:
    Exit Sub

Err_Datenbankfenster_Anzeigen_Click:
    MsgBox Err.Description
    Resume Exit_Datenbankfenster_Anzeigen_Click

End Sub


Private Sub FehlererfassungTest_Click()
Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim strKrit As String

'get_p_ReklNr () = InputBox("Geben Sie die Reklamationsnummer ein!", "Eingabe Vorgang")    'IC0217
'p_ReklNr = InputBox(getTC_FCFromCode("IC0200"), getTC_FCFromCode("IC0219"))
p_ReklNr = getValueFromInputBox(getTC_FCFromCode("IC0200"), getTC_FCFromCode("IC0219"))
If StrPtr(p_ReklNr) = 0 Then   ' Prüfen ob Abbrechen gedrückt wurde
    GoTo end_InterneFehlererfassung4_Click
End If
If get_p_ReklNr() <> "" Then
    strKrit = "ID = " & get_p_ReklNr()
    If DCount("ID", "Mastertab", strKrit) = 0 Then
        MsgBox getTC_FCFromCode("TC0205") & get_p_ReklNr() & getTC_FCFromCode("TC0379")    'MsgBox ("Reklamation mit der Nummer: " & get_p_ReklNr () & " ist nicht existent!")   'TC0357
        GoTo end_InterneFehlererfassung4_Click
    Else
         Set ldb = openAccDBConnection("openDB")
        Set rs = ldb.OpenRecordset("Select * from Mastertab where ID = " & get_p_ReklNr() & ";")
        If rs.RecordCount Then
            If (Not IsNull(rs("ErfasserIntern").Value)) Then
                If Me.OptionReklAuswahl = 1 Then
'                    If (MsgBox("Soll der Vorgang mit der externe Fehlererfasung geöffnet werden?", vbYesNo, "Achtung") = vbYes) Then      'TC0358
'                        NewReklamationE "Bearbeiten", get_p_ReklNr ()
                    If (MsgBox(getTC_FCFromCode("TC0358", ldb), vbYesNo, getTC_FCFromCode("TC0408"), ldb) = vbYes) Then
                        NewReklamationE "Bearbeiten", get_p_ReklNr()
                    End If
                Else
                    If get_p_ReklNr() = "" Then
                        NewReklamationI
                    Else
                        OpenReklamationI
                    End If
                End If
            Else
                If Me.OptionReklAuswahl = 2 Then
'                    If (MsgBox("Die externe Reklamation kann nicht über die intern Fehlererfassung bearbeitet werden. Soll der Vorgang mit der externe Fehlererfasung geöffnet werden?", vbYesNo, "Achtung") = vbYes) Then    'TC0359
'                        NewReklamationE "Bearbeiten", get_p_ReklNr ()
                    If (MsgBox(getTC_FCFromCode("TC0359", ldb), vbYesNo, getTC_FCFromCode("TC0408"), ldb) = vbYes) Then    'TC0359
                        NewReklamationE "Bearbeiten", get_p_ReklNr()
                    End If
                Else
                    NewReklamationE "Bearbeiten", get_p_ReklNr()
                End If
            End If
        End If
        Set rs = Nothing 'Nothing
         '' Set ldb = openAccDBConnection("closeDB")
    End If
Else
    If Me.OptionReklAuswahl = 1 Then 'Extern
        NewReklamationE "Erfassen", ""
    Else
        MsgBox getTC_FCFromCode("TC0212"), vbOKOnly  'Die interne Fehlererfassung ist aktuell nur über den Thin-Client möglich!", vbOKOnly
        'NewReklamationI
    End If

End If
end_InterneFehlererfassung4_Click:

End Sub

Private Sub Befehl5_Click()
On Error GoTo Err_Befehl5_Click

    If CurrentProject.AllForms("frmFormularROS").IsLoaded Then
        Forms("frmFormularROS").SetFocus
        DoCmd.Close
    End If

    Call ClientAbmelden
    FormKoorSave Me
    openAccDBConnection ("closeDB")
    DoCmd.Quit

'    Dim wmi, programm, process, system
'    programm = "MSACCESS.EXE"
'
'    Set wmi = GetObject("winmgmts:")
'    Set system = wmi.instancesOf("win32_process")
'    For Each process In system
'        If LCase(process.Name) = programm Then
'            process.Terminate (0)
'        End If
'    Next

Exit_Befehl5_Click:
    Exit Sub

Err_Befehl5_Click:
    MsgBox Err.Description    'FC0202
    Resume Exit_Befehl5_Click

End Sub

Private Sub cboReiterAdmin_Konfig_Click()

    Select Case cboReiterAdmin_Konfig
        Case 18     '"LQ-Fehlerortbestimmung"
          Call OpenForm_LQFehlerortbestimmung
        Case 19     '"Konfig Zuordnung Fehlerbild zu WS - Analyse"
          Call OpenForm_LQ_ZuordnungWSAnalyse
        Case 20     '"Konfig Zuordnung Fehlerbild zu WS - Erfasser"
          Call OpenForm_LQ_ZuordnungWSFTErfasser
        Case 21     '"Konfig LQ-Tätigkeiten"
          Call OpenForm_LQ_Taetigkeiten
        Case 22     '"Append&TableDefs"
          Call Append_TableDefs
        Case 23     '"Verbindungstest Avista"
          Call ConnectASA1
        Case 24     '"Konfiguration Enviroment"
          Call KonfReklDBEnvir
        Case 25     '"VBA-Editor On/Off"
          Call ToggleVBA_Editor
        Case 29     '"Vorgang Status Zuordnung"
            Call VorgangStatusZuordnung
        Case 31     '"Konfig Ticket-Typ"
            Call OpenForm_TicketTypKonfig
        Case 32     '"Konfig Ticket-Typ zu Kürzel"
            Call ZuordnungKürzelZuTicket
        Case 36     '"Bearbeitungszeit ReklQS"
            Call ReklBaerbZeit
        Case 37     '"Konfig MsgBox"
            Call KonfigMsgBox
        Case 38     '"Zuornung AVISTA-Ws ReklDB-WS"
            Call OpenFormx_frmAccessWsID_AvistaWsID
        Case 40 'Liefer- oder andere Infos Kunde
            Call OpenFormx_frmInfoTagKunde
        Case 43 'Konfig Email Service
            Call OpenFormx_frmEmailService_Daten
        Case 44 'frmRSP_Lokalisierung
            Call OpenFormx_frmRSP_Lokalisierung

     End Select

    Me.cboReiterAdmin_Konfig = ""
    Me.InterneFehlererfassung4.SetFocus

End Sub

Private Sub cboReiterAdmin_Konfig_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CursorWechseln(IDC_HAND)
    Me.cboReiterAdmin_Konfig.SetFocus
    Me.cboReiterAdmin_Konfig.Dropdown
End Sub

Private Sub cboReiterHilfe_Click()

    Select Case cboReiterHilfe
        Case 26     '"LQ ReklamationDB Hilfe"
'          Call
        Case 27     '"Info"
'          Call
    End Select

    Me.cboReiterHilfe = ""
    Me.InterneFehlererfassung4.SetFocus

End Sub

Private Sub cboReiterHilfe_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CursorWechseln(IDC_HAND)
    Me.cboReiterHilfe.SetFocus
    Me.cboReiterHilfe.Dropdown

End Sub

Private Sub cboReiterDatei_Click()

    Select Case cboReiterDatei
        Case 1      '"Ein-/Ausblenden Hauptmenü"
          Call ToggleDB_Win
        Case 2      '"Schließen"
          Call DBSchließen
        Case 28      '"Ribbon On/Off"
          Call cboRibbonOn_Click
       End Select

    Me.cboReiterDatei = ""
    Me.InterneFehlererfassung4.SetFocus

End Sub

Private Sub cboReiterMaßnahmen_Click()

    Select Case cboReiterMaßnahmen
        Case 8      '"Maßnahmen Dialog"
          Call OpenDialogMassnahmen
       End Select

    Me.cboReiterMaßnahmen = ""
    Me.InterneFehlererfassung4.SetFocus

End Sub

Private Sub cboReiterMaßnahmen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CursorWechseln(IDC_HAND)
    Me.cboReiterMaßnahmen.SetFocus
    Me.cboReiterMaßnahmen.Dropdown
End Sub

Private Sub cboReiterReklamationen_Click()

    Select Case cboReiterReklamationen
        Case 3      '"Erfassen Interner Vorgang"
            Call ReklamationI
        Case 4      '"Erfassen Externer Vorgang"
            Call ReklamationE
        Case 5      '"Nach Meldungsnummer suchen"
            Call OpenReklamation2
        Case 6      '"Bearbeiten"
            Call OpenReklamation
        Case 7      '"Löschen"
            Call DelReklVorgang
       End Select

    Me.cboReiterReklamationen = ""
    Me.InterneFehlererfassung4.SetFocus

End Sub

Private Sub cboReiterReklamationen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CursorWechseln(IDC_HAND)
    Me.cboReiterReklamationen.SetFocus
    Me.cboReiterReklamationen.Dropdown
End Sub

Private Sub cboReiterÜbersichten_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CursorWechseln(IDC_HAND)
    Me.cboReiterÜbersichten.SetFocus
    Me.cboReiterÜbersichten.Dropdown
End Sub

Private Sub cboReiterDatei_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CursorWechseln(IDC_HAND)
    Me.cboReiterDatei.SetFocus
    Me.cboReiterDatei.Dropdown
End Sub

Private Sub cmbHautfenster_Click()

    AccessFenster

End Sub

Private Sub CmbVorgangAnlegen_Click()
On Error GoTo err_CmbVorgangAnlegen_Click

Dim ldb As DAO.Database
Dim i As Long
Dim iVorgang As Integer


    Set ldb = openAccDBConnection("openDB")

'    iVorgang = InputBox("Wieviele Vorgänge sollen angelegt werden?", "Vorgänge anlegen")
'    iVorgang = InputBox(getTC_FCFromCode("IC3851"), getTC_FCFromCode("IC3852"))
    iVorgang = getValueFromInputBox(getTC_FCFromCode("IC3851"), getTC_FCFromCode("IC3852"))
    For i = 1 To iVorgang

'        InterneFehlererfassung4_Click "New", "", "", False, False
        NewReklamationE "New", "", "", False, False
        ldb.Execute ("Update Mastertab set [Datum] = #9999-01-01#  where ID = " & p_ReklNr & ";")
        ldb.Execute ("Update Mastertab set [Kategorie-Status] = 'Platzhalter' where ID = " & p_ReklNr & ";")
        ldb.Execute ("Update Mastertab set [Verursacher] = 'LQ' where ID = " & p_ReklNr & ";")
    Next i
Exit Sub

err_CmbVorgangAnlegen_Click:

'    Debug.Print "Error: " & Err & " - " & Error

End Sub



Private Sub CmbVorgangsBerichtDrucken_Click()

    DoCmd.OpenReport "VorgangsBerichtTabelle", acViewNormal
'    DoCmd.OpenReport "VorgangsBegleitschein", acViewNormal

End Sub

Private Sub cmdDynFunc_Click(fIndex As Integer)

    Dim s As String
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim lForm As Form

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("SELECT * from tblSoftKeys where SchbearbeiterKuerzel = '" & get_pKuerzel() & "' and SoftkeyNummer = " & fIndex & ";", dbOpenSnapshot)
    If rs.RecordCount Then
        Application.Run rs("SoftkeyMakro")
    End If

    Set rs = Nothing 'Nothing

End Sub

Private Sub cmdDynFunc1_Click()

    cmdDynFunc_Click 1

End Sub

Private Sub cmdDynFunc2_Click()

cmdDynFunc_Click 2

End Sub


Private Sub cmdDynFunc3_Click()

cmdDynFunc_Click 3

End Sub


Private Sub cmdDynFunc4_Click()

cmdDynFunc_Click 4

End Sub


Private Sub cmdDynFunc5_Click()

cmdDynFunc_Click 5

End Sub

Private Sub Form_Activate()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim i As Integer
    Dim ctl As Control
    Dim l_i As Integer
    Dim l_ctlName As String

    setCmbViewFalse Me
    If get_ServiceZuordnung = True Then getCmbView Me, "Service"

    If get_AdminZuordnung = True Then
        Me.BezEntwicklung.Visible = True
        Me.OptEntwicklungAuswahl.Visible = True
        Me.BezJa.Visible = True
        Me.BezNein.Visible = True
        Me.OptJa.Visible = True
        Me.OptNein.Visible = True
        Me.cmbHautfenster.Visible = True
    Else
        Me.BezEntwicklung.Visible = False
        Me.OptEntwicklungAuswahl.Visible = False
        Me.BezJa.Visible = False
        Me.BezNein.Visible = False
        Me.OptJa.Visible = False
        Me.OptNein.Visible = False
        Me.cmbHautfenster.Visible = False
    End If

'#######################################################    20190906-ROS
'Anzahl Sofkeys wird ermittelt damit bei hinzufügen von weitern Softkeys die Variable i automatisch hochgezählt wird
    For Each ctl In Me.Controls
        l_ctlName = Left(ctl.Name, 10)
        If l_ctlName = "cmdDynFunc" Then l_i = l_i + 1
    Next ctl
'#######################################################

    'SoftKey Bearbeitung
    Set ldb = openAccDBConnection("openDB")
    'Über alle SoftKeys
    For i = 1 To l_i
        Set rs = ldb.OpenRecordset("SELECT * from tblSoftKeys where SchbearbeiterKuerzel = '" & get_pKuerzel() & "' and SoftkeyNummer = " & i & ";", dbOpenSnapshot)
        If rs.RecordCount Then
            Me.Controls("cmdDynFunc" & i).Caption = rs("SoftKexBezeichnung")
            Me.Controls("cmdDynFunc" & i).Visible = True
        Else
            Me.Controls("cmdDynFunc" & i).Visible = False
        End If
    Next i

    Set rs = Nothing 'Nothing

End Sub



Private Sub Form_Close()

    Call ClientAbmelden
    FormKoorSave Me

End Sub



Private Sub Form_Load()

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim Batch_Pfad As String
Dim Entwicklung As String
Dim lDatum As String

    Set ldb = openAccDBConnection("openDB")

    Entwicklung = DLookup("Value", "KonfigurationReklDBLocal", "Konfiguration = 'Entwicklung'")
    If Entwicklung = "Nein" Then
        Me.OptEntwicklungAuswahl = 2
    End If
    If get_ServiceZuordnung = True Then
        lDatum = DLookup("Value", "KonfigurationReklDB", "Konfiguration = 'DatenbankCopy'")
        If Entwicklung = "Nein" And lDatum <> Date Then
            ldb.Execute ("Delete * from [Mastertab-DSBerechtigung] where [Datum] < (Date() - 28);")
            ldb.Execute ("Update [Mastertab-DSBerechtigung] set [Zustand] = 'Frei', [Datum] = date() where ([Zustand] = 'Gesperrt' And [Datum] < date());")
            ldb.Execute ("Delete * from [tblHost] where [Datum] < (Date() - 28);")
        End If
    End If

    Call ClientAnmelden

    setCmbViewFalse Me
    If get_ServiceZuordnung = True Then getCmbView Me, "Service"

    If get_AdminZuordnung = True Then
        Me.BezEntwicklung.Visible = True
        Me.OptEntwicklungAuswahl.Visible = True
        Me.BezJa.Visible = True
        Me.BezNein.Visible = True
        Me.OptJa.Visible = True
        Me.OptNein.Visible = True
        Me.cmbHautfenster.Visible = True
    Else
        Me.BezEntwicklung.Visible = False
        Me.OptEntwicklungAuswahl.Visible = False
        Me.BezJa.Visible = False
        Me.BezNein.Visible = False
        Me.OptJa.Visible = False
        Me.OptNein.Visible = False
        Me.cmbHautfenster.Visible = False
    End If

    Dim objMouseOverForm As clsMouseoverForm
    Set objMouseOverForm = New clsMouseoverForm
    With objMouseOverForm
        Set .Form = Me
    End With

End Sub



Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Form_Open
    'RS - 2012-06-11 Verionausgabe eingebunden
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim prpLoop As DAO.Property '20150825-RSPTag
    Dim Newprp As DAO.Property '20150825-RSPTag
    Dim TEMP() As String '20150825-RSPTag

'    get_pKuerzel

    AccessFenster
    blFenster = True
'    FormOpenKoor Me
'    Call FormEinUndAusBlenden(Me)

    Call ClientParameterAbfragen

    Set ldb = openAccDBConnection("openDB")      '20191001-ROS
    Set rs = ldb.OpenRecordset("Select * from KonfigurationReklDB where Konfiguration = 'ClientShutdown';", dbOpenSnapshot)
    'Set rs = getOpenRecordset("Select * from KonfigurationReklDB where Konfiguration = 'ClientShutdown';", dbOpenSnapshot)
    If rs("Value") = "Ja" And Not getClientShutdownOwner = Environ("COMPUTERNAME") Then
'            MsgBox ("Datenbank steht momentan nicht zur Verfügung")
        MsgBox getTC_FCFromCode("TC3217"), vbExclamation, getTC_FCFromCode("TC3218")
        DoCmd.Quit
    End If
'Set rs = setCloseRecordset(rs) 'Nothing
    Set rs = Nothing 'Nothing


'    'einblenden
'    DoCmd.SelectObject acTable, , True
'
'    'Ausblenden Hauptfenster
'    DoCmd.SelectObject acTable, , True
'    RunCommand acCmdWindowHide
'    bToggleDB_Win = True
'    bToggleVBA_Editor = True

    'DoCmd.Maximize

    'RS - 2012-06-11 Verionausgabe eingebunden
    Me.Version = getClientVersion()
    'RSP - 20131110 Erweiterung
    'Anzeige der zentralen Datenbank auf der Startmaske mit dem der Client verbunden ist
    Me.ServerMDB = getServerLocation()

    '20190415-ROS Abfrage ob Tatsächlich auf Entwicklungsdatenbank zugegriffen werden soll
    If DLookup("Value", "KonfigurationReklDB", "[Konfiguration] = 'Datenbank'") = "Entwicklung" Then
        If MsgBox(getTC_FCFromCode("TC3855"), vbQuestion + vbDefaultButton1 + vbOKCancel, getTC_FCFromCode("TC3856")) = vbCancel Then
            DoCmd.Quit
        End If
    End If


    '20160205-ROS###Öffnet Dialog Append_TableDefs falls keine Datenbank ausgewählt ist (Menüleiste fehlt)
    If getServerLocation() = "" Then
        Call Append_TableDefs
    End If


'    CommandBars.ActiveMenuBar.Enabled = False
'    CommandBars("Database").Visible = False


    'Setze Globale Variabelen aus der KonfigurationsDatenbank
    'Set ldb = openAccDBConnection("openDB")
    'Set rs = getOpenRecordset("Select * from KonfigurationReklDBLocal;", dbOpenSnapshot)


    'Init Status Ribbon Anzeige
    'db.Execute ("Update KonfigurationReklDBLocal set RibbonOnOff = 'Aus';")

    ldb.Execute ("Update KonfigurationReklDBLocal set [Value] = 'Aus' where [Konfiguration] = 'RibbonOnOff';")
    'db.Execute ("Update KonfigurationReklDBLocal set [Value] = 'Aus' where [Konfiguration] = 'RibbonOnOff';")

    'Menüleisten und Ribbons init
    'neueMLeiste
'    DoCmd.ShowToolbar "Ribbon", acToolbarNo '20160513-ROS Ribbon weg

    'Name der Applikation an den Dateinamen anpassen 20150825-RSPTag
'    Set ldb = openAccDBConnection("openDB")
    TEMP = Split(db.Properties("Name"), "\")
    ldb.Properties!AppTitle = Left(TEMP(UBound(TEMP)), Len(TEMP(UBound(TEMP))) - 4)
    Application.RefreshTitleBar


    'Reklamationsauswahl festlegen

    Select Case Application.CurrentProject.Name
        Case "QS-FehlererfassungEntwicklung.mdb"
            Me.OptionReklAuswahl = 1
        Case "QS-FehlererfassungExtern.mdb"
            Me.OptionReklAuswahl = 1
        Case "QS-FehlererfassungIntern.mdb"
            Me.OptionReklAuswahl = 2
            Me.OptionReklAuswahl.Locked = True
        Case Else
            Me.OptionReklAuswahl = 1
    End Select

'#######################################################    20190906-ROS
Dim i As Integer
Dim ctl As Control
Dim l_i As Integer
Dim l_ctlName As String
'Anzahl Sofkeys wird ermittelt damit bei hinzufügen von weitern Softkeys die Variable i automatisch hochgezählt wird
    For Each ctl In Me.Controls
        l_ctlName = Left(ctl.Name, 10)
        If l_ctlName = "cmdDynFunc" Then l_i = l_i + 1
    Next ctl
'#######################################################

    'SoftKey Bearbeitung
'    Set ldb = openAccDBConnection("openDB")
    'Über alle SoftKeys
    For i = 1 To l_i
        Set rs = ldb.OpenRecordset("SELECT * from tblSoftKeys where SchbearbeiterKuerzel = '" & get_pKuerzel() & "' and SoftkeyNummer = " & i & ";", dbOpenSnapshot)
        If rs.RecordCount Then
            Me.Controls("cmdDynFunc" & i).Caption = rs("SoftKexBezeichnung")
            Me.Controls("cmdDynFunc" & i).Visible = True
        Else
            Me.Controls("cmdDynFunc" & i).Visible = False
        End If
    Next i

    Set rs = Nothing 'Nothing

End_Form_Open:
        Exit Sub

Err_Form_Open:

        If Err = -2147467259 Or Err = 5 Then
            Resume Next
        ElseIf Err = 3270 Then '20150825-RSPTag
            Set Newprp = ldb.CreateProperty("AppTitle", dbText, Left(TEMP(UBound(TEMP)), Len(TEMP(UBound(TEMP))) - 4))
            ldb.Properties.Append Newprp
            Resume
        End If
End Sub

Private Sub Form_Timer()
'On Error Resume Next
'
'Dim ldb As DAO.Database
'Dim rs As DAO.Recordset
'
'
'    DoCmd.OpenForm "frmWarnmeldung", , , , , acHidden
'
'
'    If Not get_AdminZuordnung Or Not getClientShutdownOwner = Environ("COMPUTERNAME") Then
'        Set ldb = openAccDBConnection("openDB")
'
'        Set rs = ldb.OpenRecordset("Select * from KonfigurationReklDB where Konfiguration = 'ClientShutdown';", dbOpenSnapshot)
'
'        If rs("Value") = "Nein" Then
'            Set rs = Nothing 'Nothing
'        Else
'            If gi_ShutdownInterval = 0 Then
'                Forms("frmWarnmeldung").SetFocus
'                Forms("frmWarnmeldung").Controls("Bezeichnungsfeld0").Caption = vbNewLine & _
'                        "Achtung! Achtung! Achtung! Achtung! Achtung! Achtung! Achtung! " & vbNewLine & vbNewLine & _
'                        "Bitte alle Änderungen speichen und Datenbank-Client beenden." & vbNewLine & _
'                        "Es müssen Wartungsarbeiten an der Datenbank durchgeführt werden." & vbNewLine & _
'                        "Der Datenbank-Client wird in ca 5min. automatisch heruntergefahren!"
'                Forms("frmWarnmeldung")!Bezeichnungsfeld0.Caption.Refresh
'                gi_ShutdownInterval = 1
'            Else
'                gi_ShutdownMinuten = gi_ShutdownMinuten + 1
'                Forms("frmWarnmeldung").SetFocus
'                Forms("frmWarnmeldung").Controls("Bezeichnungsfeld0").Caption = vbNewLine & _
'                "Datenbank-Client wird in wenigen Minuten automatisch heruntergefahren!"
'                Forms("frmWarnmeldung")!Bezeichnungsfeld0.Caption.Refresh
'                If gi_ShutdownInterval = 1 Then     '5 Minuten Wartezeit Formular-Timmer steht auf 1 min
'                    If gi_ShutdownMinuten = 5 Then
'                        gi_ShutdownMinuten = 0
'                        gi_ShutdownInterval = gi_ShutdownInterval + 1
'                    End If
'                ElseIf gi_ShutdownInterval = 2 Then '2 Minuten Overtime
'                    If gi_ShutdownMinuten = 2 Then
'                        gi_ShutdownInterval = 0
'                        gi_ShutdownMinuten = 0
'                        DoCmd.Save
'                        DoCmd.Quit ' Befehl zum Herunterfahren
'                    End If
'                End If
'            End If
'        End If
'        Set rs = Nothing 'Nothing
'    End If

End Sub

Private Sub InterneFehlererfassung4_Click()

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim strKrit As String

'get_p_ReklNr () = InputBox("Geben Sie die Reklamationsnummer ein!", "Eingabe Vorgang")    'IC0218
'p_ReklNr = InputBox(getTC_FCFromCode("IC0200"), getTC_FCFromCode("IC0219"))

p_ReklNr = getValueFromInputBox(getTC_FCFromCode("IC0200"), getTC_FCFromCode("IC0219"))

If p_ReklNr = "X" Then   ' Prüfen ob Abbrechen gedrückt wurde
    GoTo end_InterneFehlererfassung4_Click
End If
If p_ReklNr <> "" Then
'    fprintLogRSP "InterneFehlererfassung4_Click", "Benutzer: " & get_pKuerzel & " - Open ReklID: " & p_ReklNr, "Mail"
    strKrit = "ID = " & get_p_ReklNr()
    If IsNumeric(get_p_ReklNr()) = False Then
        MsgBox getTC_FCFromCode("TC4244"), vbInformation, getTC_FCFromCode("TC4245")
        Exit Sub
    End If
    If DCount("ID", "Mastertab", strKrit) = 0 Then
        MsgBox getTC_FCFromCode("TC0205") & get_p_ReklNr() & getTC_FCFromCode("TC0379")    'MsgBox ("Reklamation mit der Nummer: " & get_p_ReklNr () & " ist nicht existent!")   'TC0362
        GoTo end_InterneFehlererfassung4_Click
    Else
        Set ldb = openAccDBConnection("openDB")
        Set rs = ldb.OpenRecordset("Select * from Mastertab where ID = " & get_p_ReklNr() & ";")
        If rs.RecordCount Then
            If (Not IsNull(rs("ErfasserIntern").Value)) Then
                If Me.OptionReklAuswahl = 1 Then
'                    If (MsgBox("Soll der Vorgang mit der externe Fehlererfasung geöffnet werden?", vbYesNo, "Achtung") = vbYes) Then     'TC0363
'                        NewReklamationE "Bearbeiten", get_p_ReklNr ()
'                    If (MsgBox(getTC_FCFromCode("TC0358"), vbYesNo, getTC_FCFromCode("TC0408")) = vbYes) Then      'TC0362     '20190704-ROS
                        NewReklamationE "Bearbeiten", get_p_ReklNr()
'                    End If
                Else
                    If get_p_ReklNr() = "" Then
                        NewReklamationI
                    Else
                        OpenReklamationI
                    End If
                End If
            Else
                If Me.OptionReklAuswahl = 2 Then
'                    If (MsgBox("Die externe Reklamation kann nicht über die intern Fehlererfassung bearbeitet werden. Soll der Vorgang mit der externe Fehlererfasung geöffnet werden?", vbYesNo, "Achtung") = vbYes) Then
'                        NewReklamationE "Bearbeiten", get_p_ReklNr ()     'TC0351
'                    If (MsgBox(getTC_FCFromCode("TC0351"), vbYesNo, getTC_FCFromCode("TC0408")) = vbYes) Then       '20190704-ROS
                        NewReklamationE "Bearbeiten", get_p_ReklNr()
'                    End If
                Else
                    NewReklamationE "Bearbeiten", get_p_ReklNr()
                End If
            End If
        End If
        Set rs = Nothing
         '' Set ldb = openAccDBConnection("closeDB")
    End If
Else

    If get_ServiceZuordnung = True Then '20161129-ROS
        If Me.OptionReklAuswahl = 1 Then 'Extern
            NewReklamationE "Erfassen", ""
        Else
            MsgBox getTC_FCFromCode("TC0212"), vbOKOnly  'Die interne Fehlererfassung ist aktuell nur über den Thin-Client möglich!", vbOKOnly
            'NewReklamationI
        End If
    End If


End If
end_InterneFehlererfassung4_Click:

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown KeyCode, Shift, Me

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyUp KeyCode, Shift

End Sub


Private Sub cboReiterÜbersichten_Click()

    Select Case cboReiterÜbersichten
        Case 9      '"Sachbearbeiter"
            Call OpenFormx_UebersichtSachbeabeiter
        Case 10     '"Bearbeiter"
            Call OpenFormx_UebersichtBeabeiter
        Case 11     '"Reklamationen Übersicht
            Call OpenForm_Reklamationen_Übersicht
        Case 12     '"Übersicht Mahnwesen Lieferant"
            Call UebersichtMahnwesenLief
        Case 13     '"Übersicht Mahnwesen QSV"
            Call UebersichtMahnwesenQSV
        Case 14     '"Übersicht TageRueklieferung"
            Call Uebersicht_RueklTage
        Case 15    '"ExportTageRueklieferung"
            Call Export_Query_RueklTage
        Case 16     '"Lieferanten"
            Call OpenFormx_UebersichtLieferanten
        Case 17     '"Kunden/Ansprechpartner"
            Call OpenFormx_UebersichtKunden
        Case 30     '"Ticket Ubersicht"
            Call OpenFormx_frmTicketListe
        Case 33     '"Notizen Ubersicht"
            Call OpenFormx_frmÜbersichtNotizen
        Case 34     '"Konfig Soft Key"
            Call OpenFormx_frmKonfiguartionSoftkey
        Case 35     '"Persönlicher Filter"
            Call OpenFormx_frmPersönlicherFilter
        Case 39     '"Termine Stellungnhame"
            Call OpenFormx_frm_TerminStellungnahmeKundeLieferant
        Case 41     '"Bestand Service-Lager"
            Call OpenFormx_frm_BestandServiceLager
        Case 42     '"Offene Ab´s Service"
            Call OpenFormx_frmAB_Temp
       End Select

    Me.cboReiterÜbersichten = ""
    Me.InterneFehlererfassung4.SetFocus

End Sub
Private Sub cboRibbonOn_Click()

Dim ldb As DAO.Database
Dim rs As DAO.Recordset


If get_AdminZuordnung = True Then
    'Init Ribbon-Button
    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select [Value] from KonfigurationReklDBLocal where Konfiguration = 'RibbonOnOff';")
    If rs("Value") = "An" Then
'        Me.cboRibbonOn.Caption = "Ribbon on"
        DoCmd.ShowToolbar "Ribbon", acToolbarNo
        Call Application.SetOption("Show Status Bar", False)
        ldb.Execute ("Update KonfigurationReklDBLocal set [Value] = 'Aus' where [Konfiguration] = 'RibbonOnOff';")
    Else
        DoCmd.ShowToolbar "Ribbon", acToolbarYes
        Call Application.SetOption("Show Status Bar", True)
'        Me.cboRibbonOn.Caption = "Ribbon off"
        ldb.Execute ("Update KonfigurationReklDBLocal set [Value] = 'An' where [Konfiguration] = 'RibbonOnOff';")
    End If
    Set rs = Nothing
End If

End Sub

Private Sub VorgangsBericht_Click()
On Error GoTo Err_VorgangsBericht_Click


    DoCmd.RunCommand acCmdSelectRecord
    DoCmd.PrintOut acSelection

Exit_VorgangsBericht_Click:
    Exit Sub

Err_VorgangsBericht_Click:
    MsgBox Err.Description
    Resume Exit_VorgangsBericht_Click

End Sub

Sub MessageBox_zeitgesteuert()

'Dim DeinMsg As Object
'Dim sec As Long
'
'sec = 9
'
'Set DeinMsg = CreateObject("WScript.Shell")
'
'    DeinMsg.PopUp "Bitte alle Änderungen speichen und Datenbank-Client beenden." & _
'                  "Es müssen Wartungsarbeiten an der Datenbank durchgeführt werden." & _
'                  "Der Datenbank-Client wird in ca 5min. automatisch heruntergefahren!", sec, _
'                  "Achtung!-Achtung!-Achtung!-Achtung!-Achtung!-Achtung!-Achtung!-Achtung!", 48

End Sub

Sub DatenbankCopy()

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim Batch_Pfad As String
Dim Datum As String

Datum = Date
Batch_Pfad = "\\link\dfs\QM_Reklamationen_Datenbank\Access-Datenbank\Reklamationen Daten\Datenbank_Copy.bat"

    If get_ServiceZuordnung = True Then
        Set ldb = openAccDBConnection("openDB")
        Set rs = ldb.OpenRecordset("Select * from KonfigurationReklDB where Konfiguration = 'DatenbankCopy';")
        If rs("Value") = Datum Then
            Set rs = Nothing 'Nothing
            Exit Sub
        Else
            rs.Edit
            rs("Value") = Date
            rs.Update
'            Shell Batch_Pfad, 1
        End If
        Set rs = Nothing
    End If

End Sub


Private Sub OptEntwicklungAuswahl_AfterUpdate()

Dim ldb As DAO.Database
Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from KonfigurationReklDBLocal where Konfiguration = 'Entwicklung';")

    Select Case OptEntwicklungAuswahl
             Case 1
                rs.Edit
                rs("Value") = "Ja"
                rs.Update
             Case 2
                rs.Edit
                rs("Value") = "Nein"
                rs.Update
'                Set rs = getOpenRecordset("Select * from tblVerweise")
'                    rs.MoveFirst
'                    Do Until rs.EOF
'                    rs.Edit
'                    rs("VerweisName") = ""
'                    rs.Update
'                    rs.MoveNext
'                Loop
'                Application.References.Remove Application.References("PDFCreator")
         End Select
    Set rs = Nothing 'Nothing

End Sub

Private Sub OptionMeldungsnummer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CursorWechseln(IDC_HAND)
End Sub

Private Sub OptionReklAuswahl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CursorWechseln(IDC_HAND)
End Sub

Private Sub OptionServiceMA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call CursorWechseln(IDC_HAND)
End Sub
