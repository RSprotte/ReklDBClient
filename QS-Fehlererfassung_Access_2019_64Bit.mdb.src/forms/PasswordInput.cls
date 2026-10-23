Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Befehl2_Click()
On Error GoTo Err_Befehl2_Click


    Screen.PreviousControl.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_Befehl2_Click:
    Exit Sub

Err_Befehl2_Click:
    MsgBox Err.Description
    Resume Exit_Befehl2_Click

End Sub


Private Sub Form_Open(Cancel As Integer)


    If get_AdminZuordnung = True Then     '20170125-ROS
            Me.Password = getPWDBAdmin()
            Password_AfterUpdate
    End If

End Sub

Private Sub Password_AfterUpdate()

    Dim ll_KundeID As Long
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from KonfigurationReklDB where Konfiguration = 'AdminPW';")

    If Me.Password = getPWDBAdmin() Then
        Select Case sCurrentDlg
            Case "Vorgang Status Zuordnung"
                DoCmd.OpenForm "ZuordnungFolgeStatus"
            Case "AppendDef"
                Append_TableDefs2
            Case "ToggleDBWin"
                ToggleDB_Win2
            Case "DelReklamation"
                DelReklVorgang2
            Case "ConnectASA"
                ConnectASA
            Case "KonfigReklDB"
                KonfReklDBEnvir2
            Case "DelAnsprechpartnerKunde"
                 Set ldb = openAccDBConnection("openDB")
                ll_KundeID = Forms![KundenUebersicht]!KundeAnsprechpartnerÜbersicht![Kunden-ID]
                ldb.Execute ("Delete * from Kunden where [Kunden-ID] = " & Forms![KundenUebersicht]!KundeAnsprechpartnerÜbersicht![Kunden-ID] & ";")
                ldb.Execute ("Update Mastertab set [Kunden-ID] = Null where [Kunden-ID] = " & ll_KundeID & ";")
                Forms![KundenUebersicht].Requery
                Forms![KundenUebersicht].Refresh
            Case "AdminLQFehlerortbestimmung"
                DoCmd.OpenForm "AdminLQFehlerortbestimmung", acNormal
            Case "ZuordnungFehlerbildWSAnalyse"
                DoCmd.OpenForm "ZuordnungEinzeltaetigkeitenZurWS", acNormal
            Case "ZuordnungFehlerbildWSErfasser"
                DoCmd.OpenForm "ZuordnungFehlerbildWS_Meldersicht", acNormal
            Case "KonfigTaetigkeiten"
                DoCmd.OpenForm "Taetigkeiten", acNormal
            Case "ToggleVBAEditor"
                ToggleVBA_Editor2
            Case "Bearbeiter"
                DoCmd.OpenForm "Bearbeiter", acNormal
            Case "Konfig Ticket-Typ"
                DoCmd.OpenForm "frmTicketTyp"
            Case "ZuordnungKürzelZuTicket"
                DoCmd.OpenForm "frmZuordnungKürzelZuTicket"
            Case "ReklBaerbZeit"
                DoCmd.OpenForm "frmReklBearbZeit"
            Case "KonfigMsgBox"
                DoCmd.OpenForm "frmMsgBoxKonfig"
            Case "ZuornungAVISTAWsReklDBWs"
                DoCmd.OpenForm ("frmAccessWsID_AvistaWsID")
            Case "frmInfoTagKunde"
                DoCmd.OpenForm ("frmInfoTagKunde")
            Case "frmEmailServiceDaten"
                DoCmd.OpenForm ("frmEmailServiceDaten")
            Case "frmRSP_Lokalisierung"
                DoCmd.OpenForm ("frmRSP_Lokalisierung")

        End Select
        DoCmd.Close acForm, "PasswordInput"

'20170125-ROS

    ElseIf Me.Password = getPWDBUser() Then
        Select Case sCurrentDlg
            Case "DelAnsprechpartnerKunde"
                 Set ldb = openAccDBConnection("openDB")
                ll_KundeID = Forms![KundenUebersicht]!KundeAnsprechpartnerÜbersicht![Kunden-ID]
                ldb.Execute ("Delete * from Kunden where [Kunden-ID] = " & Forms![KundenUebersicht]!KundeAnsprechpartnerÜbersicht![Kunden-ID] & ";")
                ldb.Execute ("Update Mastertab set [Kunden-ID] = Null where [Kunden-ID] = " & ll_KundeID & ";")
                Forms![KundenUebersicht].Requery
                Forms![KundenUebersicht].Refresh
            Case "ZuordnungFehlerbildWSAnalyse"
                DoCmd.OpenForm "ZuordnungEinzeltaetigkeitenZurWS", acNormal
            Case "ZuordnungFehlerbildWSErfasser"
                DoCmd.OpenForm "ZuordnungFehlerbildWS_Meldersicht", acNormal
            Case "KonfigTaetigkeiten"
                DoCmd.OpenForm "Taetigkeiten", acNormal
            Case "AdminLQFehlerortbestimmung"
                DoCmd.OpenForm "AdminLQFehlerortbestimmung", acNormal
            Case "ZuornungAVISTAWsReklDBWs"
                DoCmd.OpenForm ("frmAccessWsID_AvistaWsID")
            Case Else
                MsgBox getTC_FCFromCode("TC0204")  'MsgBox ("Falsches Password!") 'TC0204
        End Select
        DoCmd.Close acForm, "PasswordInput"

        Else
            MsgBox getTC_FCFromCode("TC0204")  'MsgBox ("Falsches Password!") 'TC0204
    End If

    Forms("TB:autostart").Controls("Version") = getClientVersion()
    Forms("TB:autostart").Controls("ServerMDB") = getServerLocation()
    Forms("TB:autostart").Refresh

End Sub
