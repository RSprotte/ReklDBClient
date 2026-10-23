Attribute VB_Name = "CommandBarRekl"
Option Compare Database
Public Const PW_AppendDB = "Rekl321"
Public bToggleDB_Win As Boolean
Public bToggleVBA_Editor As Boolean


Dim cmdBar As CommandBar
Dim mDatei As CommandBarPopup       'Datei
Dim mOpen As CommandBarButton       'Öffnen
Dim mClose As CommandBarButton      'Schließen
Dim mToggleDB_Win As CommandBarButton 'Aus- / Einblenden Datenbankfenster
Dim mReklamation As CommandBarPopup       'Datei
Dim mReklamationBearbeiten As CommandBarButton       'Öffnen
Dim mReklamationNeuI As CommandBarButton       'Öffnen
Dim mReklamationNeuE As CommandBarButton       'Öffnen
Dim mReklamationLöschen As CommandBarButton       'Öffnen
Dim mMeldungsnummerSuchen As CommandBarButton
Dim mFehlerarten As CommandBarPopup       'Datei
Dim mFehlerartenGrob As CommandBarButton
Dim mFehlerartenFein As CommandBarButton
Dim mMaßnahmen As CommandBarPopup       'Datei
Dim mMaßnahmenDialog As CommandBarButton       'RSP - 20131110 - Neu
Dim mMaßnahmenBearbeitungsort As CommandBarButton       'RSP - 20131110 - Neu
Dim mUebersichten As CommandBarPopup       'Datei
Dim mUebersichtenReklMa As CommandBarButton       'Öffnen
Dim mUebersichtenRekl As CommandBarButton       'Öffnen
Dim mExportTageRücklieferung As CommandBarButton       'Öffnen
Dim mRuecklieferTage As CommandBarButton       'Öffnen
Dim mKunden As CommandBarButton
Dim mLieferanten As CommandBarButton       'Öffnen
Dim mSachbearbeiter As CommandBarButton       'Öffnen
Dim mUebersichtMahnungLief As CommandBarButton
Dim mUebersichtMahnungQSV As CommandBarButton
Dim mSuchen As CommandBarButton       'Öffnen
Dim mAdmin As CommandBarPopup       'Datei
Dim mKonfigLQFehlerort As CommandBarButton       'Öffnen
Dim mTableDefAppend As CommandBarButton       'Öffnen
Dim mDelReklVorgang As CommandBarButton       'Öffnen
Dim mConnAvista As CommandBarButton
Dim mKonfigReklDBEnv As CommandBarButton
Dim mKonfigTaetigkeiten As CommandBarButton
Dim mZuordnungWSFT As CommandBarButton
Dim mZuordnungWSFTMelder As CommandBarButton
Dim mHilfe As CommandBarPopup       'Datei
Dim mReklHilfe As CommandBarButton
Dim mInfo As CommandBarButton
Dim mVBAEditorToggle As CommandBarButton
Dim mBearbeiter As CommandBarButton

Public sCurrentDlg As String

Sub Beenden()

    DoCmd.Close

End Sub

Sub OpenFormTest()

    DoCmd.OpenForm "Test"

End Sub

Public Function getLocTextecmdBar() As CommandBar

    Set getLocTextecmdBar = Application.CommandBars("Reklamationsmenue")

End Function

Sub OpenForm_FehlerbeschreibungGrob()

    DoCmd.OpenForm "FehlerartenDerFehlerklassen", acNormal
    Forms("FehlerartenDerFehlerklassen").Controls("Fehlerart").Locked = True
    Forms("FehlerartenDerFehlerklassen").UFo1.Controls("Beschreibung").Locked = True
    Forms("FehlerartenDerFehlerklassen").Controls("Übernehmen").Enabled = False

End Sub
Sub OpenForm_TicketTypKonfig()

        sCurrentDlg = "Konfig Ticket-Typ"
        DoCmd.OpenForm "PasswordInput"

End Sub
Sub OpenForm_LQFehlerortbestimmung()

    sCurrentDlg = "AdminLQFehlerortbestimmung"
    DoCmd.OpenForm "PasswordInput"

End Sub
Sub OpenForm_LQ_Taetigkeiten()

    sCurrentDlg = "KonfigTaetigkeiten"
    DoCmd.OpenForm "PasswordInput"


End Sub
Sub OpenForm_LQ_ZuordnungWSAnalyse()


    sCurrentDlg = "ZuordnungFehlerbildWSAnalyse"
    DoCmd.OpenForm "PasswordInput"


End Sub
Sub OpenForm_LQ_ZuordnungWSFTErfasser()

    sCurrentDlg = "ZuordnungFehlerbildWSErfasser"
    DoCmd.OpenForm "PasswordInput"


End Sub

Sub OpenDatenblatt_Reklamationen_ÜbersichtMa()

    Dim stDocName As String
    stDocName = "Abfrage LINK-Mitarbeiter"

    DoCmd.OpenQuery stDocName, acNormal, acEdit

End Sub
Sub OpenForm_Reklamationen_Übersicht() '20160229-ROS

    Dim Date1 As String
    Dim Schaltjahr As String

'    Date1 = "29.02."
'    Schaltjahr = Format(Date, "dd.mm.")
'    If Schaltjahr = Date1 Then
'        sVon = Format(Date - 1, "dd.mm") & "." & LTrim(Str(Val(Format(Date, "yyyy")) - 1))
'    Else
'        sVon = Format(Date, "dd.mm") & "." & LTrim(Str(Val(Format(Date, "yyyy")) - 1))
'    End If
'    sBis = Date
'    sAuftrag = " AND ([Auftragsdatum] between #" & Format(sVon, "mm\/dd\/yyyy") & "# and #" & _
'                                                Format(sBis, "mm\/dd\/yyyy") & "#)"
    Date1 = "02/29/"
    Schaltjahr = format(Date, "mm\/dd\/")
    If Schaltjahr = Date1 Then
        sVon = format(Date - 1, "mm\/dd\/") & LTrim(Str(Val(format(Date, "yyyy")) - 1))
    Else
        sVon = format(Date, "mm\/dd\/") & LTrim(Str(Val(format(Date, "yyyy")) - 1))
    End If
    sBis = Date
    sAuftrag = " AND ([Auftragsdatum] between #" & format(sVon, "mm\/dd\/yyyy") & "# and #" & _
                                                format(sBis, "mm\/dd\/yyyy") & "#)"

    Location = DLookup("Value", "KonfigurationReklDBLocal", "Konfiguration = 'Location'")
    If Location = "English" Then
        sBis = format(Date, "mm\/dd\/yyyy")
        sVon = format(sVon, "dd\/mm\/yyyy")
    Else
        sBis = format(Date, "dd\/mm\/yyyy")
        sVon = format(sVon, "mm\/dd\/yyyy")
    End If

    If Forms("TB:autostart").Controls("OptionReklAuswahl") = 2 Then '2 = intern
        DoCmd.OpenForm "Reklamationen_ÜbersichtIntern", acNormal
        Forms("Reklamationen_ÜbersichtIntern").Filter = "[intern] = True" & sAuftrag
        Forms("Reklamationen_ÜbersichtIntern").FilterOn = True
        Forms("Reklamationen_ÜbersichtIntern").FilterView = Forms("Reklamationen_ÜbersichtIntern").Filter
        Forms("Reklamationen_ÜbersichtIntern").Von = sVon
        Forms("Reklamationen_ÜbersichtIntern").Bis = sBis
    Else ' else = extern = 1
        DoCmd.OpenForm "Reklamationen_Übersicht", acNormal
        Forms("Reklamationen_Übersicht").Filter = "([extern] = True" & sAuftrag & ")"
        Forms("Reklamationen_Übersicht").FilterOn = True
        Forms("Reklamationen_Übersicht").FilterView = Forms("Reklamationen_Übersicht").Filter
        Forms("Reklamationen_Übersicht").Von = sVon
        Forms("Reklamationen_Übersicht").Bis = sBis
    End If


End Sub


Sub Export_Query_RueklTage()

    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, _
                 "RSP_AnzahlTageRücklieferungBisAuslieferungLief", Date & "_Export_RueckliefTage.xls", True, _
                 "Export_Übersicht"

End Sub
Sub Uebersicht_RueklTage()

     DoCmd.OpenForm "RSP_Übersicht_RueckLiefTage"

End Sub
Sub OpenFormx_frmAB_Temp()

    DoCmd.OpenForm "frmAB_Temp"

End Sub

Sub OpenFormx_frm_BestandServiceLager()

    DoCmd.OpenForm "frmBestandServiceLager"

End Sub
Sub OpenFormx_frm_TerminStellungnahmeKundeLieferant()

    DoCmd.OpenForm "frm_TerminStellungnahmeKundeLieferant"

End Sub

Sub OpenFormx_UebersichtLieferanten()

DoCmd.OpenForm "LieferantenUebersicht"

End Sub

Sub OpenFormx_UebersichtKunden()

DoCmd.OpenForm "KundenUebersicht"

End Sub
Sub OpenFormx_UebersichtSachbeabeiter()

    DoCmd.OpenForm "SachbearbeiterUebersicht"

End Sub
Sub OpenFormx_frmTicketListe()

    DoCmd.OpenForm "frmTicketListe", , , , , , "Liste komplett"
    Forms("frmTicketListe").Controls("txtReklID") = ""

End Sub
Sub OpenFormx_frmÜbersichtNotizen()

    If get_ServiceZuordnung = True Then
        DoCmd.OpenForm "frmÜbersichtNotizen"
    Else
        Exit Sub
    End If

End Sub
Sub OpenFormx_UebersichtBeabeiter()

    If get_ServiceZuordnung = True Then
        DoCmd.OpenForm "Bearbeiter"
    Else
        sCurrentDlg = "Bearbeiter"
        DoCmd.OpenForm "PasswordInput"
    End If

End Sub
Sub Append_TableDefs()

    sCurrentDlg = "AppendDef"
    DoCmd.OpenForm "PasswordInput"

End Sub

Public Sub Append_TableDefs2()

    Dim dbSource As DAO.Database
    Dim dbDestination As DAO.Database

    Dim tdf As DAO.TableDef
    Dim tds As DAO.TableDef

    Dim sDatei As String
    Dim sDropTable As String

    On Error GoTo Append_TableDefs2_Err



    sDatei = DateiOeffnen()
    If sDatei = "" Then
'     sDatei = InputBox("Eingabe Datenbankpfad + Datenbank", "Eingabe Datenbank")
'     sDatei = getValueFromInputBox("Eingabe Datenbankpfad + Datenbank", "Eingabe Datenbank")
     sDatei = getValueFromInputBox(getTC_FCFromCode("IC4285"), getTC_FCFromCode("IC4286"))
    End If
    'Set dbSource = DBEngine.Workspaces(0).OpenDatabase("C:\Users\ralf sprotte\Desktop\QS-Fehlererfassung_Daten.mdb")
    Set dbSource = DBEngine.Workspaces(0).OpenDatabase(sDatei)

    Set dbDestination = CurrentDb

    For Each tdf In dbDestination.TableDefs
        If tdf.Connect <> "" Then
            sDropTable = "DROP TABLE " & tdf.Name
            dbDestination.Execute "DROP TABLE [" & tdf.Name & "]", dbFailOnError
        End If
    Next tdf

    For Each tds In dbSource.TableDefs
        If Not Left(tds.Name, 5) = "MySys" Then
            Set tdf = dbDestination.CreateTableDef(tds.Name)
            tdf.Connect = ";DATABASE=" & sDatei
            tdf.SourceTableName = tds.Name
            dbDestination.TableDefs.Append tdf
        End If
    Next tds


    Forms("TB:autostart").Controls("Version") = getClientVersion()
    Forms("TB:autostart").Controls("ServerMDB") = getServerLocation()
    Forms("TB:autostart").Refresh

Append_TableDefs2_Exit:

    dbSource.Close
    Set dbSource = Nothing
    Set dbDestination = Nothing
    Set tdf = Nothing

    Exit Sub

Append_TableDefs2_Err:

    If Err = 3211 Or Err = 3033 Or Err = 3012 Then
        Resume Next
    Else
        MsgBox "Error: " & Err & " - " & Error
        Resume Next
    End If

End Sub
Sub VorgangStatusZuordnung()

    sCurrentDlg = "Vorgang Status Zuordnung"
    DoCmd.OpenForm "PasswordInput"

End Sub
Sub DelReklVorgang()

    sCurrentDlg = "DelReklamation"
    DoCmd.OpenForm "PasswordInput"

End Sub


Public Sub DelReklVorgang2()

On Error GoTo DelReklVorgang2_Err
    Dim ldb As DAO.Database
    Dim strKrit As String

'    get_p_ReklNr () = InputBox("Geben Sie die Reklamationsnummer ein!", "Eingabe Vorgang")    'IC0200
'    p_ReklNr = InputBox(getTC_FCFromCode("IC0200"), getTC_FCFromCode("IC0219"))  '20160729-ROS
    p_ReklNr = getValueFromInputBox(getTC_FCFromCode("IC0200"), getTC_FCFromCode("IC0219"))
    If StrPtr(p_ReklNr) = 0 Then   ' Prüfen ob Abbrechen gedrückt wurde
        GoTo end_DelReklVorgang2
    End If
    strKrit = "ID = " & get_p_ReklNr()
    If DCount("ID", "Mastertab", strKrit) = 0 Then
        MsgBox getTC_FCFromCode("TC0205" & get_p_ReklNr() & "TC0379")    'MsgBox ("Reklamation mit der Nummer: " & get_p_ReklNr () & " ist nicht existent!")   'TC0205
        GoTo end_DelReklVorgang2
    Else
        Set ldb = openAccDBConnection("openDB")
        ldb.Execute ("Delete * from Mastertab where [ID] = " & get_p_ReklNr() & ";")
        ldb.Execute ("Delete * from [Mastertab-DSBerechtigung] where [ID] = " & get_p_ReklNr() & ";")
        ldb.Execute ("Delete * from ReklMemo where [ID] = " & get_p_ReklNr() & ";")
        ldb.Execute ("Delete * from tblTicketListe where [ReklID] = " & get_p_ReklNr() & ";")    '20240822-ROS
         ' Set ldb = openAccDBConnection("closeDB")
    End If

DelReklVorgang2_Err:
end_DelReklVorgang2:

End Sub

Public Function CheckAdminPW() As Boolean

'    If Not InputBox(getTC_FCFromCode("TC0430") & ": ", "Append DB") = getPWDBAdmin() Then
    If Not getValueFromInputBox(getTC_FCFromCode("TC0430"), "Append DB") = getPWDBAdmin() Then
        MsgBox getTC_FCFromCode("TC0204")  'MsgBox ("Falsches Password!")  'TC0206
        CheckAdminPW = False
    Else
        CheckAdminPW = True
    End If

End Function

Sub OpenDialogMassnahmen()
'RSP - 20131110 - Neu
DoCmd.OpenForm ("Maßnahmen")


End Sub

Public Sub ToggleDB_Win()

'    Select Case fOSUserName  '20160729-ROS
'        Case "ralf sprotte", "roberto santarelli"
'            Call ToggleDB_Win2
'            Case Else
'                sCurrentDlg = "ToggleDBWin"
'                DoCmd.OpenForm "PasswordInput"
'    End Select

    If get_AdminZuordnung = True Then   '20170125-ROS
        Call ToggleDB_Win2
    Else
        sCurrentDlg = "ToggleDBWin"
        DoCmd.OpenForm "PasswordInput"
    End If


End Sub

Public Sub ToggleVBA_Editor()

'    Select Case fOSUserName  '20160729-ROS
'        Case "ralf sprotte", "roberto santarelli"
'            Call ToggleVBA_Editor2
'            Case Else
'                sCurrentDlg = "ToggleVBAEditor"
'                DoCmd.OpenForm "PasswordInput"
'    End Select

    If get_AdminZuordnung = True Then   '20170125-ROS
        Call ToggleVBA_Editor2
    Else
        sCurrentDlg = "ToggleVBAEditor"
        DoCmd.OpenForm "PasswordInput"
    End If


End Sub

Public Sub ToggleDB_Win2()


    If bToggleDB_Win Then 'Ausgeblendet
        'einblenden
        CommandBars.ActiveMenuBar.Enabled = True
        CommandBars("Database").Visible = True
        DoCmd.SelectObject acTable, , True
        bToggleDB_Win = False
    Else
        'ausblenden
        CommandBars.ActiveMenuBar.Enabled = False
        CommandBars("Database").Visible = False
        DoCmd.SelectObject acTable, , True
        RunCommand acCmdWindowHide
        bToggleDB_Win = True
    End If


End Sub
Public Sub ToggleVBA_Editor2()


    If bToggleVBA_Editor Then 'Ausgeblendet
        'einblenden
        With Application.VBE.MainWindow
            .SetFocus
            .Visible = True
        End With
        bToggleVBA_Editor = False
    Else
        'ausblenden
        With Application.VBE.MainWindow
            .SetFocus
            .Visible = False
        End With
        bToggleVBA_Editor = True
    End If


End Sub

Sub DBSchließen()

DoCmd.Quit

End Sub
Sub ConnectASA1()

    sCurrentDlg = "ConnectASA"
    DoCmd.OpenForm "PasswordInput"

End Sub


Public Sub ConnectASA()
On Error GoTo err_ConnectASA_DB_ADO_Click

   Dim cnn As ADODB.Connection
   Dim rs As ADODB.Recordset


   Set cnn = New ADODB.Connection

   ' Open a connection using an ODBC DSN "Pubs".
   'cnn.ConnectionString = "DSN=av_test_he;UID=dba;PWD=phoenix;"
   cnn.ConnectionString = "DSN=avista;UID=dba;PWD=phoenix;"
   cnn.Open

   ' Find out if the attempt to connect worked.
   If cnn.State = adStateOpen Then
      MsgBox "Welcome to Avista HE!"
   Else
      MsgBox "Sorry. No Avista HE today."
      GoTo close_db
   End If
   'ut_qsm_reklamation.ticket_nummer
   ' Create a Recordset by executing an SQL statement.
   Set rs = cnn.Execute(AvistaArtikelSQL_Select & "1100077' " & AvistaArtikelSQL_Order)
    'SELECT tm_artikel.ast_nummer, tm_artikel_var.avar_nummer, tm_artikelvar_kd.avkd_artnr_kd, tm_artikel.ast_bezeichnung, va_produktgruppe.pg_bez
   ' Show the first author.
   rs.MoveFirst
   While Not rs.EOF
   MsgBox ("ArtikelNr: " & rs("ast_nummer") & "-" & rs("avar_nummer") & Chr(13) & Chr(10) & _
           "Knden-Nr.: " & rs("avkd_artnr_kd") & Chr(13) & Chr(10) & _
           "Bezeichnung: " & rs("ast_bezeichnung") & Chr(13) & Chr(10) & _
           "Produktgruppe: " & rs("pg_bez"))
           rs.MoveNext
   Wend
   ' Close the connection.
   rs.Close
   Set rs = Nothing

close_db:
   ' Close the connection.
   cnn.Close
   Exit Sub

err_ConnectASA_DB_ADO_Click:
    MsgBox "error: " & Err & " - " & Error
    Resume Next

End Sub

Sub OpenReklamation()

Dim ldb As DAO.Database
Dim rs As DAO.Recordset

'get_p_ReklNr () = InputBox("Geben Sie die Reklamationsnummer ein!", "Eingabe Vorgang")    'IC0202
'p_ReklNr = InputBox(getTC_FCFromCode("IC0200"), getTC_FCFromCode("IC0219"))
p_ReklNr = getValueFromInputBox(getTC_FCFromCode("IC0200"), getTC_FCFromCode("IC0219"))
If StrPtr(p_ReklNr) = 0 Then   ' Prüfen ob Abbrechen gedrückt wurde
    GoTo end_OpenReklamation
End If
strKrit = "ID = " & get_p_ReklNr()
If DCount("ID", "Mastertab", strKrit) = 0 Then
    MsgBox getTC_FCFromCode("TC0205") & get_p_ReklNr() & getTC_FCFromCode("TC0379")    'MsgBox ("Reklamation mit der Nummer: " & get_p_ReklNr () & " ist nicht existent!") 'TC0209
    GoTo end_OpenReklamation
Else
     Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from Mastertab where ID = " & get_p_ReklNr() & ";")
    If rs.RecordCount Then
        If (Not IsNull(rs("ErfasserIntern").Value)) Then
            OpenReklamationI
        Else
            NewReklamationE "Bearbeiten", get_p_ReklNr()
        End If
    End If
    Set rs = Nothing
     ' Set ldb = openAccDBConnection("closeDB")
End If
end_OpenReklamation:

End Sub
Public Function setKuerzel() As String

    setKuerzel = get_pKuerzel()

End Function

Public Function getKundenIDfromKuerzel(k As String) As Long
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

     Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from Kunden where [Kurzzeichen] = '" & k & "';", dbOpenSnapshot)
    If rs.RecordCount Then
        pi_Kuerzel = rs("Kunden-ID")
        getKundenIDfromKuerzel = rs("Kunden-ID")
    Else
        getKundenIDfromKuerzel = 0
    End If

    Set rs = Nothing
     ' Set ldb = openAccDBConnection("closeDB")

End Function
Public Function getSachbearbeiterIDfromKuerzel(k As String) As Long
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from Sachbearbeiter where [Kurzzeichen] = '" & k & "';", dbOpenSnapshot)
    If rs.RecordCount Then
        pi_Kuerzel = rs("SB-ID")
        getSachbearbeiterIDfromKuerzel = rs("SB-ID")
    Else
        getSachbearbeiterIDfromKuerzel = Null
    End If

    Set rs = Nothing
     ' Set ldb = openAccDBConnection("closeDB")

End Function

Sub ReklamationE()

    NewReklamationE "Erfassen", ""

End Sub
Sub ReklamationI()

    MsgBox getTC_FCFromCode("TC0212"), vbOKOnly  'Die interne Fehlererfassung ist aktuell nur über den Thin-Client möglich!", vbOKOnly
    'NewReklamationI

End Sub


Sub KonfReklDBEnvir()

    sCurrentDlg = "KonfigReklDB"
    DoCmd.OpenForm "PasswordInput"

End Sub

Public Sub KonfReklDBEnvir2()

    DoCmd.OpenForm "KonfiguartionLQReklDB"

End Sub
Sub OpenReklamation2() '20160613-ROS
On Error GoTo err_RosTest

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim strKrit As String
Dim sÜbersicht As String
Dim QS As DAO.QueryDef
Dim Mnrku As String


    'MnrKu = InputBox("Geben Sie die Meldungsnummer ein!", "Eingabe Vorgang")    'IC0204
'    Mnrku = InputBox(getTC_FCFromCode("IC0204"), getTC_FCFromCode("IC0219"))
    Mnrku = getValueFromInputBox(getTC_FCFromCode("IC0204"), getTC_FCFromCode("IC0219"))
'    If StrPtr(Mnrku) = 0 Then ' Prüfen ob Abbrechen gedrückt wurde
'        GoTo end_OpenReklamation2
'    End If
    If Mnrku = "0" Then  ' Prüfen ob Abbrechen gedrückt wurde
        GoTo end_OpenReklamation2
    End If

    If Mnrku <> "" Then
    strKrit = "[Meldungsnummer Kunde] = '" & Mnrku & "'"
    If DCount("[Meldungsnummer Kunde]", "Mastertab", strKrit) = 0 Then
        MsgBox getTC_FCFromCode("TC0205") & " - " & Mnrku & " - " & getTC_FCFromCode("TC0379")
        'MsgBox ("Reklamation mit der Meldungsnummer: " & MnrKu & " ist nicht existent!") 'TC0213
        GoTo end_OpenReklamation2
    Else
        Set ldb = openAccDBConnection("openDB")
        Set rs = ldb.OpenRecordset("Select * from Mastertab where [Meldungsnummer Kunde] = '" & Mnrku & "' and ([Referenznummer] Is Null)")
            rs.MoveLast
            rs.MoveFirst
        If rs.RecordCount = 1 Then
            p_ReklNr = rs("ID")
            NewReklamationE "Bearbeiten", get_p_ReklNr()
           Else

            DoCmd.OpenForm "Reklamationen_Übersicht_Konfigurator"

                sÜbersicht = "Reklamationen_Übersicht"
                bFilterOn = False

                sfilter = "([Meldungsnummer Kunde] like '" & Mnrku & "*')"
                bFilterOn = True

                If Not IsLoaded(sÜbersicht) Then
                    DoCmd.OpenForm sÜbersicht
                End If
                'Filter in Übersicht setzen und aktualisieren
                Forms(sÜbersicht).Controls("FilterView") = sfilter
                Forms(sÜbersicht).Filter = sfilter
                Forms(sÜbersicht).FilterOn = True
                Forms(sÜbersicht).Requery
                Forms(sÜbersicht).Refresh

            Set QS = ldb.QueryDefs("RSP_Ausw_Basisdaten_1_Export")
                If Not bFilterOn Then
                    If InStr(QS.sql, "Where") Then
                        QS.sql = Left(QS.sql, InStr(QS.sql, "Where") - 1) & sfilter & ";"
                    End If
                    Else
                    If InStr(QS.sql, "Where") Then
                        QS.sql = Left(QS.sql, InStr(QS.sql, "Where") + 5) & sfilter & ";"
                    Else
                        QS.sql = Left(QS.sql, InStr(QS.sql, ";") - 1) & " Where " & sfilter & ";"
                    End If
                End If

exit_RosTest:
                Set rs = Nothing

                Forms("Reklamationen_Übersicht_Konfigurator").SetFocus
                DoCmd.Close acForm, "Reklamationen_Übersicht_Konfigurator"

                Forms(sÜbersicht).SetFocus
                Exit Sub

err_RosTest:
                If Err = 3021 Then
                    Resume Next
                Else
                    GoTo exit_RosTest
                End If
            End If
        End If
Set rs = Nothing
    End If
end_OpenReklamation2:


End Sub

Sub UebersichtMahnwesenLief()

    DoCmd.OpenForm "Übersicht_Mahnwesen_Lief"

End Sub
Sub UebersichtMahnwesenQSV()

    DoCmd.OpenForm "Übersicht_Mahnwesen_QSV"

End Sub
Sub ZuordnungKürzelZuTicket()

    sCurrentDlg = "ZuordnungKürzelZuTicket"
    DoCmd.OpenForm "PasswordInput"

End Sub
Sub ZeiterfassungÖffnen()

    DoCmd.OpenForm ("frmZeiterfassung")

End Sub

Sub ZeiterfassungROSÖffnen()

    DoCmd.OpenForm ("frmZeiterfassungROS")

End Sub
Sub OpenFormx_frmKonfiguartionSoftkey()

    DoCmd.OpenForm ("frmKonfiguartionSoftkey")

End Sub

Sub RücklieferungKunde()
'20210323-ROS: Code-Änderung zum öffnen eiens Formulars für die bearbeitung der Rücklieferungen

'    DoCmd.OpenReport "Rücklieferung_Kunde", acViewReport
'    If MsgBox("Soll der Bericht ausgedruckt werden?", vbYesNo) = vbYes Then
'        DoCmd.OpenReport "Rücklieferung_Kunde", acViewNormal
'     Else
'        Exit Sub
'     End If

    DoCmd.OpenForm ("frmRücklieferung_Kunde")
'    If MsgBox(getTC_FCFromCode("TC4214"), vbYesNo) = vbYes Then
'        Forms("frmRücklieferung_Kunde").Printer.Orientation = acPRORLandscape
'        DoCmd.PrintOut acPrintAll
'     Else
'        Exit Sub
'     End If

End Sub
Sub OpenFormx_frmPersönlicherFilter()

    DoCmd.OpenForm ("frmPersönlicherFilter")

End Sub
Sub OpenQueryx_GesperrteVorgänge()

    DoCmd.OpenForm "frmGesperrteVorgänge"

End Sub
Sub ReklBaerbZeit()

    sCurrentDlg = "ReklBaerbZeit"
    DoCmd.OpenForm "PasswordInput"

End Sub
Sub OpenFormx_frmFrachtkostenZuordnung()

    If get_ServiceZuordnung = True Then DoCmd.OpenForm ("frmFrachtkostenZuordnung")

End Sub

Sub KonfigMsgBox()

    sCurrentDlg = "KonfigMsgBox"
    DoCmd.OpenForm "PasswordInput"

End Sub
Sub OpenFormx_frmAccessWsID_AvistaWsID()

    sCurrentDlg = "ZuornungAVISTAWsReklDBWs"
    DoCmd.OpenForm "PasswordInput"

End Sub
Sub OpenFormx_frmEmailService()

    DoCmd.OpenForm "frmEmailService"

End Sub

Sub OpenFormx_frmInfoTagKunde()

    sCurrentDlg = "frmInfoTagKunde"
    DoCmd.OpenForm "PasswordInput"

End Sub
Sub OpenFormx_frmEmailService_Daten()

    sCurrentDlg = "frmEmailServiceDaten"
    DoCmd.OpenForm "PasswordInput"

End Sub

Sub OpenFormx_frmRSP_Lokalisierung()

    sCurrentDlg = "frmRSP_Lokalisierung"
    DoCmd.OpenForm "PasswordInput"

End Sub
