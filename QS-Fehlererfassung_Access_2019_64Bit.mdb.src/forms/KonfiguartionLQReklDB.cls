Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const cExportUeSQuery = "SELECT [RSP_UeS_XXX].Anzeige, [RSP_UeS_XXX].AnzeigeD, [RSP_UeS_XXX].LängeBez, [RSP_UeS_XXX].AnzeigeE, [RSP_UeS_XXX].AnzeigeCN " & _
                        "From [RSP_UeS_XXX] WHERE [RSP_UeS_XXX].Geändert=True;"
Const cChangeTableItemE = "UPDATE YYY, RSP_Lokalisierung SET YYY.Anzeige = [RSP_Lokalisierung].[TextE] " & _
                        "WHERE (((YYY.XXX)=[RSP_Lokalisierung].[BasisText]));"

Const cUpdateUeSBezUnicode = "UPDATE RSP_BezUnicode, RSP_Export_BezUnicode_TempImp SET RSP_BezUnicode.BezGerman = [RSP_Export_BezUnicode_TempImp].[BezGerman], " & _
                             "RSP_BezUnicode.BezEnglish = [RSP_Export_BezUnicode_TempImp].[BezEnglish], RSP_BezUnicode.BezChinese = [RSP_Export_BezUnicode_TempImp].[BezChinese] " & _
                             "WHERE RSP_BezUnicode.Formular=[RSP_Export_BezUnicode_TempImp].[Formular] AND RSP_BezUnicode.Bezeichnungsfeld)=[RSP_Export_BezUnicode_TempImp].[Bezeichnungsfeld] " & _
                             "AND RSP_BezUnicode.TypeOfObject=[RSP_Export_BezUnicode_TempImp].[TypeOfObject];"

Private Sub Befehl64_Click()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from RSP_Lokalisierung where OjectID_Name = '" & KonfigLQReklDBDatei & "';")
        rs.Edit
        rs("BasisText") = Me.[Pfad-DE]
        rs.Update
    Set rs = Nothing

End Sub

Private Sub Befehl70_Click()
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim Datei As String

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from RSP_Lokalisierung where OjectID_Name = '" & KonfigLQReklDBDatei & "';")
        rs.Edit
        rs("TextE") = Me.[Pfad-E]
        rs.Update
    Set rs = Nothing
End Sub

Private Sub Befehl71_Click()
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim Datei As String

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from RSP_Lokalisierung where OjectID_Name = '" & KonfigLQReklDBDatei & "';")
        rs.Edit
        rs("TextCN") = Me.[Pfad-CN]
        rs.Update
    Set rs = Nothing
End Sub

Private Sub ButAppendDB_Click()

    Append_TableDefs
    'neueMLeiste

End Sub

Private Sub ButMahnungLief_Click()
    CheckAndSendMahnungenLieferant
End Sub

Private Sub ButMahnungQSV_Click()
    CheckAndSendMahnungenQSV
End Sub

Private Sub ButMahnwesenMngm_Click()

    Dim i As Integer
    i = MahnwesenAusfuehren()

End Sub

Private Sub ButSpracheUmstellen_Click()
On Error GoTo err_ButSpracheUmstellen_Click

    Dim ldb As DAO.Database
    Dim sCaption As String
    Dim frm As Object
    Dim Rpt As Object
    Dim ctl As Control
    Dim ctll As Control
    Dim dbSource As DAO.Database
    Dim rsImportObjFilter, _
        rsQuelle, _
        rsZiel _
    As DAO.Recordset
    Dim prop As Variant
    Dim oTab As DAO.TableDef
    Dim oCol As DAO.Field
    Dim pr As DAO.Property
    Dim Tb As TableDef
    Dim sDB, _
        sTabFieldName, _
        sLangField, _
        sSQL _
    As String
    Dim qExportUeSQuery As DAO.QueryDef

    Set ldb = openAccDBConnection("openDB")
    Set Tb = ldb.TableDefs("Versionierung")

    sDB = Right(Tb.Connect, Len(Tb.Connect) - InStr(1, Tb.Connect, "="))
'    Debug.Print sDB
    Set Tb = Nothing

    Set ldb = DBEngine.Workspaces(0).OpenDatabase(sDB)

    'Es werden alle Lokalisierungseinträge mit den Spalten TypeofObject = Foprmular, Bericht, usw. und den Namen des Objektes: z.B. Fehlererfassung - Register geöffnet
    Set rsImportObjFilter = CurrentDb.OpenRecordset("SELECT RSP_Lokalisierung.TypeOfObject, RSP_Lokalisierung.OjectID_Name " & _
                                     "From RSP_Lokalisierung " & _
                                     "WHERE [TypeOfObject] = 'Form' " & _
                                     "GROUP BY RSP_Lokalisierung.TypeOfObject, RSP_Lokalisierung.OjectID_Name;", dbOpenSnapshot)

    'Auf den bestehenden Recordset werden alle Einträge der Objekte = Formular gefiltert
    'Grund: Alle sichtbaren Texte in den Formularen müssen bei der Spachumstellung verändert werden
    rsImportObjFilter.MoveFirst
    While Not rsImportObjFilter.EOF
'        Debug.Print rsImportObjFilter("OjectID_Name")
        'Texte des Konfiguration-Dialog werden nicht verändert
        If Not rsImportObjFilter("OjectID_Name") = "KonfiguartionLQReklDB" Then
            'Aktuelles Formular wird in der Entwurfsansicht im Hintergrund geöffnet
            DoCmd.OpenForm rsImportObjFilter("OjectID_Name"), acDesign, , , , acHidden
            'Alle Formulartexte (Labels, Überschriften, Textfelder, usw.) werden für das aktuell geöffnet Formular ermittelt
            Set rsQuelle = CurrentDb.OpenRecordset("SELECT * from RSP_Lokalisierung where TypeOfObject = 'Form' and OjectID_Name = '" & rsImportObjFilter("OjectID_Name") & "';", dbOpenSnapshot)
            'Umstellung der Sprache für den Client
            Select Case Me.KombLang01
                Case "German"
                    sLangField = "BasisText"
                    CurrentDb.Execute ("Update KonfigurationReklDBLocal set [Value] = 'German' where [Konfiguration] = 'Location';")
                Case "English"
                    sLangField = "TextE"
                    CurrentDb.Execute ("Update KonfigurationReklDBLocal set [Value] = 'English' where [Konfiguration] = 'Location';")
                Case "Chinese"
                    sLangField = "TextCN"
                    CurrentDb.Execute ("Update KonfigurationReklDBLocal set [Value] = 'Chinese' where [Konfiguration] = 'Location';")
            End Select
            rsQuelle.MoveFirst
            'Umstellung der Sprache für alle ermittelten Steuerlemente
            While Not rsQuelle.EOF
                If rsImportObjFilter("OjectID_Name") = "TB:autostart" And rsQuelle("TextID") = "Bezeichnungsfeld66" Then
                    'fprintLogRSP "ButSpracheUmstellen_Click", rsQuelle("TextID")
                End If
                If Not IsNull(rsQuelle(sLangField)) Then
                    'fprintLogRSP "ButSpracheUmstellen_Click", "VorUmstelung: " & Forms(rsImportObjFilter("OjectID_Name")).Controls(rsQuelle("TextID")).Properties("Caption")
                    'Textersetzung
                    Forms(rsImportObjFilter("OjectID_Name")).Controls(rsQuelle("TextID")).Properties("Caption") = rsQuelle(sLangField)
                    'fprintLogRSP "ButSpracheUmstellen_Click", "NachUmstelung: " & Forms(rsImportObjFilter("OjectID_Name")).Controls(rsQuelle("TextID")).Properties("Caption")
                    'db.execute("Update [RSP_Lokalisierung] set [TextE] = '" & rsQuelle("")
                End If
                rsQuelle.MoveNext
            Wend
            Set rsQuelle = Nothing
            'Alle Datumsfelder des aktuellen geöffneten Formular werden sprachabhängig angepasst.
            Dim ls_rowSource As String
            For Each ctl In Forms(rsImportObjFilter("OjectID_Name")).Controls
                'In der englischen Access Version muss der Verweis auf ein Control über den
                'Objektname [Forms] erfolgen und nicht über [Formulare] (deutsche Version),
                'daher muss in allen Steuerlementen List- und Comboboxen in allen Formularen der  Datenherkunft
                'in die englische Version gewandelt werden
                '20191014-ROS
                If ctl.ControlType = 111 Or ctl.ControlType = 110 Then
                    If ctl.RowSource <> "" Then
                        ls_rowSource = ctl.RowSource
                        'Debug.Print "Form: " & rsImportObjFilter("OjectID_Name") & " - Ctrl-Name: " & ctl.Name & " - RowSource: " & ls_rowSource
                        If InStr(ls_rowSource, "[Formulare]") Then
                            ctl.RowSource = Replace(ctl.RowSource, "[Formulare]", "[Forms]")
                        End If
                    End If
                End If


                If ctl.ControlType = 109 Then
                    If ctl.Properties("Format") = "Short Date" Then
                        If getCurrLanguage() = "English" Then
                            ctl.Properties("InputMask") = "00\/00\/00;0;*"
                        Else
                            'Alle anderen: Deutsch / Chinesisch
                            ctl.Properties("InputMask") = "00,00,00;0;*"
                        End If
                    End If
                End If
            Next ctl

            'Formular wird geschlossen und gesichert
            DoCmd.Close acForm, rsImportObjFilter("OjectID_Name"), acSaveYes
        End If
next_rsImportObjFilter:
        rsImportObjFilter.MoveNext
    Wend


    Set rsImportObjFilter = Nothing

Call CmbTBautostartöffnen_Click

     Exit Sub
err_ButSpracheUmstellen_Click:

    If Err = 2465 Then
        ldb.Execute ("delete * from RSP_Lokalisierung where [Text-ID]=" & rsQuelle("Text-ID") & ";")
        'rsQuelle.Delete
        Resume Next
    ElseIf Err = 3078 Then
        Resume Next
    ElseIf Err = 2102 Then
        ldb.Execute ("delete * from RSP_Lokalisierung where [OjectID_Name]='" & rsImportObjFilter("OjectID_Name") & "';")
        Resume next_rsImportObjFilter
    End If

End Sub


Private Sub ButSpracheUmstellenDB_Click()
On Error GoTo err_ButSpracheUmstellenDB_Click

    Dim ldb As DAO.Database
    Dim sCaption As String
    Dim frm As Object
    Dim Rpt As Object
    Dim ctl As Control
    Dim ctll As Control
    Dim dbSource _
    As DAO.Database
    Dim rsImportObj As DAO.Recordset
    Dim rsImportObjFilter As DAO.Recordset
    Dim rsQuelle As DAO.Recordset
    Dim rsZiel As DAO.Recordset
    Dim prop As Variant
    Dim oTab As DAO.TableDef
    Dim oCol As DAO.Field
    Dim pr As DAO.Property
    Dim Tb As TableDef
    Dim sDB, _
        sTabFieldName, _
        sLangField, _
        sSQL _
    As String
    Dim qExportUeSQuery As DAO.QueryDef
    Dim sCurrentLanguage As String
    Dim sLanguageSpalte As String
    Dim s As String

    Set ldb = openAccDBConnection("openDB")
    Set Tb = ldb.TableDefs("Versionierung")

    sDB = Right(Tb.Connect, Len(Tb.Connect) - InStr(1, Tb.Connect, "="))
'    Debug.Print sDB
    Set Tb = Nothing
     ' Set ldb = openAccDBConnection("closeDB")

    Set ldb = DBEngine.Workspaces(0).OpenDatabase(sDB)


    sCurrentLanguage = getCurrLanguageDB()

    Set rsImportObj = CurrentDb.OpenRecordset("SELECT RSP_Lokalisierung.TypeOfObject, RSP_Lokalisierung.OjectID_Name " & _
                                     "From RSP_Lokalisierung " & _
                                     "GROUP BY RSP_Lokalisierung.TypeOfObject, RSP_Lokalisierung.OjectID_Name;")

    Select Case Me.KombLangDB '20160122_RSP
        Case "German"
            sLangField = "BasisText"
            CurrentDb.Execute ("Update KonfigurationReklDB set [Value] = 'German' where [Konfiguration] = 'Location';")
        Case "English"
            sLangField = "TextE"
            CurrentDb.Execute ("Update KonfigurationReklDB set [Value] = 'English' where [Konfiguration] = 'Location';")
        Case "Chinese"
            sLangField = "TextCN"
            CurrentDb.Execute ("Update KonfigurationReklDB set [Value] = 'Chinese' where [Konfiguration] = 'Location';")
    End Select


    Set rsImportObjFilter = Nothing
    rsImportObj.Filter = "[TypeOfObject] = 'TableItem'"
    Set rsImportObjFilter = rsImportObj.OpenRecordset
    rsImportObjFilter.MoveFirst
    While Not rsImportObjFilter.EOF
        Select Case rsImportObjFilter("OjectID_Name")
            Case "Fehlergruppen"
                sTabFieldName = "Fehlergruppe"
                'db.Execute ("Update " & rs("TabellenName") & " set [Anzeige] = [Fehlergruppe];")
            Case "Kategorien-Bearbeitung I"
                sTabFieldName = "Kategorie-Bearbeitung I"
            Case "Kategorien-Bearbeitung II"
                sTabFieldName = "Kategorie-Bearbeitung II"
            Case "Kategorien-Fehleranalyse"
                sTabFieldName = "Kategorie-Fehleranalyse"
            Case "Kategorien-Problemstatus"
                sTabFieldName = "Kategorie-Problemstatus"
            Case "Kategorien-Vorgangstatus"
                sTabFieldName = "Kategorie-Vorgangstatus"
            Case "Vorgang-Text"
                sTabFieldName = "Vorgang-Beschreibung"
            Case "ZuordnungDefTaetigkeit"
                sTabFieldName = "DefTaetigkeit"
            Case "Taetigkeiten"
                sTabFieldName = "Taetigkeit"
            Case "Mahnstufen"
                sTabFieldName = "Mahnstufe"
            Case "Meldungstyp"
                sTabFieldName = "Medlungstyp"
            Case "MaßnahmeUmsetzungFrist"
                sTabFieldName = "MaßnahmeFrist"
            Case "TaetigkeitenEinkauf"
                sTabFieldName = "Taetigkeit"
            Case "tblMenüleiste"
                sTabFieldName = "Unterregister"
            Case "tblRegister"
                sTabFieldName = "Name"
            Case "tblZuordnung"
                sTabFieldName = "Zuordnung"
            Case "tblTicketStatus"
                sTabFieldName = "TicketStatus"
            Case "MaßnahmenHerkunft"
                sTabFieldName = "MaßnahmenHerkunft"
            Case "tblTicketMahnungStatus"
                sTabFieldName = "TicketStatus"
            Case "tblTicketTyp"
                sTabFieldName = "TicketTyp"
            Case "tblSoftKeyMakros"
                sTabFieldName = "DefaultMakroBezeichnung"
            Case Else
                sTabFieldName = rsImportObjFilter("OjectID_Name")
        End Select
'        sSQL = "UPDATE [" & rsImportObjFilter("OjectID_Name") & "], RSP_Lokalisierung SET [" & rsImportObjFilter("OjectID_Name") & "].Anzeige = [RSP_Lokalisierung].[" & sLangField & "] " & _
'              "WHERE ([" & rsImportObjFilter("OjectID_Name") & "].[" & sTabFieldName & "] = [RSP_Lokalisierung].[BasisText]);"
'20190625-ROS
        sSQL = "UPDATE [" & rsImportObjFilter("OjectID_Name") & "], RSP_Lokalisierung SET [" & rsImportObjFilter("OjectID_Name") & "].Anzeige = [RSP_Lokalisierung].[" & sLangField & "] " & _
              "WHERE ([" & rsImportObjFilter("OjectID_Name") & "].[" & sTabFieldName & "] = [RSP_Lokalisierung].[BasisText]) and ([RSP_Lokalisierung].[OjectID_Name] = '" & rsImportObjFilter("OjectID_Name") & "') ;"

        ldb.Execute (sSQL)

        rsImportObjFilter.MoveNext
    Wend


    '########################################
    'Generierung derTabelle TicketManagement (TicketTyp-Tab)
    'Es wird in der Lokalisierungstabelle je nach TextTyp MailBetreff und MailText nach dem eingestellten Text gesucht und in Abhängigekit der neu eingestellten Sprache der zugehörige Text
    ' in die Zieltabelle eingetragen.
    '########################################
    'Textspalte nach ursprünglich eingestellter Sprache ermitteln
    Select Case sCurrentLanguage
        Case "German"
            sLanguageSpalte = "BasisText"
        Case "English"
            sLanguageSpalte = "TextE"
        Case "Chinese"
            sLanguageSpalte = "TextCN"
    End Select

    Set rsQuelle = ldb.OpenRecordset("select * from [tblTicketTyp];")
    If rsQuelle.RecordCount > 0 Then
        rsQuelle.MoveFirst
        While Not rsQuelle.EOF
            rsQuelle.Edit
            Set rsZiel = CurrentDb.OpenRecordset("Select * from RSP_Lokalisierung where TypeOfObject = 'TableItem2' and OjectID_Name = 'tblTicketTyp' and " & _
                                                 "TextID = 'MailBetreff' and [" & sLanguageSpalte & "] = '" & rsQuelle("MailBetreff").Value & "';")
            '1. Schritt Betreff Text austauschen
            If rsZiel.RecordCount Then rsQuelle("MailBetreff") = rsZiel(sLangField)
            Set rsZiel = CurrentDb.OpenRecordset("Select * from RSP_Lokalisierung where TypeOfObject = 'TableItem2' and OjectID_Name = 'tblTicketTyp' and " & _
                                                 "TextID = 'MailText' and [" & sLanguageSpalte & "] = '" & rsQuelle("MailText").Value & "';")
            '2. Schritt Mail Text austauschen
            If rsZiel.RecordCount Then rsQuelle("MailText") = rsZiel(sLangField)

            rsQuelle.Update
            rsQuelle.MoveNext
        Wend
    End If
    Set rsQuelle = Nothing





    'Set rsImportObj = Nothing

     ' Set ldb = openAccDBConnection("closeDB")
     Exit Sub
err_ButSpracheUmstellenDB_Click:

    If Err = 2465 Then
        rsQuelle.Delete
        Resume Next
    ElseIf Err = 3078 Then
        Resume Next
    End If
End Sub

Private Sub cboAutoAus_AfterUpdate()

Dim ldb As DAO.Database
Dim rs As DAO.Recordset

Set ldb = openAccDBConnection("openDB")

Set rs = ldb.OpenRecordset("Select * from KonfigurationReklDB where Konfiguration = 'ClientShutdown';")
    rs.Edit
    rs("Value") = Me.cboAutoAus
    rs.Update

    If Me.cboAutoAus = "Ja" Then
         ldb.Execute ("Update KonfigurationReklDB set [Value] = '" & Environ("COMPUTERNAME") & "' where Konfiguration = 'ClientShutdownOwner';")
    Else
         ldb.Execute ("Update KonfigurationReklDB set [Value] = '' where Konfiguration = 'ClientShutdownOwner';")
    End If
    Set rs = Nothing

End Sub

Private Sub cmbAccessFensterSetzten_Click()

    'Alle Accesss Fenster auf gleichen Bildschirm öffnen, Koordinaten werden an Formular TB:autostart angepasst
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim strWert As String

    strWert = DLookup("[Koordinaten]", "TAB_Koordinaten", "[FormName] = 'TB:autostart'")
    strWert = Left(strWert, InStr(strWert, "/") - 1)

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from [TAB_Koordinaten];")
    rs.MoveLast
    rs.MoveFirst
    Do While Not rs.EOF
        rs.Edit
        rs("Koordinaten") = strWert & "/" & MID(rs("Koordinaten"), InStr(rs("Koordinaten"), "/") + 1, InStrRev(rs("Koordinaten"), "/"))
        rs.Update
        rs.MoveNext
    Loop

    Set rs = Nothing

End Sub

Private Sub cmbAendern_Click()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    ldb.Execute ("update [Vorgang-Text] set [Anzeige] = '" & Me.txtVorgang & "', " & _
                                         "[Vorgang-Beschreibung] = '" & Me.txtBezeichnung & "', " & _
                                         "[Aktiv] = '" & Me.ctlAktiv & "', " & _
                                         "[Datenbank] = '" & Me.cbxDatabase & "' " & _
                                         "Where [Vorgang-ID] = " & Me.txtID)
    Me.lstVorgaenge.Requery
    Me.lstVorgaenge = Me.txtID
'    Me.txtID = rs("ID")
'    Me.txtVorgang = rs("Anzeige")
'    Me.txtBezeichnung = rs("Beschreibung")
'    Me.ctlAktiv = rs("Aktiv")
'    Me.cbxDatabase = rs("Datenbank")
'
'    Set rs = Nothing



End Sub

Private Sub cmbAendern_pKuerzel_Click()

    Aendern_pKuerzel

End Sub

Private Sub cmbAufgabenInTicket_Click()

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim fs As Object
Dim fsfold As Object
Dim fsFile As Object
Dim sPfad As String
Dim ll_TicketID As Long
Dim la_Anhaenge() As String


    DoCmd.Hourglass True
    Set ldb = openAccDBConnection("openDB")

    Set rs = ldb.OpenRecordset(" Select * from qryLQAufgabenInTicketUmwandeln")

    If rs.RecordCount Then
        rs.MoveFirst
        Do While Not rs.EOF
            ll_TicketID = TicketMWAnlegen(rs("ID"), 1, rs("QSWeiterleitung"), rs("QSBearbeitungBis_Soll"), rs("QSBearbeitungBis_Verl"), 2, rs("QSWeiterleitung"), rs("MailVon"), rs("MailTo"), rs("MailCC"), rs("MailBCC"))
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set rs = ldb.OpenRecordset(" Select * from qryLiefAufgabenInTicketUmwandeln")
    If rs.RecordCount Then
        rs.MoveFirst
        Do While Not rs.EOF
            ll_TicketID = TicketMWAnlegen(rs("ID"), 2, rs("Rückmeldung_Lieferant_Soll"), rs("Rückmeldung_Lieferant_Soll"), rs("Rückmeldung_Lieferant_Verlängertl"), 2, rs("Rückmeldung_Lieferant_Soll"), rs("MailVon"), rs("MailTo"), rs("MailCC"), "Operativer-Einkauf@de.lq-group.com", rs("TelefonBeruflich"))
            If IsNull(rs("Meldungsnummer Kunde")) Then
                sPfad = getReklVerzeichnis(rs("ID"), "2017")
            Else
                sPfad = getReklVerzeichnis(rs("ID"), "2017", rs("Meldungsnummer Kunde"))
            End If
            If sPfad <> "" Then
                If fs.FileExists(sPfad & rs("ID") & "_Reklamationsbericht_Lieferant.pdf") = True Then
                    ldb.Execute ("Insert into tblTicketAnhaenge ([TicketID],[ReklID],[AnhangPfad]) values (" & ll_TicketID & "," & rs("ID") & ",'" & sPfad & rs("ID") & "_Reklamationsbericht_Lieferant.pdf" & "');")
                End If
                If fs.FolderExists(sPfad & "\Attachment") = True Then
                    Set fsfold = fs.GetFolder(sPfad & "\Attachment")
                    For Each fsFile In fsfold.Files

                        la_Anhaenge = Split(fsFile, "\")
                        Select Case la_Anhaenge(UBound(la_Anhaenge))
                            Case "An.txt"
                                ldb.Execute ("Update tblTicketListe set [MailAn] = '" & readTextFile(fsFile) & "' where [TicketID] = " & ll_TicketID & ";")
                            Case "CC.txt"
                                ldb.Execute ("Update tblTicketListe set [MailCC] = '" & readTextFile(fsFile) & "' where [TicketID] = " & ll_TicketID & ";")
                            Case "Body.txt"
                                ldb.Execute ("Update tblTicketListe set [MailText] = '" & readTextFile(fsFile) & "' where [TicketID] = " & ll_TicketID & ";")
                            Case Else
                                ldb.Execute ("Insert into tblTicketAnhaenge ([TicketID],[ReklID],[AnhangPfad]) values (" & ll_TicketID & "," & rs("ID") & ",'" & fsFile & "');")
                        End Select
                    Next
                End If
            End If
            rs.MoveNext
        Loop
        Set rs = Nothing

    End If

    DoCmd.Hourglass False

    Exit Sub

End Sub

Private Sub cmbClientVersionUebernehmen_Click()     '20181203-ROS

Dim ldb As DAO.Database
Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select [Value] from KonfigurationReklDB where [Konfiguration] = 'ClientVersion';")
        rs.Edit
        rs("Value") = get_ClientVersion()
        rs.Update
    Set rs = Nothing

    Me.txtAktuelleVersion = get_AkuelleClientVersion()

    Me.Refresh
    Me.Requery

End Sub

Private Sub cmbDatumFormatAnpassen_Click()
    Dim ldb As DAO.Database
    Dim sCaption As String
    Dim frm As Object
    Dim Rpt As Object
    Dim ctl As Control
    Dim Tb As TableDef
    Dim sDB, _
        sTabFieldName _
    As String
    Dim dbSource _
    As DAO.Database
    Dim rs, _
        rsQuelle, _
        rsZiel _
    As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set Tb = ldb.TableDefs("Versionierung")

    sDB = Right(Tb.Connect, Len(Tb.Connect) - InStr(1, Tb.Connect, "="))
    'Debug.Print sDB
    Set Tb = Nothing
    Set ldb = openAccDBConnection("closeDB")

    Set ldb = DBEngine.Workspaces(0).OpenDatabase(sDB)
    'Rücksetzen des Update Datum in der Lokalisierung-Tabelle
    'db.Execute ("Update RSP_Lokalisierung set DateLastUpdate = Null;")
    'Geht über alle Formulare der aktuellen Client Applikation
    For Each frm In CurrentProject.AllForms
'        Debug.Print frm.Name
        'Alle Formulare werden behandelt, außer das für den Administrator, das bleibt in Deutsch
        If Not frm.Name = "KonfiguartionLQReklDB" Then
            'Das Formular muss im Hintergrund geööfnet werden um
            DoCmd.OpenForm frm.Name, acDesign, , , , acHidden
            For Each ctl In Forms(frm.Name).Controls
                If ctl.ControlType = 109 Then
                    If ctl.Properties("Format") = "Short Date" Then
                        If getCurrLanguage() = "English" Then
                            ctl.Properties("InputMask") = "00\/00\/00;0;*"
                        Else
                            ctl.Properties("InputMask") = "00,00,00;0;*"
                        End If
                    End If
                End If
            Next ctl
            DoCmd.Close acForm, frm.Name, acSave
        End If
    Next frm
End Sub

Private Sub cmbDatumInMailtextAendern_Click()

    Call fct_changedateinTicketlist

End Sub

Private Sub cmdDel_Click(fIndex As Integer)


    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("SELECT * from tblSoftKeyS where SchbearbeiterKuerzel = '" & get_pKuerzel() & "' and SoftkeyNummer = " & fIndex & ";")
    If rs.RecordCount Then
        rs.Delete
    End If
    Me.Controls("txtMakroSoftKey" & fIndex) = ""
    Me.Controls("txtMakroSoftKeyBez" & fIndex) = ""

    If IsLoaded("TB:autostart") Then
        If Forms("TB:autostart").Controls("cmdDynFunc" & fIndex).Visible = True Then
            Forms("TB:autostart").Controls("cmdDynFunc" & fIndex).Visible = False
        End If
    End If

    Set rs = Nothing


End Sub

Private Sub cmbNeu_Click()

    Me.txtID = ""
    Me.txtVorgang = ""
    Me.txtBezeichnung = ""
    Me.ctlAktiv = False
    Me.cbxDatabase = "ReklDB"


    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from [Vorgang-Text] where Anzeige = 'Neuer Vorgang';", dbOpenSnapshot)

    If rs.RecordCount = 0 Then
        ldb.Execute ("Insert into [Vorgang-Text] ([Anzeige],[Vorgang-Beschreibung],[Aktiv],[Datenbank]) values ('Neuer Vorgang','Neuer Vorgang',False,'ReklDB')")
        Set rs = ldb.OpenRecordset("Select * from [Vorgang-Text] where Anzeige = 'Neuer Vorgang';", dbOpenSnapshot)
    End If
    Me.txtID = rs("Vorgang-ID")
    Me.txtVorgang = rs("Anzeige")
    Me.txtBezeichnung = rs("Vorgang-Beschreibung")
    Me.ctlAktiv = rs("Aktiv")
    Me.cbxDatabase = rs("Datenbank")

    Set rs = Nothing

End Sub

Private Sub cmbOffeneAB_Click()

    DoCmd.OpenForm ("frmAB_Temp")

End Sub

Private Sub cmbProgrammCode_Click()
'20260803-ROS
    ProgrammcodeDurchsuchen InputBox("Programmcode", "Bitte zu suchenden Programmcode eingeben.")

End Sub

Private Sub cmbStundensatzCN_Click()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim Datei As String

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from RSP_Lokalisierung where OjectID_Name = '" & KonfigLQReklDBDatei & "';")
        rs.Edit
        rs("TextCN") = Me.[txtStundensatzCN]
        rs.Update
    Set rs = Nothing

End Sub

Private Sub cmbStundensatzDE_Click()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim Datei As String

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from RSP_Lokalisierung where OjectID_Name = '" & KonfigLQReklDBDatei & "';")
        rs.Edit
        rs("BasisText") = Me.[txtStundensatzDE]
        rs.Update
    Set rs = Nothing

End Sub

Private Sub cmbStundensatzE_Click()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim Datei As String

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from RSP_Lokalisierung where OjectID_Name = '" & KonfigLQReklDBDatei & "';")
        rs.Edit
        rs("TextE") = Me.[txtStundensatzE]
        rs.Update
    Set rs = Nothing

End Sub

Private Sub cmdDel1_Click()
    cmdDel_Click 1
End Sub
Private Sub cmdDel2_Click()
    cmdDel_Click 2
End Sub
Private Sub cmdDel3_Click()
    cmdDel_Click 3
End Sub
Private Sub cmdDel4_Click()
    cmdDel_Click 4
End Sub
Private Sub cmdDel5_Click()
    cmdDel_Click 5
End Sub

Private Sub cmdNeu_Click()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("SELECT * from tblSoftKeyMakros where MakroName = '" & Me.txtMakro & "';")
    If rs.RecordCount = 0 Then
        rs.AddNew
        rs("MakroName") = Me.txtMakro
        rs("DefaultMakroBezeichnung") = Me.txtDefaultMakroBez
        rs("Anzeige") = Me.txtDefaultMakroBez
        rs.Update
        Me.lstSofkeyMakro.Requery
    Else
'        MsgBox ("Makro schon vorhanden!")
        MsgBox getTC_FCFromCode("TC3237"), vbInformation, getTC_FCFromCode("TC3200")
    End If
    Set rs = Nothing
    Set db = Nothing


End Sub

Private Sub cmdSet_Click(fIndex As Integer)

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("SELECT * from tblSoftKeyS where SchbearbeiterKuerzel = '" & get_pKuerzel() & "' and SoftkeyNummer = " & fIndex & ";")
    If rs.RecordCount Then
        rs.Edit
    Else
        rs.AddNew
        rs("SoftkeyNummer") = fIndex
        rs("SchbearbeiterKuerzel") = get_pKuerzel()

    End If
    rs("SoftkeyMakro") = Me.lstSoftKeyMakroZuordnung.Column(1)
    Me.Controls("txtMakroSoftKey" & fIndex) = Me.lstSoftKeyMakroZuordnung.Column(1)
    rs("SoftKexBezeichnung") = Me.lstSoftKeyMakroZuordnung.Column(2)
    Me.Controls("txtMakroSoftKeyBez" & fIndex) = Me.lstSoftKeyMakroZuordnung.Column(2)
    rs.Update
    'me.

    If IsLoaded("TB:autostart") Then
        Forms("TB:autostart").Controls("cmdDynFunc" & fIndex).Caption = Me.Controls("txtMakroSoftKeyBez" & fIndex)
        If Forms("TB:autostart").Controls("cmdDynFunc" & fIndex).Visible = False Then
            Forms("TB:autostart").Controls("cmdDynFunc" & fIndex).Visible = True
        End If
    End If

    Set rs = Nothing
    Set db = Nothing


End Sub
Private Sub cmdSet1_Click()

    cmdSet_Click 1

End Sub
Private Sub cmdSet2_Click()

    cmdSet_Click 2

End Sub
Private Sub cmdSet3_Click()

    cmdSet_Click 3

End Sub
Private Sub cmdSet4_Click()

    cmdSet_Click 4

End Sub
Private Sub cmdSet5_Click()

    cmdSet_Click 5

End Sub

Private Sub cmdStatusNeuBerechnen_Click()

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim rs1 As DAO.Recordset
Dim rs2 As DAO.Recordset
Dim liHour As Integer
Dim liDays As Integer
Dim liMinute As Integer
Dim lNow As Date
Dim lDate2h As Date

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("SELECT ReklNr from tblStatusWechsel group by ReklNr;")
    rs.MoveFirst
    While Not rs.EOF
        'Basis der Zustände die neu berechnet werden müssen
        Set rs1 = ldb.OpenRecordset("SELECT * from tblStatusWechsel where ReklNr = " & rs("ReklNr") & " order by Datum;")
        'Zur Ermittlung des Folgezustandes zur Berecnung der Zeitdifferenz
        Set rs2 = ldb.OpenRecordset("SELECT * from tblStatusWechsel where ReklNr = " & rs("ReklNr") & " order by Datum;", dbOpenSnapshot)
        rs1.MoveFirst
        rs2.MoveFirst
            While Not rs1.EOF
                rs2.MoveNext
                rs1.Edit
                If Not rs2.EOF Then
'                    Debug.Print rs2("Status")
'                    Debug.Print rs1("Status")
                    rs1!AT = CalcArbeitszeitDiffDate(rs1!Datum, rs2("Datum")) 'Eintrag lokale Variable
                    rs1!AT = getHourMinuteByAnzMinute(rs1!AT, liDays, liHour, liMinute)
                    rs1!Zeit_AT = liDays
                    rs1!Zeit_h = liHour
                    rs1!Zeit_min = liMinute
                Else
                    rs1("Zeit_AT") = 0
                    rs1("Zeit_h") = 0
                    rs1("Zeit_min") = 0
                    rs1("AT") = 0
                    'Kein Folgezustand vorhanden - abbrechen oder .......
                End If
                rs1.Update
                rs1.MoveNext
            Wend
        rs.MoveNext
    Wend



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

    FormOpenKoor Me
    Call FormEinUndAusBlenden(Me)

End Sub

Private Sub Form_Open(Cancel As Integer)

Dim ldb As DAO.Database
Dim rs, rsl As DAO.Recordset

Set ldb = openAccDBConnection("openDB")
Set rs = ldb.OpenRecordset("Select * from KonfigurationReklDBLocal;", dbOpenSnapshot)

'Laden der einzelne  Werte aus Tabelle Lokal
rs.MoveFirst
While Not rs.EOF
    Select Case rs("Konfiguration")
        Case "LoggingClient"
            Me.InLoggingClient = rs("Value")
    End Select
rs.MoveNext
Wend

Set rs = ldb.OpenRecordset("Select * from KonfigurationReklDB;", dbOpenSnapshot)
'Laden der einzelne  Werte aus Tabelle Global
rs.MoveFirst
While Not rs.EOF
    Select Case rs("Konfiguration")
        Case "LoggingClientAll"
            If rs("Value") = "Aktiv" Then
                Me.InLoggingClientsAll = True
            Else
                Me.InLoggingClientsAll = False
            End If
    End Select
rs.MoveNext
Wend

Set rs = Nothing
 ' Set ldb = openAccDBConnection("closeDB")

End Sub


Private Sub Import_Click()

    Dim lexcelprov As clsExcelProvider



If (Me.QuelleImport <> "") Then

    Select Case Me.ImportVorgang
        Case "Import Tätigkeiten-Fehlerbilder"



        Case "Import Zurodnung Fehlerbild-WS Erfasser"
        Case "Import Zurodnung Fehlerbild-WS Analyse"""

    End Select
End If


End Sub


Function dbFieldExists(TabDef As TableDef, _
  ByVal sFieldName As String) As Boolean

  Dim sName As String

  On Error Resume Next
  sName = TabDef.Fields(sFieldName).Name
  dbFieldExists = (Err = 0)
  On Error GoTo 0
End Function

Private Sub ButImpUebersetzungstabellen_Click()
'Update_2016_01_19_RSP
'Einlesen der Übersetzungstabelle optimiert


    Dim sPathDatei, sDatei As String
    Dim aSplitPath() As String
    Dim aSplit() As String
    Dim sTempTab As String

    'Öffnen der 1. ImportDatei für die Steuerelemente der Formulare und Reports
    'Mit dem Öffnen das Quellverzeichnis speichern um die übrigend Übersetzunghsdateien einlesen zu können
    'Öffnen: xx.xx.xxx_RSP_Export_BezUnicode.xls
    sPathDatei = DateiOeffnen()
    aSplitPath = Split(sPathDatei, "\")
    aSplit = Split(aSplitPath(UBound(aSplitPath)), ".")
    If Right(aSplit(0), 1) = "E" Then
        sTempTab = "E"
    Else
        sTempTab = "CN"
    End If
    'DoCmd.TransferDatabase acImport, , , acTable, sDatei, "RSP_Export_BezUnicode_TempImp", False

    DoCmd.TransferSpreadsheet transfertype:=acImport, _
            TableName:="RSP_Export_Lokalisierung_TempImp", _
            FileName:=sPathDatei, Hasfieldnames:=True, _
             SpreadsheetType:=5
    'Import Englische Texte
    CurrentDb.Execute ("UPDATE RSP_Lokalisierung INNER JOIN RSP_Export_Lokalisierung_TempImp ON RSP_Lokalisierung.[Text-ID] = RSP_Export_Lokalisierung_TempImp.[Text-ID] " & _
                       "SET RSP_Lokalisierung.Text" & sTempTab & " = RSP_Export_Lokalisierung_TempImp.Text" & sTempTab & ", RSP_Lokalisierung.Übersetzt" & sTempTab & " = True, " & _
                       "DateLastUpdate = #" & format(Date, "mm\/dd\/yyyy") & "# " & _
                       "WHERE (((RSP_Export_Lokalisierung_TempImp.Text" & sTempTab & ")<>''));")
'   Import Chineschische Texte
'    CurrentDb.Execute ("UPDATE RSP_Lokalisierung INNER JOIN RSP_Export_Lokalisierung_TempImp ON RSP_Lokalisierung.[Text-ID] = RSP_Export_Lokalisierung_TempImp.[Text-ID] " & _
'                       "SET RSP_Lokalisierung.TextCN = RSP_Export_Lokalisierung_TempImp.TextCN, RSP_Lokalisierung.ÜbersetztCN = True " & _
'                       "WHERE (((RSP_Export_Lokalisierung_TempImp.TextCN)<>''));")
    DoCmd.DeleteObject acTable, "RSP_Export_Lokalisierung_TempImp"

            'The Spreadsheet type = 5 specifies an Excel 5.0/7.0 file
            'format
End Sub

Private Sub ButGenUebersetzungstabellen_Click()
On Error GoTo err_ButGenUebersetzungstabellen_Click
'Funktionsbeschreibung:
'Die Funktion generiert in erster Linie Übersetzungstabellen, um die Applikation in verschieden Sprachen betreiben zu können. Diese Funktion schafft die Umgebung alle sichbaren
'Beschreibungen (Überschriften, Feldbeschreibungen, Spaltenüberschriften, usw.) in verschieden Sprachen (Deutsch, Englsich, Chinesisch) darstellen zu können. Man muss hier
'die verschieden Arten von Beschreibungen unterscheiden
'   1. Beschreibungen in Formularen
'   2. Spaltenüberschriften in Tabellen (Alias)
'   3. Und Auswahlfelder mit Texten, die sich in Tabelleninhalten wiederfinden
'Alle 3 Beschreibungsarten haben unterschiedliche Bereich, in den der sichtbare Text verwaltet wird. Daher müssen auch alle 3 Bereiche unabhängig in der generierung der
'Lokaliserung bearbeitet werden.
'1. Lokalisierung der Formularbeschreibungen
'Die Beschreibungen der Steuerelemente in Formularen ist in der Eigenschaft des jeweiligen Steuerlements entahlten, ohne Möglichkeit eine weitere Beschreibung (Fremdsprache)
'unter zu bringen. Daher muss eine Tabelle generiert werden, die in Abhängig jedes Formular und deren Steuerelemente das Beschreibungsfeld plus weitere Felder für die Lokalisierung
'beschreibt.
'Anschließend müssen alle Formulare auf Steuerelemente mit der Eigenschaft "Caption" durchsucht und in die Lokalisierungstabelle für Formulare eingetragen werden.
'Ausnahme: Alle Befehlschaltflächen mit Bildern werden automatisch im Tabellenfeld "Basistext" mit dem Anfangstext "Befehl" gekennzeichnet. Beim Export der Übersetzungstabellen werden
'          diese Einträge gefiltert und auf Übersetzt gekennzeichnet, da diese Bilder anstelle Text beinhalten.
'Aufbau der Tabelle:
'   1. ID - Steuerelement
'   2. Formular - Heimat des Steuerelements
'   3. Caption - Inhalt Steuerelement
'   4. ControlTyp - Typ welches ein Caption beinhaltet
'   5. Geändert - Sitzt auf True, wenn das Steuerlement das erste mal eingetargen und wenn es sich geändert hat. Übersetzung ist notwendig
'   6. Länge Caption - Anzahl Zeichen der Caption. Ist notwendig damit die Übersetzung im Rahmen der Größe des Original Caption übersetzt wird
'   7. BezEnglisch - Englische Übersetzung
'   8. BezChinese - Chinische Übersetzung
'Die Tabelle kann im Anschluss über eine Query exportiert werden, die nur die zu ändernden Captions beinhaltet. Bei ersten Export spllten alle Captions beinhaltet sein. Jeder weitere
'Export beinhaltet dann nur noch neu hinzugekommen Steuerelement bzw. geänderte Captions von bestehenden Steuerlementen.
'
'2. Lokalisierung der Spaltenüberschriften von Tabellen
'Bei Tabellen verhält es sich ähnlich wie bei Formularen. Jede Spalte hat ein Beschriftungsfeld (Caption) genannt Alias, das nur eine Sprache enthalten kann.
'Das Alias ersetzt den Feldname einer Tabelle, wenn diese in der Tabellenansicht angezeigt wird. Die Übersetzungen der Captions muss wie bei Formularen in einer separaten Tabelle
'für alle Tabellen und Feldnamen generiert werden. Die Lokalisierungstabelle für Tabelle ist identsich aufgebaut wir die Formualare.

    Dim ldb As DAO.Database
    Dim sCaption As String
    Dim frm As Object
    Dim Rpt As Object
    Dim ctl As Control
    Dim ctll As Control
    Dim dbSource As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsQuelle As DAO.Recordset
    Dim rsZiel As DAO.Recordset
    Dim prop As Variant
    Dim oTab As DAO.TableDef
    Dim oCol As DAO.Field
    Dim pr As DAO.Property
    Dim Tb As TableDef
    Dim sDB, _
        sTabFieldName _
    As String
    Dim qExportUeSQuery As DAO.QueryDef

    Dim cbr As CommandBar
    Dim cbc, cbc2 As CommandBarControl


    Set ldb = openAccDBConnection("openDB")
    'Ref der Tabelle Versionierung um heraus zu bekommen wo diese Tabelle liegt
    Set Tb = ldb.TableDefs("Versionierung")

    If getCurrLanguage = "English" Then
'        MsgBox ("Bitte Sprache erst auf Deutsch umstellen!")
        MsgBox getTC_FCFromCode("TC3238"), vbInformation, getTC_FCFromCode("TC3239")
        Exit Sub
    End If



    sDB = Right(Tb.Connect, Len(Tb.Connect) - InStr(1, Tb.Connect, "="))
'    Debug.Print sDB
    Set Tb = Nothing
     ' Set ldb = openAccDBConnection("closeDB")

    Set ldb = DBEngine.Workspaces(0).OpenDatabase(sDB)
    'Rücksetzen des Update Datum in der Lokalisierung-Tabelle
    ldb.Execute ("Update RSP_Lokalisierung set DateLastUpdate = Null;")
    'Geht über alle Formulare der aktuellen Client Applikation
    For Each frm In CurrentProject.AllForms
'        Debug.Print frm.Name
        'Alle Formulare werden behandelt, außer das für den Administrator, das bleibt in Deutsch
        If Not frm.Name = "KonfiguartionLQReklDB" Then
            'Das Formular muss im Hintergrund geööfnet werden um
            DoCmd.OpenForm frm.Name, acDesign, , , , acHidden
            For Each ctl In Forms(frm.Name).Controls
                sCaption = "NoCaption"
                For Each prop In ctl.Properties
                    If prop.Name = "Caption" Then
                        sCaption = ctl.Caption
                        Exit For
                    End If
                Next prop
                If sCaption <> "NoCaption" Then
                    Set rs = ldb.OpenRecordset("Select * from RSP_Lokalisierung where TypeOfObject = 'Form' and OjectID_Name = '" & frm.Name & "' and TextID = '" & ctl.Name & "';")
                    If rs.RecordCount = 0 Then
                        ldb.Execute ("Insert into RSP_Lokalisierung (TypeOfObject, OjectID_Name, TextID, BasisText, LängeText, DateLastUpdate) values ( 'Form','" & frm.Name & "', '" & ctl.Name & "', '" & sCaption & "', " & Len(sCaption) & ", " & getDateUSSQL(Date) & ");")
                    Else
                        rs.Edit
                        If rs("BasisText").Value <> sCaption Then
                            rs("BasisText").Value = sCaption
                            rs("LängeText").Value = Len(sCaption)
                            rs("TextE").Value = ""
                            rs("TextCN").Value = ""
                            rs("ÜbersetztE").Value = True
                            rs("ÜbersetztCN").Value = True
                        End If
                        rs("DateLastUpdate").Value = Date
                        rs.Update
                    End If
                End If
next_Ctl:
            Next ctl
            DoCmd.Close acForm, frm.Name, acSaveNo
            If frm.Name = "TB:autostart" Then
                DoCmd.OpenForm frm.Name, acNormal
            End If
        End If
next_Form:
    Next frm

     'Geht über alle Bericht der aktuellen Client Applikation
    For Each Rpt In CurrentProject.AllReports
'        Debug.Print Rpt.Name
        'Alle Berichte werden behandelt, außer das für den Administrator, das bleibt in Deutsch
        If Rpt.Name = "MaßnahmenBericht" Then
            'Der Bericht muss im Hintergrund geöfnet werden um
            DoCmd.OpenReport Rpt.Name, acDesign, , , , acHidden
            For Each ctll In Reports(Rpt.Name).Controls
                sCaption = "NoCaption"
                For Each prop In ctll.Properties
                    If prop.Name = "Caption" Then
                        sCaption = ctll.Caption
                        Exit For
                    End If
                Next prop
                If sCaption <> "NoCaption" Then
                'If ctl.ControlType = acLabel Then
                    Set rs = ldb.OpenRecordset("Select * from RSP_Lokalisierung where TypeOfObject = 'Report' and OjectID_Name = '" & Rpt.Name & "' and TextID = '" & ctll.Name & "';")
                    If rs.RecordCount = 0 Then
                        ldb.Execute ("Insert into RSP_Lokalisierung (TypeOfObject, OjectID_Name, TextID, BasisText, LängeText, DateLastUpdate) values ( 'Report','" & Rpt.Name & "', '" & ctll.Name & "', '" & sCaption & "', " & Len(sCaption) & ", " & Date & ");")
                    Else
                        rs.Edit
                        If rs("BasisText").Value <> sCaption Then
                            rs("BasisText").Value = sCaption
                            rs("LängeText").Value = Len(sCaption)
                            rs("TextE").Value = ""
                            rs("TextCN").Value = ""
                            rs("ÜbersetztE").Value = True
                            rs("ÜbersetztCN").Value = True
                        End If
                        rs("DateLastUpdate").Value = Date
                        rs.Update
                    End If
                End If
next_Ctll:
            Next ctll
'            If Frm.Name = "TB:autostart" Then
'                DoCmd.OpenForm Frm.Name, acNormal
'            Else
            DoCmd.Close acReport, Rpt.Name, acSaveNo
'            End If
        End If
next_Report:
    Next Rpt


     For Each oTab In ldb.TableDefs
'        Debug.Print oTab.Name
        If Left(oTab.Name, 4) <> "MSys" Then
            'DoCmd.OpenForm Frm.Name, acDesign, , , , acHidden
            For Each oCol In oTab.Fields
'                Debug.Print oCol.Name
                sCaption = "NoCaption"
                For Each prop In oCol.Properties
                    If prop.Name = "Caption" Then
                        sCaption = prop.Value
                        Exit For
                    End If
                Next prop
                Set rs = ldb.OpenRecordset("Select * from RSP_Lokalisierung where TypeOfObject = 'TableAlias' and OjectID_Name = '" & oTab.Name & "' and TextID = '" & oCol.Name & "';")
                If rs.RecordCount = 0 Then
                    If sCaption = "NoCaption" Then
                        Set pr = oCol.CreateProperty("Caption", dbText)
                        pr.Value = oCol.Name
                        oCol.Properties.Append pr
                        'oCol.Properties("Caption") = oCol.Name
                        sCaption = oCol.Name
                    End If
                    ldb.Execute ("Insert into RSP_Lokalisierung (TypeOfObject, OjectID_Name, TextID, BasisText, LängeText, DateLastUpdate) values ( 'TableAlias','" & oTab.Name & "', '" & oCol.Name & "', '" & sCaption & "', " & Len(sCaption) & ", " & getDateUSSQL(Date) & ");")
                Else
                    rs.Edit
                    If rs("BasisText").Value <> sCaption Then
                        rs("BasisText").Value = sCaption
                        rs("LängeText").Value = Len(sCaption)
                        rs("TextE").Value = ""
                        rs("TextCN").Value = ""
                        rs("ÜbersetztE").Value = True
                        rs("ÜbersetztCN").Value = True
                    End If
                    rs("DateLastUpdate").Value = Date
                    rs.Update
                    'db.Execute ("update RSP_BezUnicode set [BezGerman] = '" & sCaption & "', [Länge Bezeichnung] = " & Len(ctl.Caption) & " where Formular = '" & Frm.Name & "' and  Bezeichnungsfeld = '" & ctl.Name "';")
                End If
next_oCol:
            Next oCol
'            If Frm.Name = "TB:autostart" Then
'                DoCmd.OpenForm Frm.Name, acNormal
'            Else
'                DoCmd.Close acForm, Frm.Name, acSaveNo
'            End If
        End If
next_Tab:
    Next oTab
    '########################################
    'Generierung aller Übersetzungstabellen
    '########################################
    Set rs = ldb.OpenRecordset("Select * from RSP_Tabellen_Lokalisierung;")
    rs.MoveFirst
    While Not rs.EOF
        If IsNull(DLookup("[Name]", "MSysObjects", "[Name] = '" & rs("TabellenName") & "'")) = False Then
            '########################################
            'Erweitern der bestehenden Tabellen um die Lokaliserungsfelder
            '########################################
''            Debug.Print rs("TabellenName")
            Select Case rs("TabellenName")
                Case "Fehlergruppen"
                    sTabFieldName = "Fehlergruppe"
                    'db.Execute ("Update " & rs("TabellenName") & " set [Anzeige] = [Fehlergruppe];")
                Case "Kategorien-Bearbeitung I"
                    sTabFieldName = "Kategorie-Bearbeitung I"
                Case "Kategorien-Bearbeitung II"
                    sTabFieldName = "Kategorie-Bearbeitung II"
'                Case "Kategorien-Fehleranalyse"
'                    sTabFieldName = "Kategorie-Fehleranalyse"
                Case "Kategorien-Problemstatus"
                    sTabFieldName = "Kategorie-Problemstatus"
                Case "Kategorien-Vorgangstatus"
                    sTabFieldName = "Kategorie-Vorgangstatus"
                Case "Vorgang-Text"
                    sTabFieldName = "Vorgang-Beschreibung"
                Case "ZuordnungDefTaetigkeit"
                    sTabFieldName = "DefTaetigkeit"
                Case "Taetigkeiten"
                    sTabFieldName = "Taetigkeit"
                Case "Mahnstufen"
                    sTabFieldName = "Mahnstufe"
                Case "Meldungstyp"
                    sTabFieldName = "Medlungstyp"
                Case "MaßnahmeUmsetzungFrist"
                    sTabFieldName = "MaßnahmeFrist"
                Case "TaetigkeitenEinkauf"
                    sTabFieldName = "Taetigkeit"
                Case "tblMenüleiste"
                    sTabFieldName = "Unterregister"
                Case "tblRegister"
                    sTabFieldName = "Name"
                Case "tblZuordnung"
                    sTabFieldName = "Zuordnung"
                Case "tblTicketStatus"
                    sTabFieldName = "TicketStatus"
                Case "MaßnahmenHerkunft"
                    sTabFieldName = "MaßnahmenHerkunft"
                Case "tblTicketMahnungStatus"
                    sTabFieldName = "TicketStatus"
                Case "tblTicketTyp"
                    sTabFieldName = "TicketTyp"
                Case "tblSoftKeyMakros"
                    sTabFieldName = "DefaultMakroBezeichnung"
                Case Else
                    sTabFieldName = rs("TabellenName")
            End Select

            If Not dbFieldExists(db.TableDefs(rs("TabellenName")), "Anzeige") Then
                Set Tb = ldb.TableDefs(rs("TabellenName"))
                With Tb
                    .Fields.Append .CreateField("Anzeige", dbText)
                End With
                Set Tb = Nothing
                'Kopiere die Bezugsspalte in die Anzeigespalte
                ldb.Execute ("Update [" & rs("TabellenName") & "] set [Anzeige] = [" & sTabFieldName & "];")
            End If

            'Fülle die Übersetungstabelle
            'Über alle Einträge der Bezugstabelle
'            If rs("TabellenName") = "tblTicketStatus" Then
'                Debug.Print rs("TabellenName")
'            End If
            Set rsQuelle = ldb.OpenRecordset("select * from [" & rs("TabellenName") & "];")
            If rsQuelle.RecordCount > 0 Then 'Update_2016_01_19_RSP Abfrage auf Recordcount > 0
                rsQuelle.MoveFirst
                While Not rsQuelle.EOF
                    Set rsZiel = ldb.OpenRecordset("Select * from RSP_Lokalisierung where TypeOfObject = 'TableItem' and OjectID_Name = '" & rs("TabellenName") & "' and TextID = '" & sTabFieldName & "' and BasisText = '" & rsQuelle(sTabFieldName).Value & "';")
                    If rsZiel.RecordCount = 0 Then
                        ldb.Execute ("Insert into RSP_Lokalisierung (TypeOfObject, OjectID_Name, TextID, BasisText, LängeText, DateLastUpdate) values ( 'TableItem','" & rs("TabellenName") & "', '" & sTabFieldName & "', '" & rsQuelle(sTabFieldName).Value & "', " & Len(rsQuelle(sTabFieldName).Value) & ", " & getDateUSSQL(Date) & ");")
                    Else
                        If rsZiel("BasisText").Value <> rsQuelle(sTabFieldName).Value Then
                            rsZiel.Edit
                            rsZiel("BasisText").Value = rsQuelle(sTabFieldName).Value
                            rsZiel("LängeBez").Value = Len(rsQuelle(sTabFieldName).Value)
                            rsZiel("TextE").Value = ""
                            rsZiel("TextCN").Value = ""
                            rsZiel("ÜbersetztE").Value = False
                            rsZiel("ÜbersetztCN").Value = False
                            rs("DateLastUpdate").Value = Date
                            rsZiel.Update
                        End If
                    End If
                    Set rsZiel = Nothing
                    rsQuelle.MoveNext
                Wend
            End If
            Set rsQuelle = Nothing

        End If
        Set Tb = Nothing

        rs.MoveNext
    Wend


    '########################################
    'Generierung derTabelle TicketManagement (TicketTyp-Tab)
    'Es wird in der Lokalisierungstabelle je nach TextTyp MailBetreff und MailText nach dem eingestellten Text gesucht und in Abhängigekit der neu eingestellten Sprache der zugehörige Text
    ' in die Zieltabelle eingetragen.
    '########################################
    Set rsQuelle = ldb.OpenRecordset("select * from [tblTicketTyp];")
    If rsQuelle.RecordCount > 0 Then
        rsQuelle.MoveFirst
        While Not rsQuelle.EOF
            Set rsZiel = ldb.OpenRecordset("Select * from RSP_Lokalisierung where TypeOfObject = 'TableItem2' and OjectID_Name = 'tblTicketTyp' and " & _
                                                 "TextID = 'MailBetreff' and [BasisText] = '" & rsQuelle("MailBetreff").Value & "';")
            '1. Schritt Betreff Text austauschen
            If rsZiel.RecordCount = 0 Then
                ldb.Execute ("Insert into RSP_Lokalisierung (TypeOfObject, OjectID_Name, TextID, BasisText, LängeText, DateLastUpdate) values ( 'TableItem2','tblTicketTyp', 'MailBetreff', '" & rsQuelle("MailBetreff").Value & "', " & Len(rsQuelle("MailBetreff").Value) & ", " & getDateUSSQL(Date) & ");")
            End If
            Set rsZiel = ldb.OpenRecordset("Select * from RSP_Lokalisierung where TypeOfObject = 'TableItem2' and OjectID_Name = 'tblTicketTyp' and " & _
                                                 "TextID = 'MailText' and [BasisText] = '" & rsQuelle("MailText").Value & "';")
            '2. Schritt Mail Text austauschen
            If rsZiel.RecordCount = 0 And Not IsNull(rsQuelle("MailText").Value) Then
                ldb.Execute ("Insert into RSP_Lokalisierung (TypeOfObject, OjectID_Name, TextID, BasisText, LängeText, DateLastUpdate) values ( 'TableItem2','tblTicketTyp', 'MailText', '" & rsQuelle("MailText").Value & "', " & Len(rsQuelle("MailText").Value) & ", " & getDateUSSQL(Date) & ");")
            End If
            rsQuelle.MoveNext
        Wend
    End If
    Set rsQuelle = Nothing





    'Ermitteln aller Menü-Texte
    Set cbr = getLocTextecmdBar()
'    Debug.Print cbr.Name, cbr.NameLocal, cbr.Type, _
        cbr.Position, cbr.BuiltIn
    For Each cbc In cbr.Controls
        If cbc.Type = msoControlPopup Then
            Set rsZiel = ldb.OpenRecordset("Select * from RSP_Lokalisierung where TypeOfObject = 'Menue' and OjectID_Name = 'CommandBarPopup' and TextID = '" & cbc.Caption & "' and BasisText = '" & cbc.Caption & "';")
            If rsZiel.RecordCount = 0 Then
                ldb.Execute ("Insert into RSP_Lokalisierung (TypeOfObject, OjectID_Name, TextID, BasisText, LängeText, DateLastUpdate) values ( 'Menue','CommandBarPopup', '" & cbc.Caption & "', '" & cbc.Caption & "', " & Len(cbc.Caption) & ", " & getDateUSSQL(Date) & ");")
            Else
                If rsZiel("BasisText").Value <> cbc.Caption Then
                    rsZiel.Edit
                    rsZiel("BasisText").Value = cbc.Caption
                    rsZiel("LängeBez").Value = Len(cbc.Caption)
                    rsZiel("TextE").Value = ""
                    rsZiel("TextCN").Value = ""
                    rsZiel("ÜbersetztE").Value = False
                    rsZiel("ÜbersetztCN").Value = False
                    rs("DateLastUpdate").Value = Date
                    rsZiel.Update
                End If
            End If
            Set rsZiel = Nothing
            For Each cbc2 In cbc.CommandBar.Controls
                If cbc2.Type = msoControlButton Then
                    Set rsZiel = ldb.OpenRecordset("Select * from RSP_Lokalisierung where TypeOfObject = 'Menue' and OjectID_Name = 'CommandBarButton' and TextID = '" & cbc2.Caption & "' and BasisText = '" & cbc2.Caption & "';")
                    If rsZiel.RecordCount = 0 Then
                        ldb.Execute ("Insert into RSP_Lokalisierung (TypeOfObject, OjectID_Name, TextID, BasisText, LängeText, DateLastUpdate) values ( 'Menue','CommandBarButton', '" & cbc2.Caption & "', '" & cbc2.Caption & "', " & Len(cbc2.Caption) & ", " & getDateUSSQL(Date) & ");")
                    Else
                        If rsZiel("BasisText").Value <> cbc2.Caption Then
                            rsZiel.Edit
                            rsZiel("BasisText").Value = cbc2.Caption
                            rsZiel("LängeBez").Value = Len(cbc2.Caption)
                            rsZiel("TextE").Value = ""
                            rsZiel("TextCN").Value = ""
                            rsZiel("ÜbersetztE").Value = False
                            rsZiel("ÜbersetztCN").Value = False
                            rs("DateLastUpdate").Value = Date
                            rsZiel.Update
                        End If
                    End If
                    Set rsZiel = Nothing
                End If
            Next cbc2
         End If
     Next cbc

    'Einbinden aller neu genrierten Tabellen
    Append_TableDefs2
    'Eportieren der Übersetzungstabellen
    'Alle Befehlsschaltflächen mit Bildern werden mit Übersetzt gekennzeichnet
    DoCmd.RunSQL "UPDATE RSP_Export_Lokalisierung_E SET RSP_Export_Lokalisierung_E.ÜbersetztE = True WHERE (((RSP_Export_Lokalisierung_E.BasisText) Like 'Befehl*'));"
    DoCmd.RunSQL "UPDATE RSP_Export_Lokalisierung_CN SET RSP_Export_Lokalisierung_CN.ÜbersetztCN = True WHERE (((RSP_Export_Lokalisierung_CN.BasisText) Like 'Befehl*'));"
    DoCmd.RunSQL "UPDATE RSP_Lokalisierung SET RSP_Lokalisierung.LängeText = len(RSP_Lokalisierung.BasisText);"


    If IsNull(DLookup("[Name]", "MSysObjects", "[Name] = 'RSP_Lokalisierung'")) = False Then
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, _
                 "RSP_Export_Lokalisierung_E", "RSP_Export_Lokalisierung_E.xls", True, _
                 "Übersetzung"
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, _
                 "RSP_Export_Lokalisierung_CN", "RSP_Export_Lokalisierung_CN.xls", True, _
                 "Übersetzung"
'        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, _
'                 "RSP_Export_Lokalisierung_E", "C:\" & Date & "_RSP_Export_Lokalisierung_E.xls", True, _
'                 "Übersetzung"
'        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, _
'                 "RSP_Export_Lokalisierung_CN", "C:\" & Date & "_RSP_Export_Lokalisierung_CN.xls", True, _
'                 "Übersetzung"
    End If

    Set rs = Nothing
    Set frm = Nothing
    Set Rpt = Nothing
    Set ctl = Nothing
     ' Set ldb = openAccDBConnection("closeDB")

Exit Sub
err_ButGenUebersetzungstabellen_Click:
    If Err = 7784 Then
        GoTo next_Form
    ElseIf Err = 438 Then
        Set ctl = Nothing
        GoTo next_Ctl
    End If

End Sub

Private Sub KonfigLQReklDBDatei_AfterUpdate()

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from RSP_Lokalisierung where TextID = '" & Me.Text73 & "';")
    If rs.RecordCount Then
        If Me.Text73 = "DP0008" Then
            Me.txtStundensatzDE = rs("BasisText")
            Me.txtStundensatzE = rs("TextE")
            Me.txtStundensatzCN = rs("TextCN")
        Else
            Me.Pfad_DE = rs("BasisText")
            Me.Pfad_E = rs("TextE")
            Me.Pfad_CN = rs("TextCN")
        End If

    End If
    Set rs = Nothing

End Sub

Private Sub lstVorgaenge_Click()


    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from [Vorgang-Text] where [Vorgang-ID] = " & Me.lstVorgaenge)

    Me.txtID = rs("Vorgang-ID")
    Me.txtVorgang = rs("Anzeige")
    Me.txtBezeichnung = rs("Vorgang-Beschreibung")
    Me.ctlAktiv = rs("Aktiv")
    Me.cbxDatabase = rs("Datenbank")

    Set rs = Nothing

End Sub

Private Sub OKAllgemein_Click()

ÜbernehmenAllgemein_Click
DoCmd.Close acForm, "KonfiguartionLQReklDB"

End Sub

Private Sub QuelleImport_Click()

    Dim sDatei As String
    Dim sTitel, sfilter As String

    Select Case Me.ImportVorgang
        Case "Import Tätigkeiten-Fehlerbilder"
            sTitel = "Bitte die Liste der fehlerbilder auswählen (*.xls)"
            sfilter = "*.xls*"
        Case "Import Zurodnung Fehlerbild-WS Erfasser"
        Case "Import Zurodnung Fehlerbild-WS Analyse"""

    End Select
    sDatei = DateiOeffnen(sTitel, sfilter)
    If (sDatei <> "") Then
        Me.QuelleInput = sDatei
    End If
End Sub

Private Sub RegisterStr0_Change()

Select Case Me.RegisterStr0
    Case 8 'SoftKey Konfig User
        'SoftKey Felder initialisieren
        Dim ldb As DAO.Database
        Dim rs As DAO.Recordset

        Set ldb = openAccDBConnection("openDB")
        Set rs = ldb.OpenRecordset("Select * from tblSoftkeys where SchbearbeiterKuerzel = '" & get_pKuerzel() & "' order by SoftkeyNummer;")

        While Not rs.EOF
            Me.Controls("txtMakroSoftKey" & rs("SoftkeyNummer")) = rs("SoftkeyMakro")
            Me.Controls("txtMakroSoftKeyBez" & rs("SoftkeyNummer")) = rs("SoftKexBezeichnung")
            rs.MoveNext
        Wend

        Set rs = Nothing
        Set db = Nothing


End Select
End Sub

Private Sub Schließen_Click()

    FormKoorSave Me
    DoCmd.Close

End Sub
Private Sub Befehl33_Click()
On Error GoTo Err_Befehl33_Click


    Screen.PreviousControl.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_Befehl33_Click:
    Exit Sub

Err_Befehl33_Click:
    MsgBox Err.Description  'FC0171
    Resume Exit_Befehl33_Click

End Sub



Private Sub txtMakroSoftKeyBez_AfterUpdate(fIndex As Integer)
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")

    If Me.Controls("txtMakroSoftKeyBez" & fIndex).Text = "" Then
        MsgBox "Leere Zeichenfolge ist nicht erlaubt!"
    Else
        Set ldb = openAccDBConnection("openDB")
        Set rs = ldb.OpenRecordset("Select * from tblSoftKeys where SchbearbeiterKuerzel = '" & get_pKuerzel() & "' and SoftkeyNummer = " & fIndex & ";")
        If rs.RecordCount Then
            rs.Edit
            rs("SoftKexBezeichnung") = Me.Controls("txtMakroSoftKeyBez" & fIndex).Text
            rs.Update
            If IsLoaded("TB:autostart") Then
                Forms("TB:autostart").Controls("cmdDynFunc" & fIndex).Caption = Me.Controls("txtMakroSoftKeyBez" & fIndex).Text
            End If
        End If
        Set rs = Nothing
    End If
End Sub

Private Sub txtMakroSoftKeyBez1_AfterUpdate()

    txtMakroSoftKeyBez_AfterUpdate 1

End Sub
Private Sub txtMakroSoftKeyBez2_AfterUpdate()

    txtMakroSoftKeyBez_AfterUpdate 2

End Sub
Private Sub txtMakroSoftKeyBez3_AfterUpdate()

    txtMakroSoftKeyBez_AfterUpdate 3

End Sub
Private Sub txtMakroSoftKeyBez4_AfterUpdate()

    txtMakroSoftKeyBez_AfterUpdate 4

End Sub
Private Sub txtMakroSoftKeyBez5_AfterUpdate()

    txtMakroSoftKeyBez_AfterUpdate 5

End Sub

Private Sub ÜbernehmenAllgemein_Click()

Dim ldb As DAO.Database
Dim rs As DAO.Recordset

Set ldb = openAccDBConnection("openDB")
If Me.InLoggingClientsAll Then
    ldb.Execute ("UPDATE KonfigurationReklDB SET KonfigurationReklDB.[Value] = 'Aktiv' " & _
                "WHERE (((KonfigurationReklDB.Konfiguration)='LoggingClientsAll'));")
    gLoggingClientsAll = "Aktiv"
Else
    ldb.Execute ("UPDATE KonfigurationReklDB SET KonfigurationReklDB.[Value] = 'Deaktiv' " & _
                "WHERE (((KonfigurationReklDB.Konfiguration)='LoggingClientsAll'));")
    gLoggingClientsAll = "Deaktiv"
End If
If Me.InLoggingClient = "Aktiv" Then
    ldb.Execute ("UPDATE KonfigurationReklDBLocal SET KonfigurationReklDBLocal.[Value] = 'Aktiv' " & _
                "WHERE (((KonfigurationReklDBLocal.Konfiguration)='LoggingClient'));")
    gLoggingClientsAll = "Aktiv"
Else
    ldb.Execute ("UPDATE KonfigurationReklDBLocal SET KonfigurationReklDBLocal.[Value] = 'Deaktiv' " & _
                "WHERE (((KonfigurationReklDBLocal.Konfiguration)='LoggingClient'));")
    gLoggingClientsAll = "Deaktiv"
End If

Set rs = Nothing
 ' Set ldb = openAccDBConnection("closeDB")

End Sub
Private Sub Befehl46_Click()
On Error GoTo Err_Befehl46_Click


    Screen.PreviousControl.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_Befehl46_Click:
    Exit Sub

Err_Befehl46_Click:
    MsgBox Err.Description    'FC0172
    Resume Exit_Befehl46_Click

End Sub

Private Sub Befehl72_Click()
    Me.Pfad_CN = Null
    Me.Pfad_DE = Null
    Me.Pfad_E = Null
    Me.txtStundensatzCN = Null
    Me.txtStundensatzDE = Null
    Me.txtStundensatzE = Null
End Sub
Private Sub CmbTBautostartöffnen_Click()
On Error GoTo Err_CmbTBautostartöffnen_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "TB:autostart"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_CmbTBautostartöffnen_Click:
    Exit Sub

Err_CmbTBautostartöffnen_Click:
    MsgBox Err.Description
    Resume Exit_CmbTBautostartöffnen_Click

End Sub
