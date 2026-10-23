Attribute VB_Name = "AccessControlMastertab"
Option Compare Database
Public db As DAO.Database

Sub lockDataset(fsReklnr As String)

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from [Mastertab-DSBerechtigung] where ID = " & fsReklnr & ";")
    If rs.RecordCount = 0 Then
    '    Me.Auftragsdatum = ""
        rs.AddNew
        rs("ID").Value = get_p_ReklNr()
    Else
        rs.Edit
    End If
    rs("Zustand").Value = "Gesperrt"
    rs("Kürzel").Value = get_pKuerzel()
    rs("Datum") = Date
    rs("Client") = Environ("COMPUTERNAME")
    rs.Update
    rs.Close

    Set rs = Nothing
     'Set ldb = openAccDBConnection("closeDB")

End Sub


Sub unlockDataset(fsReklnr As String)

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from [Mastertab-DSBerechtigung] where ID = " & fsReklnr & ";")
    If rs.RecordCount = 0 Then
    '    Me.Auftragsdatum = ""
        rs.AddNew
        rs("ID").Value = get_p_ReklNr()
    Else
        If (rs("Kürzel").Value = get_pKuerzel()) Then
            rs.Edit
        Else
            GoTo exit_unlockDataset
        End If
    End If
    rs("Zustand").Value = "Frei"
    rs("Kürzel").Value = get_pKuerzel()
    rs("Datum") = Date
    rs("Client") = Environ("COMPUTERNAME")
    rs.Update
    rs.Close
    Set rs = Nothing

exit_unlockDataset:
    Set rs = Nothing
     'Set ldb = openAccDBConnection("closeDB")

End Sub
Function GetStateVorgang(fsReklnr As String) As String
' RSP - 20131107 - Vorgang prüfen ob dieser abgeschlossen ist
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
     Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from [Mastertab] where ID = " & fsReklnr & ";")
    If rs.RecordCount > 0 Then
        If Not IsNull(rs("Kategorie-Status").Value) Then GetStateVorgang = rs("Kategorie-Status").Value
    Else
        GetStateVorgang = ""
    End If

    Set rs = Nothing
     'Set ldb = openAccDBConnection("closeDB")

End Function

Function proofAndLockDataset(fsReklnr As String) As String
' RSP - 20121029 - Fehlerkorrektur

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim bNewDS As Boolean   'RSP - 201020129

    bNewDS = False

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from [Mastertab-DSBerechtigung] where ID = " & fsReklnr & ";")
    If rs.RecordCount = 0 Then  'RSP - 201020129 - Fehlerkorrektur: ABfrage bei nicht vorhanden DS ergänzt
        rs.AddNew
        rs("ID").Value = get_p_ReklNr()
        rs("Zustand") = "Frei"
        rs("Datum") = Date
        rs("Client") = Environ("COMPUTERNAME")
        bNewDS = True
    End If

    If (rs("Zustand").Value = "Frei") Then
        If Not bNewDS Then
            rs.Edit
        End If
        rs("Zustand").Value = "Gesperrt"
        rs("Kürzel").Value = get_pKuerzel()
        rs("Datum") = Date
        rs("Client") = Environ("COMPUTERNAME")
        rs.Update
        rs.Close
        proofAndLockDataset = ""
    Else
        If rs("Kürzel").Value = get_pKuerzel() Then
            proofAndLockDataset = ""
        Else
            proofAndLockDataset = rs("Kürzel").Value
'        MsgBox "Reklamation ist durch " & rs("Kürzel").Value & " seit dem " & rs("Datum") & "in Bearbeitung." & Chr(10) & Chr(13) & _
               "Der Vorgang wird nur mit Leseberechtigung geöffnet." 'TC0200
        MsgBox getTC_FCFromCode("TC0200", ldb) & rs("Kürzel").Value & getTC_FCFromCode("TC0381", ldb) & rs("Datum") & getTC_FCFromCode("TC0382") & Chr(10) & Chr(13) & _
               getTC_FCFromCode("TC0380", ldb)
        End If
    End If
    Set rs = Nothing
     'Set ldb = openAccDBConnection("closeDB")

End Function

Public Function getLastDSMastertab() As Long

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim i As Long

    'Zahlenbereich für die Interne-Fehlererfassung über LQPP 20231023-ROS
    If getCurrLanguage = "German" Then
        i = cReklID_LQPP_DE
    ElseIf getCurrLanguage = "English" Then
        i = cReklID_LQPP_EN
    ElseIf getCurrLanguage = "Chinese" Then
        i = cReklID_LQPP_CN
    End If

    'Ermitteln letzte Reklamations-ID für Externe-Fehlererfassung unterhalb des Zahlenbereicht LQPP, ersetzt alten Programmcode "getLastDSMastertab"
    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select [ID] from Mastertab where [ID] < " & i & " ORDER BY [ID] ASC;")
    If rs.RecordCount > 0 Then
        rs.MoveLast
        getLastDSMastertab = rs("ID")
    Else
        getLastDSMastertab = 0
    End If

'    Set ldb = openAccDBConnection("openDB")
'    Set rs = ldb.OpenRecordset("Select Max(Mastertab.ID) as MaxID from Mastertab;")
'    If rs.RecordCount > 0 Then
'        getLastDSMastertab = rs("MaxID")
'    Else
'        getLastDSMastertab = 0
'    End If

    Set rs = Nothing

End Function

Public Function setNewReklDS(ByVal fReklDS As Long, ByVal fADate As String, ByVal fState As String, ByVal fKatProblem As String, ByVal fVorgang As String) As Long
On Error GoTo err_setNewReklDS

    Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")
    If fVorgang = "extern" Then
        ldb.Execute ("Insert into Mastertab ([ID],[Auftragsdatum],[Kategorie-Status],[Kategorie-Problem],[Extern],[Zeitbedarf]) values (" & fReklDS & ",'" & fADate & "','" & fState & "','" & fKatProblem & "', True,'0,25');")
    Else
        ldb.Execute ("Insert into Mastertab ([ID],[Auftragsdatum],[Kategorie-Status],[Kategorie-Problem],[Intern],[Zeitbedarf]) values (" & fReklDS & ",'" & fADate & "','" & fState & "','" & fKatProblem & "', ,True,'0,25');")
    End If
    InsertNewReklMemoDS (fReklDS)
    InsertNewEinkaufDS (fReklDS)

     'Set ldb = openAccDBConnection("closeDB")
    setNewReklDS = fReklDS

    Exit Function
err_setNewReklDS:
     'Set ldb = openAccDBConnection("closeDB")
    setNewReklDS = 0
    MsgBox Replace(getTC_FCFromCode("TC0434"), "<xxx>", fReklDS) '("Fehler beim Anlegen eines neuen Vorgangs - ReklNr: " & fReklDS & " konnte nicht angelegt werden (Eventuell schon vorhanden!?)")

End Function

Public Function openAccDBConnection(fs_Action As String, Optional f_option As String = "") As DAO.Database
On Error GoTo err_openAccDBConnection
    'Static ldb As DAO.Database
    'Dim rs As DAO.Recordset
    Dim strErgebnis As String

    If f_option = "New" Then
        Set db = CurrentDb
        Set openAccDBConnection = db
    Else
        If fs_Action = "openDB" Then
            If db Is Nothing Then
                 Set db = CurrentDb
                'MsgBox ("Initialisierung bzw. Erstellen der zentralen Datenbank-Instanz!")
            End If
            Set openAccDBConnection = db
        ElseIf fs_Action = "closeDB" Then
            db.Close
            Set db = Nothing
            Set openAccDBConnection = Nothing
        End If
    End If
    Exit Function

err_openAccDBConnection:
    'strErgebnis = MsgBox("Fehler: " & Err & " - " & Error, vbOKOnly + vbExclamation + vbDefaultButton1, "Fehler Öffnen Datanbank Verbindung")
    Resume Next
End Function

Public Sub closeAccDBConnection(ByRef fdb As DAO.Database)
'Public Function closeAccDBConnection(ByRef fdb As Variant)
On Error GoTo err_closeAccDBConnection

    Dim rs As DAO.Recordset
    Dim strErgebnis As String
    If Not fdb Is Nothing Then
        fdb.Execute ("Update KonfigurationReklDBLocal set [Value] = [Value] - 1  where Konfiguration = 'AktuelleDBVerbindungen';")
        Set rs = fdb.OpenRecordset("Select [Value] from KonfigurationReklDBLocal where Konfiguration = 'AktuelleDBVerbindungen';", ldbOpenSnapshot)
'        Forms("TB:autostart").Controls("txtAktuelleDBVerbindungen") = rs("Value")

        Set rs = Nothing
        fdb.Close
        Set fdb = Nothing
    End If
    'closeAccDBConnection = 1
    Exit Sub

err_closeAccDBConnection:
    strErgebnis = MsgBox(getTC_FCFromCode("TC0313") & Err & " - " & Error, vbOKOnly + vbExclamation + vbDefaultButton1, getTC_FCFromCode("TC0435"))
    Resume Next

End Sub
Public Function initDBEnviroment()

Dim ldb As DAO.Database

iAnzDBConnections = 0
iAnzDBRSConnections = 0

 Set ldb = openAccDBConnection("openDB")
db.Execute ("Update KonfigurationReklDBLocal set [Value] = 0 where Konfiguration = 'AktuelleDBVerbindungen';")
db.Execute ("Update KonfigurationReklDBLocal set [Value] = 0 where Konfiguration = 'AktuelleDBRSVerbindungen';")
'closeAccDBConnection ldb

End Function
Public Function getTC_FCFromCode(fCode As String, Optional fdb As DAO.Database = Nothing) As String

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsLoc As DAO.Recordset

    If fdb Is Nothing Then
        Set ldb = openAccDBConnection("openDB")
    Else
        Set ldb = fdb
    End If

    Set rs = ldb.OpenRecordset("Select * from RSP_Lokalisierung where TextID = '" & fCode & "';", dbOpenSnapshot)
    Set rsLoc = ldb.OpenRecordset("Select * from KonfigurationReklDBLocal where Konfiguration = 'Location';", dbOpenSnapshot)
    'init function
    getTC_FCFromCode = ""
    If rs.RecordCount > 0 Then
        Select Case rsLoc("Value")
            Case "German"
                getTC_FCFromCode = rs("BasisText")
            Case "English"
                If rs("TextE") <> "" Then getTC_FCFromCode = rs("TextE")
            Case "Chinese"
                If rs("TextCN") <> "" Then getTC_FCFromCode = rs("TextCN")
        End Select
    Else
        getTC_FCFromCode = ""
    End If


    Set rs = Nothing
    Set rsLoc = Nothing


End Function
Public Function getCurrLanguage() As String
    getCurrLanguage = DLookup("[Value]", "KonfigurationReklDBLocal", "[Konfiguration] = 'Location'")
End Function
Public Function getCurrLanguageDB() As String
    getCurrLanguageDB = DLookup("[Value]", "KonfigurationReklDB", "[Konfiguration] = 'Location'")
End Function

Public Function getTextFromLokalisierungsTab(fObjectTyp As String, fObject As String, fTextID As Integer, Optional fdb As DAO.Database = Nothing) As String
On Error GoTo err_getTextFromLokalisierungsTab

    '20260511_Change
    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsLoc As DAO.Recordset

    If fdb Is Nothing Then
        Set ldb = openAccDBConnection("openDB")
    Else
        Set ldb = fdb
    End If

    Set rs = ldb.OpenRecordset("Select * from RSP_Lokalisierung where TypeOfObject = '" & fObjectTyp & "' and OjectID_Name = '" & fObject & "' and [Text-ID] = " & fTextID & ";")
    Set rsLoc = ldb.OpenRecordset("Select * from KonfigurationReklDBLocal where Konfiguration = 'Location';")
    'init function
    getTextFromLokalisierungsTab = ""
    If rs.RecordCount > 0 Then
        Select Case rsLoc("Value")
            Case "German"
                getTextFromLokalisierungsTab = rs("BasisText")
            Case "English"
                If rs("TextE") <> "" Then getTextFromLokalisierungsTab = rs("TextE")
            Case "Chinese"
                If rs("TextCN") <> "" Then getTextFromLokalisierungsTab = rs("TextCN")
        End Select
    Else
        getTextFromLokalisierungsTab = ""
    End If


    Set rs = Nothing
    Set rsLoc = Nothing

    Exit Function
err_getTextFromLokalisierungsTab:
    Resume Next
End Function

Public Function getSQLStatInsertTable(fTable As String, fSQLBefehl As String, Optional fReklID As Long = 0) As String

    Dim sFields As String
    Dim sValues As String

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")

    sFields = ""
    sValues = ""
    Set rs = ldb.OpenRecordset("Select * from InsertTabTemplateWerte Where [TableZuordnung] = '" & fTable & "';")
    rs.MoveFirst
    While Not rs.EOF
        If (fSQLBefehl = "Insert") Then
            sFields = sFields & rs("FieldName") & ","
            If (rs("FieldType") = "String") Then
                sValues = sValues & "'" & rs("FieldValue") & "',"

            Else
                sValues = sValues & rs("FieldValue") & ","
            End If
        Else
            sFields = sFields & rs("FieldName") & "="
            If (rs("FieldTyp") = "String") Then
                sValues = sValues & "'" & rs("FieldValue") & "',"

            Else
                sValues = sValues & rs("FieldValue") & ","
            End If

        End If
        rs.MoveNext
    Wend


    If (fSQLBefehl = "Insert") Then
        sFields = Left(sFields, Len(sFields) - 1)
        sValues = Left(sValues, Len(sValues) - 1)
        getSQLStatInsertTable = "Insert into " & fTable & " (" & sFields & ") values (" & sValues & ");"
    Else
        sValues = Left(sValues, Len(sValues))
        getSQLStatInsertTable = "Update " & fTable & " set " & sValues & " where [ID]=" & fReklID & ";"
    End If

    Set rs = Nothing

End Function
