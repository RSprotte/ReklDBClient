Attribute VB_Name = "Ticket"
Option Compare Database
Option Explicit
'########################################## 20240822-ROS ###############################################
'#######################################################################################################
'Folgende Daten die für die Erstellung eines Ticket benötigt:
'TicketID               -           Autowert
'TicketTypID            Benötigt    Wird übergeben
'ReklID                 Benötigt    Wird Übergeben
'MailVon                Optional    Individuell
'MailAn                 Optional    Individuell
'MailCC                 Optional    Individuell
'MailBCC                Optional    DatenTicketTyp(TicketTypID, "EskalationsAdresse")
'MailBetreff            Optional    DatenTicketTyp(TicketTypID, "MailBetreff")
'MailText               Optional    DatenTicketTyp(TicketTypID, "MailText")
'MailTextHTML           Optional    HtmlTextCreate(DatenTicketTyp(TicketTypID, "MailText"), z.B strTextHtml)
'MailErstellt           Optional    Datum Heute
'MailDatum              Benötigt    Datum Mail versendet bzw. beim Start der Laufzeit-Tickets (LZ-.....)
'BearbeitungBis         Optional    Datum Heute + DatenTicketTyp(TicketTypID, "Arbeitstage")
'VerlängertBis          Optional    bei Mail versenden = BearbeitungBis, kann danach geändert werden
'ErledigtAm             -           wird nach Beendigung befüllt
'TicketStatusID         Benötigt    wird bei erstellen befüllt und nach Mail versendet oder Beendigung geändert
'LinkZuVorgangsOrdner   Optional    Individuell
'TicketMahnwesenID      -           wird von Mahnwesen befüllt
'ErledigtVon            -           wird nach Beendigung mit Kürzel des Bearbeiters befüllt/ bei Laufzeit-Tickets (LZ-.....) mit "System"
'#######################################################################################################
'Erklärung Felder TickertTyp:
'TicketTypID            Autowert
'TicketTyp              Name
'Anzeige                Name Lokalisiert
'Arbeitstage            Bearbeitungszeit des Tickets'MailBetreff
'MailText
'AnsichtSteuerung       Ansichtsteuereung für Ticketauswahl von Service und QS beim Erstellen eines Tickets / Boolean
'ErinnerungTage         Erinnerung senden nach x Tagen
'Mahnung1Tage           1 Mahnung senden nach x Tagen
'Mahnung2Tage           2 Mahnung senden nach x Tagen
'EskalationsAdresse     Wird bei 2 Mahnung angezogen
'MahnwesenAktiv         Ticket so durch das Mahnwesen überwacht werden ja oder nein / Boolean
'MehrfachTicket         TicketTyp kann mehrfach erstellt werden ja oder nein / Boolean
'TicketAnzeigen         Steuert die Anzeige in Ticketauswahl bei der Ticket Erstellung Anzeige ja oder nein / Boolean
'#######################################################################################################

Public Function TicketErstellen(ByVal fTicketTypID As Integer, ByVal fReklaID As Long) As Long
'Ticket wir erstellt mit Übergabe der VorgangsID und des TicketTyps
'Rückgabe der erstellten TicketID zur weiteren Verwendung

    Dim ldb As DAO.Database
    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    Set rs = ldb.OpenRecordset("Select * from tblTicketListe;")

    ldb.Execute ("Insert into tblTicketListe (TicketTypID, ReklID) values (" & fTicketTypID & ", " & fReklaID & ");")
    rs.MoveFirst
    rs.MoveLast
    TicketErstellen = rs("TicketID")

    Set rs = Nothing

End Function

Public Sub TicketLoeschen(ByVal fTickertID As Long)
'Löscht eine bereits erstelltes Ticket, das nicht mehr benötigt wird
'Es können nur Tickets gelöscht werden, die den Status "Erstellt" haben
'andere Tickes mussen als erledigt markiert werden

On Error GoTo Err

    Dim ldb As DAO.Database

    Set ldb = openAccDBConnection("openDB")

    'Tickets können nur gelöscht werden, wenn Ticketstatus = Erstellt
    If DLookup("[TicketStatusID]", "tblTicketListe", "[TicketID] = " & fTickertID) = 1 Then   ' 1 = Erstellt
    'Sicherheitsabfrage ob Ticket wirklich gelöscht werden soll
        If MsgBox(getTC_FCFromCode("TC4213"), vbYesNo, getTC_FCFromCode("TC3210")) = vbYes Then
            DoCmd.SetWarnings False
            ldb.Execute ("Delete * from tblTicketListe where [TicketID] = " & fTickertID & ";")
            DoCmd.SetWarnings True
            Exit Sub
        End If
    Else
        'Ticket nicht löschbar, muss abgeschlossen werden
        MsgBox getTC_FCFromCode("TC4242"), vbExclamation, getTC_FCFromCode("TC3210")
        Exit Sub
    End If

Err:
    If Err.Number = 2427 Then
'        MsgBox ("Kein Datensatz ausgewählt")
        MsgBox getTC_FCFromCode("TC3240"), vbExclamation
        Exit Sub
    End If

End Sub

Public Sub TicketFuellen(ByVal fTickertID As Long, ByVal fWert As String, fFeld As String)
'Ticket mit den benötigten Daten füllen oder Updaten

    Dim ldb As DAO.Database
'    Dim rs As DAO.Recordset

    Set ldb = openAccDBConnection("openDB")
    ldb.Execute ("Update tblTicketListe set [" & fFeld & "]  = '" & fWert & "' where [TicketID]= " & fTickertID & " ;")

'    Set rs = ldb.OpenRecordset("Select * from [tblTicketListe] where [TicketID] = " & fTickertID & ";")
'    If rs.RecordCount Then
'        rs.Edit
'        rs(fFeld) = fWert
'        rs.Update
'    End If
'
'    Set rs = Nothing

End Sub

Public Function DatenTicketTyp(ByVal fTicketTypID As Integer, ByVal fFeldName As String) As String
'Abruf der im TicketTyp hinterlegten Daten. Übergabe TicketTypID und FeldName zum Abrufen der Daten.

    DatenTicketTyp = Nz(DLookup(fFeldName, "tblTicketTyp", "[TicketTypID] = " & fTicketTypID & ""))

End Function

Public Function getTicketID(ByVal fTicketTypID As Integer, ByVal fRekla As String) As Long
'Suchen offener Tickets zum Abschluss der Tickets anhand des TicketTyps und der Vorgangsnummer
'TicketStatus 2 = in Bearbeitung

    getTicketID = Nz(DLookup("TicketID", "tblTicketListe", "[TicketTypID] = " & fTicketTypID & " and [TicketStatusID] = 2 and [ReklID] = " & fRekla & ""))

End Function

Public Function TicketExist(ByVal fTicketTypID As Integer, ByVal fRekla As String) As Integer

    TicketExist = Nz(DLookup("TicketID", "tblTicketListe", "[TicketTypID] = " & fTicketTypID & " and [TicketStatusID] = 2 and [ReklID] = " & fRekla & ""))

End Function
