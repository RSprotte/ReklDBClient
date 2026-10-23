Attribute VB_Name = "Mahnwesen"
Option Compare Database

Public Function MahnwesenAusfuehren() As Integer

    Dim ldb As DAO.Database
    Dim rsMW_Tickets  As DAO.Recordset
    Dim rsLastMW_Ticket As DAO.Recordset
    Dim li_nextMahnstufe As Integer
    Dim li_currMahnstufe As Integer
    Dim rs_EMailLoad As Recordset
    Dim rs_MahnstufeLoad As Recordset
    Dim rs_Anhaenge As Recordset
    Dim sBetreff As String
    Dim sBody As String
    Dim scc As String
    Dim sbcc As String
    Dim lsaColAnhang As Collection
    Dim sFile As String, sPfad As String


    MahnwesenAusfuehren = 0

    Set ldb = openAccDBConnection("openDB")
    Set rsMW_Tickets = ldb.OpenRecordset("Select * from qryOffeneAbgelaufeneTicket;")

    If rsMW_Tickets.RecordCount = 0 Then Exit Function

    rsMW_Tickets.MoveFirst
    While Not rsMW_Tickets.EOF
            li_currMahnstufe = 0
            li_nextMahnstufe = 0
            Set rsLastMW_Ticket = ldb.OpenRecordset("Select * from tblTicketMahnwesen where TicketMahnungID = " & rsMW_Tickets("TicketMahnwesenID") & ";")

'            If rsMW_Tickets("TicketMahnwesenID") = 9959 Then
'                Debug.Print rsMW_Tickets("TicketMahnwesenID")
'            End If

            If rsLastMW_Ticket.RecordCount Then
                'Schon Mahnungen verschickt worden.
                'Ja: Die aktuelle Mahnstufe sichern
                If rsLastMW_Ticket("Datum") >= rsMW_Tickets("VerlängertBis") Then li_currMahnstufe = rsLastMW_Ticket("TicketMahnungStatusID")
            End If

            'Berechnen der zum jetzigen Zeitpunkt notwendigen Mahnstufe
            If rsMW_Tickets("Tage") >= rsMW_Tickets("ErinnerungTage") Then li_nextMahnstufe = li_nextMahnstufe + 1  'Erinnerung senden = 1
            If rsMW_Tickets("Tage") >= rsMW_Tickets("Mahnung1Tage") Then li_nextMahnstufe = li_nextMahnstufe + 1    '1. Mahnung senden = 2
            If rsMW_Tickets("Tage") >= rsMW_Tickets("Mahnung2Tage") Then li_nextMahnstufe = li_nextMahnstufe + 1    '2. Mahnung senden = 3

            For i = li_currMahnstufe + 1 To li_nextMahnstufe
                'Versenden Mahnticket (Mail)
                'Ermitteln aller relevanten Daten zum versenden der Mail (An, CC, BCC, Betreff.......)
                Set rs_EMailLoad = ldb.OpenRecordset("select * from tblTicketliste where [TicketID] = " & rsMW_Tickets("TicketID") & ";")
                'Ermitteln Mahnstufentext für Betreff
                Set rs_MahnstufeLoad = ldb.OpenRecordset("select * from tblTicketMahnungStatus where [TicketMahnungStatusID] = " & i & ";")

                'Parameter zum versenden der Mahnung laden
                If Not IsNull(rs_EMailLoad("MailCC")) Then
                        scc = rs_EMailLoad("MailCC")
                    Else
                        scc = ""
                End If
                'BCC mit der Eskalationsadresse wird nur bei der 2. Mahnung = Stufe 3 gesendet (wenn vorhanden)
                If Not IsNull(rs_EMailLoad("MailBCC")) And i = 3 Then
                        sbcc = rs_EMailLoad("MailBCC")
                    Else
                        sbcc = ""
                End If
                sBetreff = rs_MahnstufeLoad("Anzeige") & ":  " & rs_EMailLoad("MailBetreff")
                sBody = rs_EMailLoad("MailTextHTML") & "<br>" & "<br>" & "<br>" & "<br>" & rs_EMailLoad("LinkZuVorgangsOrdner")

                Set lsaColAnhang = New Collection
                Set rs_Anhaenge = ldb.OpenRecordset("select * From tblTicketAnhaenge where TicketID = " & rsMW_Tickets("TicketID") & " ;")
                While Not rs_Anhaenge.EOF
                    'Container der Anhänge füllen
                    sPfad = rs_Anhaenge("AnhangPfad")
                    n = InStrRev(rs_Anhaenge("AnhangPfad"), "\")
                    sFile = Right$(rs_Anhaenge("AnhangPfad"), Len(rs_Anhaenge("AnhangPfad")) - n)
                    lsaColAnhang.Add sPfad, sFile
                    rs_Anhaenge.MoveNext
                Wend

                'Mail für Mahnstufe versenden
                sendMailAufgabeAnhang rs_EMailLoad("MailVon"), rs_EMailLoad("MailAn"), sBetreff, sBody, lsaColAnhang, True, scc, sbcc

                'Aktuelle Mahnungsstufe in Liste in Abhängigkeit der TicketID eintragen
                rsLastMW_Ticket.AddNew
                rsLastMW_Ticket("TicketID") = rsMW_Tickets("TicketID")
                rsLastMW_Ticket("TicketMahnungStatusID") = i
                rsLastMW_Ticket("Datum") = Date
                    'Aktuelle Mahnungsstufe in Ticket eintragen
                rsMW_Tickets.Edit
                rsMW_Tickets("TicketMahnwesenID") = rsLastMW_Ticket("TicketMahnungID")
                rsMW_Tickets.Update
                rsLastMW_Ticket.Update
                'Mahnwesen TicketStatus in Ticket übernehmen  20180125-ROS
                Select Case i
                    Case 1
                        sTicketStatus = "Erinnerung"
                    Case 2
                        sTicketStatus = "1. Mahnung"
                    Case 3
                        sTicketStatus = "2. Mahnung"
                End Select
                ldb.Execute ("Update tblTicketListe set [TicketStatusID]= " & DLookup("TicketStatusID", "tblTicketStatus", "TicketStatus ='" & sTicketStatus & "'") & _
                               " where [TicketID] = " & rsMW_Tickets("TicketID") & ";")
                'Abfrage Mahnstufe und TicketType
                '2. Mahnstufe und Type = Lieferantenticket => Status setzen
                If i = 3 And rsMW_Tickets("TicketTypID") = 2 Then ' 3 = 2. Mahnung / 2 = QS Bearbeitung Lieferant
                    ldb.Execute ("Update Mastertab set [Kategorie-Status] = 'Klärung: Lieferant/Einkauf' where [ID] = " & rsMW_Tickets("ReklID") & ";")
                End If

            Next i



        rsMW_Tickets.MoveNext
    Wend

    Set rsMW_Tickets = Nothing

End Function
