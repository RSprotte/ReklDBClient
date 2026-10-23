Attribute VB_Name = "UpDate4_0"
Option Compare Database
Option Explicit

Function AufwandslisteNeu()

    Dim ldb As DAO.Database
    Dim rsUD As DAO.Recordset
    Dim ltbdef As DAO.TableDef
    Dim lVorgangID As Long
    Dim i As Integer

     Set ldb = openAccDBConnection("openDB")
    '1. Vorgang ""Fehlerbericht erfasst." anlegen um die Pos10 einttragen zu können
    Set rsUD = ldb.OpenRecordset("select * from [Vorgang-Text];")
    rsUD.AddNew
    rsUD("Vorgang-Beschreibung") = "Fehlerbericht erfasst."
    lVorgangID = rsUD("Vorgang-ID")
    rsUD.Update
    Set rsUD = Nothing
    '2. Pos 10 übertragen
    ldb.Execute ("INSERT INTO AufwandVorgangPositionen ( MID, Datum, Pos, Zeit, Vorgang, Kürzel ) " & _
                "SELECT Mastertab.ID, Mastertab.Datum, 10 AS Pos, 0.25 AS Zeit, " & lVorgangID & " AS Vorgang, Bearbeiter.Kurzzeichen " & _
                "FROM Mastertab INNER JOIN Bearbeiter ON Mastertab.[Zuständiger Service-MA] = Bearbeiter.Name;")

    '2. Pos 20 - 80 übertragen
    ldb.Execute ("INSERT INTO AufwandVorgangPositionen ( MID, Pos, Datum, Artikel, Menge, Zeit, Vorgang, Kürzel ) " & _
                "SELECT Mastertab.ID, 20 AS Pos, Mastertab.[Datum - Pos20], Mastertab.[Art-Nr - Pos20], Mastertab.[Menge - Pos20], Mastertab.[Zeit - Pos20], " & _
                "Mastertab.[Vorgang - Pos20], Mastertab.[Kurzzeichen - Pos20] " & _
                "From Mastertab " & _
                "WHERE (((Mastertab.[Datum - Pos20]) Is Not Null));")
    ldb.Execute ("INSERT INTO AufwandVorgangPositionen ( MID, Pos, Datum, Artikel, Menge, Zeit, Vorgang, Kürzel ) " & _
                "SELECT Mastertab.ID, 30 AS Pos, Mastertab.[Datum - Pos30], Mastertab.[Art-Nr - Pos30], Mastertab.[Menge - Pos30], Mastertab.[Zeit - Pos30], " & _
                "Mastertab.[Vorgang - Pos30], Mastertab.[Kurzzeichen - Pos30] " & _
                "From Mastertab " & _
                "WHERE (((Mastertab.[Datum - Pos30]) Is Not Null));")
    ldb.Execute ("INSERT INTO AufwandVorgangPositionen ( MID, Pos, Datum, Artikel, Menge, Zeit, Vorgang, Kürzel ) " & _
                "SELECT Mastertab.ID, 40 AS Pos, Mastertab.[Datum - Pos40], Mastertab.[Art-Nr - Pos40], Mastertab.[Menge - Pos40], Mastertab.[Zeit - Pos40], " & _
                "Mastertab.[Vorgang - Pos40], Mastertab.[Kurzzeichen - Pos40] " & _
                "From Mastertab " & _
                "WHERE (((Mastertab.[Datum - Pos40]) Is Not Null));")
    ldb.Execute ("INSERT INTO AufwandVorgangPositionen ( MID, Pos, Datum, Artikel, Menge, Zeit, Vorgang, Kürzel ) " & _
                "SELECT Mastertab.ID, 50 AS Pos, Mastertab.[Datum - Pos50], Mastertab.[Art-Nr - Pos50], Mastertab.[Menge - Pos50], Mastertab.[Zeit - Pos50], " & _
                "Mastertab.[Vorgang - Pos50], Mastertab.[Kurzzeichen - Pos50] " & _
                "From Mastertab " & _
                "WHERE (((Mastertab.[Datum - Pos50]) Is Not Null));")
    ldb.Execute ("INSERT INTO AufwandVorgangPositionen ( MID, Pos, Datum, Artikel, Menge, Zeit, Vorgang, Kürzel ) " & _
                "SELECT Mastertab.ID, 60 AS Pos, Mastertab.[Datum - Pos60], Mastertab.[Art-Nr - Pos60], Mastertab.[Menge - Pos60], Mastertab.[Zeit - Pos60], " & _
                "Mastertab.[Vorgang - Pos60], Mastertab.[Kurzzeichen - Pos60] " & _
                "From Mastertab " & _
                "WHERE (((Mastertab.[Datum - Pos60]) Is Not Null));")
    ldb.Execute ("INSERT INTO AufwandVorgangPositionen ( MID, Pos, Datum, Artikel, Menge, Zeit, Vorgang, Kürzel ) " & _
                "SELECT Mastertab.ID, 70 AS Pos, Mastertab.[Datum - Pos70], Mastertab.[Art-Nr - Pos70], Mastertab.[Menge - Pos70], Mastertab.[Zeit - Pos70], " & _
                "Mastertab.[Vorgang - Pos70], Mastertab.[Kurzzeichen - Pos70] " & _
                "From Mastertab " & _
                "WHERE (((Mastertab.[Datum - Pos70]) Is Not Null));")
    ldb.Execute ("INSERT INTO AufwandVorgangPositionen ( MID, Pos, Datum, Artikel, Menge, Zeit, Vorgang, Kürzel ) " & _
                "SELECT Mastertab.ID, 80 AS Pos, Mastertab.[Datum - Pos80], Mastertab.[Art-Nr - Pos80], Mastertab.[Menge - Pos80], Mastertab.[Zeit - Pos80], " & _
                "Mastertab.[Vorgang - Pos80], Mastertab.[Kurzzeichen - Pos80] " & _
                "From Mastertab " & _
                "WHERE (((Mastertab.[Datum - Pos80]) Is Not Null));")


    Set ltbdef = ldb.TableDefs("Mastertab")
    For i = 2 To 8
        ltbdef.Fields.Delete "Datum - Pos" & i & "0"
        ltbdef.Fields.Delete "Art-Nr - Pos" & i & "0"
        ltbdef.Fields.Delete "Menge - Pos" & i & "0"
        ltbdef.Fields.Delete "Zeit - Pos" & i & "0"
        ltbdef.Fields.Delete "Vorgang - Pos" & i & "0"
        ltbdef.Fields.Delete "Kurzzeichen - Pos" & i & "0"
    Next i

     ' Set ldb = openAccDBConnection("closeDB")

    Set rsUD = Nothing


End Function


Function RenameReklBericht()

Dim ldb As DAO.Database
Dim rs As DAO.Recordset
Dim sPfadDatei As String
Dim afiles As Variant
Dim i As Integer
Dim sReplace As String
Dim FileName As String

 Set ldb = openAccDBConnection("openDB")
Set rs = ldb.OpenRecordset("SELECT Mastertab.ID, Mastertab.[Meldungsnummer Kunde], Mastertab.Auftragsdatum From Mastertab WHERE ((Not (Mastertab.[Kategorie-Status])='Abgeschlossen') AND ((Mastertab.Verursacher)='Lieferant'));")

rs.MoveFirst
While Not rs.EOF
    If IsNull(rs("Meldungsnummer Kunde")) Or rs("Meldungsnummer Kunde") = "" Then
        sPfadDatei = setReklVerzeichnis(rs("ID"), format(rs("Auftragsdatum"), "yyyy"))
    Else
        sPfadDatei = setReklVerzeichnis(rs("ID"), format(rs("Auftragsdatum"), "yyyy"), rs("Meldungsnummer Kunde"))
    End If

    FileName = Dir(sPfadDatei & "*.*")

    Do While FileName <> ""
        If InStr(1, FileName, "Reklmationsbericht_Lieferant") Then
            sReplace = Replace(FileName, "Reklmationsbericht_Lieferant", "Reklamationsbericht_Lieferant")
            Name sPfadDatei & FileName As sPfadDatei & sReplace
        End If
        FileName = Dir()
    Loop

    rs.MoveNext
Wend

Set rs = Nothing
 ' Set ldb = openAccDBConnection("closeDB")

End Function
