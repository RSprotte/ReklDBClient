SELECT
  Mastertab.ID,
  Mastertab.[Kunden-ID],
  ReklMemo.VorgangstypID,
  ReklMemo.Vorgangs_Eingang,
  Mastertab.Maschinennummer,
  Mastertab.[Bestellnummer Kunde],
  Mastertab.Rückmeldung_Kunde_Soll,
  Mastertab.Auftragsdatum,
  Mastertab.Referenznummer,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.Referenznummer,
  Mastertab.Fehlerpositionen,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.ID,
  Mastertab.Rückmeldung_Kunde_Verschiebung,
  Mastertab.[per Telefon],
  Mastertab.[per E-Mail/Fax],
  Mastertab.[mit Rücksendung],
  Mastertab.[Welcher Fehler tritt auf]
FROM
  Mastertab
  LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID
ORDER BY
  Mastertab.ID;
