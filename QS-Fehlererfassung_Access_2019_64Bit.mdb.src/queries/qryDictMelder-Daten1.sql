SELECT
  Mastertab.ID,
  Mastertab.[Kunden-ID],
  Mastertab.Auftragsdatum,
  ReklMemo.Vorgangs_Eingang,
  Mastertab.[Kategorie-Status],
  ReklMemo.KundenRefNummer,
  Mastertab.[Bestellnummer Kunde],
  ReklMemo.BestellnrRepAuftrag,
  ReklMemo.AnlieferterminRep,
  Mastertab.[Zuständiger Service-MA],
  Mastertab.Datum,
  Mastertab.Rückmeldung_Kunde_Soll,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.Fehlerpositionen,
  Mastertab.Referenznummer,
  Mastertab.[per Telefon],
  Mastertab.[per E-Mail/Fax],
  Mastertab.[mit Rücksendung],
  ReklMemo.VorgangstypID,
  ReklMemo.Portal,
  ReklMemo.[Portal-Link],
  Mastertab.intern,
  Mastertab.extern,
  Mastertab.Rückmeldung_Kunde_Verschiebung
FROM
  Mastertab
  LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID;
