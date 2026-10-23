SELECT
  Mastertab.ID,
  Mastertab.[Kategorie-Status],
  Mastertab.[Kategorie-Problem],
  Mastertab.Rückmeldung_Lieferant_Soll,
  Mastertab.Rückmeldung_Lieferant_Ist,
  Mastertab.Rückmeldung_Kunde_Ist,
  Mastertab.Rückmeldung_Kunde_Soll,
  Mastertab.Rückmeldung_Kunde_Verschiebung,
  Mastertab.[Rücklieferung an Lieferant],
  Mastertab.[Rückliefer-Menge],
  Mastertab.[Annahme mit Prüfausnahme],
  Mastertab.[Annahme mit Nacharbeit],
  Mastertab.[Meldungsnummer Kunde],
  ReklMemo.VorgangstypID,
  Mastertab.Auftragsdatum,
  Mastertab.intern,
  Mastertab.extern
FROM
  Mastertab
  LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID;
