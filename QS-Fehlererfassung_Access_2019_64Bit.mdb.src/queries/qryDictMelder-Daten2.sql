SELECT
  Mastertab.ID,
  Mastertab.Fehlergruppe,
  Mastertab.Fehlertexte,
  Mastertab.Maschinennummer,
  Mastertab.[Aufwand-TRUMPF],
  Mastertab.[Welcher Fehler tritt auf],
  Mastertab.[Bemerkungen - Maßnahmen],
  Mastertab.[Kategorie-Status],
  Mastertab.Fehlerpositionen,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.intern,
  Mastertab.extern,
  Mastertab.Auftragsdatum,
  ReklMemo.VorgangstypID
FROM
  Mastertab
  LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID;
