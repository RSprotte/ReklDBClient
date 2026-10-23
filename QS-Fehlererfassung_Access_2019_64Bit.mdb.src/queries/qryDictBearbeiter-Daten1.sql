SELECT
  Mastertab.ID,
  ReklMemo.[Aktueller Eintrag],
  Mastertab.[Bemerkungen - Maßnahmen],
  Mastertab.[Kategorie-Status],
  ReklMemo.Protokoll,
  Mastertab.[Meldungsnummer Kunde],
  ReklMemo.VorgangstypID,
  Mastertab.intern,
  Mastertab.extern,
  Mastertab.Auftragsdatum,
  Mastertab.[Bemerkungen-Vorgang]
FROM
  Mastertab
  LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID;
