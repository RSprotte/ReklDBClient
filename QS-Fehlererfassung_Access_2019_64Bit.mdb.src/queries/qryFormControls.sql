SELECT
  Mastertab.ID,
  Mastertab.Auftragsdatum,
  Mastertab.extern,
  Mastertab.intern,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.[Kategorie-Status],
  ReklMemo.VorgangstypID
FROM
  Mastertab
  LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID;
