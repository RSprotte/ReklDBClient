SELECT
  Mastertab.ID,
  Mastertab.TaetigkeitID,
  Mastertab.DefTaetigkeitID,
  Mastertab.[Kategorie-Status],
  ReklMemo.FertigungsDatum,
  Mastertab.[langfristige Maßnahmen],
  Mastertab.LQEinheitID,
  Mastertab.LQBereichID,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.Auftragsdatum,
  ReklMemo.VorgangstypID,
  Mastertab.intern,
  Mastertab.extern
FROM
  Mastertab
  LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID;
