SELECT
  Mastertab.ID,
  Mastertab.LQEinheitID,
  Mastertab.LQBereichID,
  Mastertab.Sachbearbeiter,
  Mastertab.[Welcher Fehler tritt auf],
  ReklMemo.QSBearbeitungBis_Verl,
  ReklMemo.QSBearbeitungBis_Ist,
  ReklMemo.QSBearbeitungBis_Soll,
  Mastertab.[Kategorie-Status]
FROM
  Mastertab
  INNER JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID;
