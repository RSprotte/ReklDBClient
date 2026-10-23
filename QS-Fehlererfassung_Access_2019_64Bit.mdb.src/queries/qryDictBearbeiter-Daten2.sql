SELECT
  Mastertab.ID,
  Mastertab.[Wie Warum passierte der Fehler],
  Mastertab.Sachbearbeiter,
  Mastertab.TaetigkeitID,
  Mastertab.DefTaetigkeitID,
  Mastertab.DefTaetigkeitAnalyseID,
  Mastertab.[Fertigungsauftrag-FA],
  ReklMemo.FertigungsDatum,
  Mastertab.LQEinheitID,
  Mastertab.LQBereichID,
  Mastertab.LQWorkstationID,
  ReklMemo.Lagerpüfung,
  Mastertab.Verursacher,
  Mastertab.[Lieferanten-Code 1],
  Mastertab.[Rückliefer-Nummer],
  Mastertab.Zuständig,
  Mastertab.[Ablage Rücklieferschein],
  Mastertab.[Mitarbeiter-Kurzzeichen],
  Mastertab.[LiefAnsprechp-ID],
  ROS_Daten_Einkauf.TaetigkeitLieferantID,
  Mastertab.[Kategorie-Status],
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.Auftragsdatum,
  ReklMemo.VorgangstypID,
  Mastertab.intern,
  Mastertab.extern,
  Mastertab.[Kategorie-Problem]
FROM
  (
    Mastertab
    LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID
  )
  LEFT JOIN ROS_Daten_Einkauf ON Mastertab.ID = ROS_Daten_Einkauf.[ID-ReklNr];
