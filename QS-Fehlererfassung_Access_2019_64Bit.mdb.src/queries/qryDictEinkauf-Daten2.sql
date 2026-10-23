SELECT
  Mastertab.ID,
  Mastertab.[Kategorie-Status],
  ROS_Daten_Einkauf.Prüfzyklus,
  ROS_Daten_Einkauf.ZeitaufwandjePrüfung,
  ROS_Daten_Einkauf.AnsprechpLieferant,
  ROS_Daten_Einkauf.Fazit,
  ROS_Daten_Einkauf.NichtabgeschlosseneMaßnahmendesLieferanten,
  ROS_Daten_Einkauf.ÜberprüfungsterminLQ,
  ROS_Daten_Einkauf.ÜberprüfungsterminErgebnis,
  ROS_Daten_Einkauf.Wiederholungsfehler,
  ROS_Daten_Einkauf.TaetigkeitLieferantID,
  Mastertab.[Lieferanten-Code 1],
  Mastertab.[Meldungsnummer Kunde],
  ReklMemo.VorgangstypID,
  Mastertab.Auftragsdatum,
  Mastertab.intern,
  Mastertab.extern
FROM
  (
    Mastertab
    LEFT JOIN ROS_Daten_Einkauf ON Mastertab.ID = ROS_Daten_Einkauf.[ID-ReklNr]
  )
  LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID;
