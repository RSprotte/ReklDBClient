SELECT
  Mastertab.ID,
  Mastertab.Auftragsdatum,
  Verursacher.Anzeige,
  Mastertab.[Kategorie-Bearbeitung I],
  Mastertab.[Kategorie-Bearbeitung II],
  Mastertab.[Kategorie-Problem],
  Mastertab.[Material-Kosten],
  Mastertab.Auftragsnummer,
  ROS_Daten_Einkauf.InterneDaten_ROS,
  Mastertab.[Belastungsbetrag LQ]
FROM
  (
    (
      Mastertab
      LEFT JOIN Verursacher ON Mastertab.Verursacher = Verursacher.Verursacher
    )
    LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID
  )
  LEFT JOIN ROS_Daten_Einkauf ON Mastertab.ID = ROS_Daten_Einkauf.[ID-ReklNr]
WHERE
  (
    (
      (
        Mastertab.[Kategorie-Bearbeitung I]
      ) Is Not Null
    )
    AND (
      (
        ROS_Daten_Einkauf.InterneDaten_ROS
      ) = False
    )
    AND (
      (Mastertab.extern) = True
    )
    AND (
      (Mastertab.[Kategorie-Status]) = "abgeschlossen"
    )
    AND (
      (ReklMemo.VorgangstypID) = 1
    )
  )
ORDER BY
  Mastertab.ID;
