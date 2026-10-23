SELECT
  Mastertab.intern,
  Mastertab.extern,
  Mastertab.ID,
  Format([Auftragsdatum], "mmm") AS Monat,
  Format([Auftragsdatum], "yyyy") AS Jahr,
  Mastertab.[Kategorie-Problem],
  Mastertab.[Kategorie-Status],
  Mastertab.Verursacher,
  Lieferanten.Lieferantenname,
  LQEinheit.LQEinheit,
  Mastertab.[Aufwand-TRUMPF],
  Mastertab.Zeitbedarf,
  Mastertab.[QS-Bearbeitung],
  Mastertab.[PE-Bearbeitung],
  Mastertab.[Material-Kosten],
  Mastertab.[Fracht-Kosten],
  [Artikel-all_avista_out].[LQ Art-Nr],
  [Artikel-all_avista_out].AvistaArtikelTyp,
  Lieferanten.VerantwortlichEinkauf,
  Taetigkeiten.Taetigkeit
FROM
  (
    (
      (
        (
          Mastertab
          LEFT JOIN Lieferanten ON Mastertab.[Lieferanten-Code 1] = Lieferanten.[Lieferant-ID]
        )
        LEFT JOIN LQEinheit ON Mastertab.LQEinheitID = LQEinheit.ID
      )
      LEFT JOIN [Artikel-all_avista_out] ON Mastertab.[Artikel-ID 1] = [Artikel-all_avista_out].[Artikel-ID]
    )
    LEFT JOIN ROS_Daten_Einkauf ON Mastertab.ID = ROS_Daten_Einkauf.[ID-ReklNr]
  )
  LEFT JOIN Taetigkeiten ON ROS_Daten_Einkauf.TaetigkeitLieferantID = Taetigkeiten.ID
WHERE
  (
    (
      (Mastertab.Auftragsdatum) Between #1/1/2014#
      And Date()
    )
  )
GROUP BY
  Mastertab.intern,
  Mastertab.extern,
  Mastertab.ID,
  Format([Auftragsdatum], "mmm"),
  Format([Auftragsdatum], "yyyy"),
  Mastertab.[Kategorie-Problem],
  Mastertab.[Kategorie-Status],
  Mastertab.Verursacher,
  Lieferanten.Lieferantenname,
  LQEinheit.LQEinheit,
  Mastertab.[Aufwand-TRUMPF],
  Mastertab.Zeitbedarf,
  Mastertab.[QS-Bearbeitung],
  Mastertab.[PE-Bearbeitung],
  Mastertab.[Material-Kosten],
  Mastertab.[Fracht-Kosten],
  [Artikel-all_avista_out].[LQ Art-Nr],
  [Artikel-all_avista_out].AvistaArtikelTyp,
  Lieferanten.VerantwortlichEinkauf,
  Taetigkeiten.Taetigkeit
HAVING
  (
    (
      (Mastertab.[Kategorie-Problem]) = "berechtigt"
      Or (Mastertab.[Kategorie-Problem]) = "nicht berechtigt"
    )
  )
ORDER BY
  Mastertab.ID DESC;
