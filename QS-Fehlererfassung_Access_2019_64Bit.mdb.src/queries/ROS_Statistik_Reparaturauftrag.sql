SELECT
  Mastertab.ID,
  Format([Auftragsdatum], "mmm") AS Monat,
  Mastertab.Auftragsdatum,
  Vorgangstyp.Vorgangstyp,
  Mastertab.[Kategorie-Status],
  Mastertab.[Kategorie-Bearbeitung I],
  Mastertab.[Kategorie-Bearbeitung II],
  Sum(AufwandVorgangPositionen.Zeit) AS SummevonZeit,
  Mastertab.[Material-Kosten] AS Mat_Kosten,
  [Zeitbedarf] * 45 + [Material-Kosten] + [Fracht-Kosten] + [Prüfkosten] AS Ges_Kosten,
  Mastertab.Reparaturbetrag,
  [Reparaturbetrag] -(
    [Zeitbedarf] * 45 + [Material-Kosten] + [Fracht-Kosten] + [Prüfkosten]
  ) AS Differenz
FROM
  (
    (
      Mastertab
      LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID
    )
    LEFT JOIN Vorgangstyp ON ReklMemo.VorgangstypID = Vorgangstyp.VorgangstypID
  )
  LEFT JOIN AufwandVorgangPositionen ON Mastertab.ID = AufwandVorgangPositionen.MID
WHERE
  (
    (
      (Mastertab.extern) = True
    )
  )
GROUP BY
  Mastertab.ID,
  Format([Auftragsdatum], "mmm"),
  Mastertab.Auftragsdatum,
  Vorgangstyp.Vorgangstyp,
  Mastertab.[Kategorie-Status],
  Mastertab.[Kategorie-Bearbeitung I],
  Mastertab.[Kategorie-Bearbeitung II],
  Mastertab.[Material-Kosten],
  [Zeitbedarf] * 45 + [Material-Kosten] + [Fracht-Kosten] + [Prüfkosten],
  Mastertab.Reparaturbetrag,
  [Reparaturbetrag] -(
    [Zeitbedarf] * 45 + [Material-Kosten] + [Fracht-Kosten] + [Prüfkosten]
  )
HAVING
  (
    (
      (Vorgangstyp.Vorgangstyp) = "Reparatur"
    )
    And (
      (Mastertab.[Kategorie-Status]) = "Abgeschlossen"
    )
    And (
      (
        Format$(Mastertab.Auftragsdatum, 'yyyy')
      ) = [Bitte geben Sie das Jahr an]
    )
  )
ORDER BY
  Mastertab.ID DESC;
