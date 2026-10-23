SELECT
  Mastertab.ID,
  Format([Auftragsdatum], "mmm") AS Monat,
  Format([Auftragsdatum], "yyyy") AS Jahr,
  Vorgangstyp.Vorgangstyp,
  Mastertab.[Kategorie-Status],
  Mastertab.[Kategorie-Problem],
  Mastertab.extern,
  Sum(AufwandVorgangPositionen.Zeit) AS SummevonZeit,
  Mastertab.[Material-Kosten] AS Mat_Kosten,
  [Zeitbedarf] * 45 + [Material-Kosten] + [Fracht-Kosten] + [Prüfkosten] AS Ges_Kosten
FROM
  (
    (
      Mastertab
      LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID
    )
    LEFT JOIN Vorgangstyp ON ReklMemo.VorgangstypID = Vorgangstyp.VorgangstypID
  )
  LEFT JOIN AufwandVorgangPositionen ON Mastertab.ID = AufwandVorgangPositionen.MID
GROUP BY
  Mastertab.ID,
  Format([Auftragsdatum], "mmm"),
  Format([Auftragsdatum], "yyyy"),
  Vorgangstyp.Vorgangstyp,
  Mastertab.[Kategorie-Status],
  Mastertab.[Kategorie-Problem],
  Mastertab.extern,
  Mastertab.[Material-Kosten],
  [Zeitbedarf] * 45 + [Material-Kosten] + [Fracht-Kosten] + [Prüfkosten]
HAVING
  (
    (
      (Mastertab.extern) = True
    )
  )
ORDER BY
  Mastertab.ID DESC;
