SELECT TOP 10
  [Baugruppen-avista_out].[LQ Art-Nr],
  [Baugruppen-avista_out].[Kunden-Nr],
  [Baugruppen-avista_out].[Bezeichnung 1],
  Count(Mastertab.ID) AS AnzahlvonID
FROM
  [Baugruppen-avista_out]
  LEFT JOIN Mastertab ON [Baugruppen-avista_out].[Baugruppen-ID] = Mastertab.[Baugruppen-ID]
WHERE
  (
    (
      (Mastertab.Auftragsdatum) >= #7/1/2015#
    )
    AND (
      (Mastertab.extern) = Yes
    )
  )
GROUP BY
  [Baugruppen-avista_out].[LQ Art-Nr],
  [Baugruppen-avista_out].[Kunden-Nr],
  [Baugruppen-avista_out].[Bezeichnung 1],
  Mastertab.[Kategorie-Problem]
HAVING
  (
    (
      (Mastertab.[Kategorie-Problem]) = "berechtigt"
    )
  )
ORDER BY
  Count(Mastertab.ID) DESC;
