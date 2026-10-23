SELECT
  [Baugruppen-avista_out].[Baugruppen-ID],
  Count(Mastertab.ReklIDA) AS [Anzahl Reklamationen]
FROM
  [Baugruppen-avista_out]
  LEFT JOIN Mastertab ON [Baugruppen-avista_out].[Baugruppen-ID] = Mastertab.[Baugruppen-ID]
WHERE
  (
    (
      (Mastertab.Auftragsdatum) Between Date()
      And DateAdd(
        "yyyy",
        -1,
        Date()
      )
    )
    AND (
      (Mastertab.[Kategorie-Problem]) = "berechtigt"
    )
    AND (
      (Mastertab.extern) = Yes
    )
  )
GROUP BY
  [Baugruppen-avista_out].[Baugruppen-ID]
ORDER BY
  Count(Mastertab.ReklIDA) DESC;
