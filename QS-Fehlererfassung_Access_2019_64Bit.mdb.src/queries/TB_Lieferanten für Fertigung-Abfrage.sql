SELECT
  Lieferanten.[Lieferant-ID] AS [L-Code],
  Lieferanten.Matchcode,
  Lieferanten.Lieferantenname,
  Lieferanten.[L-Nummer/AVISTA]
FROM
  Lieferanten
WHERE
  (
    (
      (Lieferanten.GESPERRT) = No
    )
    AND (
      (Lieferanten.Anzeigen) = Yes
    )
  )
ORDER BY
  Lieferanten.Matchcode;
