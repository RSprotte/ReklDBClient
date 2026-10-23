SELECT
  Lieferanten.[Lieferant-ID],
  Lieferanten.Matchcode,
  Lieferanten.Lieferantenname,
  Lieferanten.[L-Nummer/AVISTA],
  Lieferanten.[RMA-Nummer]
FROM
  Lieferanten
WHERE
  (
    (
      (Lieferanten.GESPERRT) = No
    )
  )
ORDER BY
  Lieferanten.Matchcode;
