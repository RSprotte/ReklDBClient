SELECT
  Lieferanten.Lieferantenname,
  "L" & Format([Lieferant-ID], "000") AS LCode,
  Lieferanten.[L-Nummer/AVISTA]
FROM
  Lieferanten
WHERE
  (
    (
      (Lieferanten.GESPERRT) = No
    )
  )
ORDER BY
  Lieferanten.Lieferantenname;
