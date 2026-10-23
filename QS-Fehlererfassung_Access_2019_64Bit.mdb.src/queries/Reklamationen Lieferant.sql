SELECT
  Mastertab.*,
  Mastertab.Datum
FROM
  Lieferanten
  INNER JOIN Mastertab ON Lieferanten.[Lieferant-ID] = Mastertab.[Lieferanten-Code 1]
WHERE
  (
    (
      (Mastertab.Datum) >= #7/1/2009#
    )
    AND (
      (Mastertab.[Lieferanten-Code 1]) = 50
    )
    AND (
      (Lieferanten.[L-Nummer/AVISTA]) = 70025
    )
  );
