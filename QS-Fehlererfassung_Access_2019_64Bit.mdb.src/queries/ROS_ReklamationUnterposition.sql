SELECT
  Mastertab.ID,
  Mastertab.[Kategorie-Status],
  Mastertab.Referenznummer,
  Mastertab.Fehlerpositionen
FROM
  Mastertab
WHERE
  (
    (
      (Mastertab.ID) Is Not Null
    )
  )
GROUP BY
  Mastertab.ID,
  Mastertab.[Kategorie-Status],
  Mastertab.Referenznummer,
  Mastertab.Fehlerpositionen
HAVING
  (
    (
      (Mastertab.Referenznummer) Is Not Null
    )
  )
ORDER BY
  Mastertab.ID DESC;
