SELECT
  Mastertab.ID
FROM
  Mastertab
  LEFT JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]
WHERE
  (
    (
      (Mastertab.extern) = True
    )
    AND (
      (Mastertab.[Kategorie-Status]) <> "Stellungnahme Klärung Kunde"
    )
    AND (
      (Kunden.Kontaktperson) = "Kraus, Fabian"
    )
  )
GROUP BY
  Mastertab.ID
HAVING
  (
    (
      (Mastertab.ID) = 83318
    )
  );
