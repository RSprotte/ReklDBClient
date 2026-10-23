SELECT
  Mastertab.ID,
  Mastertab.Datum,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.Fehlertexte,
  Kunden.[Kunden-ID],
  Kunden.Kontaktperson
FROM
  Kunden
  INNER JOIN Mastertab ON Kunden.[Kunden-ID] = Mastertab.[Kunden-ID]
ORDER BY
  Kunden.Kontaktperson;
