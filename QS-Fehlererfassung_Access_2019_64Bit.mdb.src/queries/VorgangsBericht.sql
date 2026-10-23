SELECT
  Mastertab.ID,
  [Artikel-all_avista_out].[LQ Art-Nr],
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.[Bestellnummer Kunde],
  tblFirma.Firma_Name,
  Kunden.Kontaktperson
FROM
  (
    (
      Mastertab
      LEFT JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]
    )
    LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
  )
  LEFT JOIN [Artikel-all_avista_out] ON Mastertab.[Artikel-ID 1] = [Artikel-all_avista_out].[Artikel-ID]
WHERE
  (
    (
      (Mastertab.Datum) = #1/1/9999#
    )
  )
GROUP BY
  Mastertab.ID,
  [Artikel-all_avista_out].[LQ Art-Nr],
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.[Bestellnummer Kunde],
  tblFirma.Firma_Name,
  Kunden.Kontaktperson;
