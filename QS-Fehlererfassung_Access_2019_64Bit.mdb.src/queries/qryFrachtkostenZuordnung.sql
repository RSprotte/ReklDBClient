SELECT
  Mastertab.ID,
  Mastertab.[WE-Datum],
  tblFirma.Firma_Name,
  [Baugruppen-avista_out].[LQ Art-Nr],
  [Baugruppen-avista_out].[Bezeichnung 1],
  [Baugruppen-avista_out].[Kunden-Nr],
  [Artikel-all_avista_out].[LQ Art-Nr],
  [Artikel-all_avista_out].[Bezeichnung 1],
  [Artikel-all_avista_out].[Kd Mat-Nr]
FROM
  (
    (
      (
        Mastertab
        LEFT JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]
      )
      LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
    )
    LEFT JOIN [Artikel-all_avista_out] ON Mastertab.[Artikel-ID 1] = [Artikel-all_avista_out].[Artikel-ID]
  )
  LEFT JOIN [Baugruppen-avista_out] ON Mastertab.[Baugruppen-ID] = [Baugruppen-avista_out].[Baugruppen-ID]
WHERE
  (
    (
      (Mastertab.extern) = True
    )
  )
GROUP BY
  Mastertab.ID,
  Mastertab.[WE-Datum],
  tblFirma.Firma_Name,
  [Baugruppen-avista_out].[LQ Art-Nr],
  [Baugruppen-avista_out].[Bezeichnung 1],
  [Baugruppen-avista_out].[Kunden-Nr],
  [Artikel-all_avista_out].[LQ Art-Nr],
  [Artikel-all_avista_out].[Bezeichnung 1],
  [Artikel-all_avista_out].[Kd Mat-Nr]
HAVING
  (
    (
      (Mastertab.[WE-Datum]) Between [Von Datum]
      And Date()
    )
    AND (
      (tblFirma.Firma_Name) Like [Anfangsbuchstabe Kunde] & "*"
    )
  )
ORDER BY
  Mastertab.[WE-Datum];
