SELECT
  Fehlergruppen.Fehlergruppe,
  Fehlertexte.Fehlertexte,
  Mastertab.ID,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.Auftragsdatum,
  tblFirma.Firma_Name,
  Mastertab.extern,
  Mastertab.[TRUMPF?]
FROM
  (
    Kunden
    RIGHT JOIN (
      Fehlertexte
      RIGHT JOIN (
        Fehlergruppen
        RIGHT JOIN Mastertab ON Fehlergruppen.ID = Mastertab.Fehlergruppe
      ) ON Fehlertexte.ID = Mastertab.Fehlertexte
    ) ON Kunden.[Kunden-ID] = Mastertab.[Kunden-ID]
  )
  LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
WHERE
  (
    (
      (Mastertab.Auftragsdatum) Between #1/1/2009#
      And #6/30/2009#
    )
    AND (
      (tblFirma.Firma_Name) = "tw"
    )
    AND (
      (Mastertab.extern) = Yes
    )
    AND (
      (Mastertab.[TRUMPF?]) = Yes
    )
  )
ORDER BY
  Mastertab.Auftragsdatum DESC;
