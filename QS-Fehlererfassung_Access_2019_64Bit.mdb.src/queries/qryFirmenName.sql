SELECT
  Kunden.[Kunden-ID],
  tblFirma.Firma_Name,
  tblFirma.KndNr_Avista
FROM
  Kunden
  INNER JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
GROUP BY
  Kunden.[Kunden-ID],
  tblFirma.Firma_Name,
  tblFirma.KndNr_Avista;
