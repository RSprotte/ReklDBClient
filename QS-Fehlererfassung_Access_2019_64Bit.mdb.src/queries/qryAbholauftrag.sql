SELECT
  Mastertab.ID,
  tblFirma.KndNr_Avista,
  tblFirma.Firma_Name,
  [Artikel-all_avista_out].[LQ Art-Nr] AS LQArtOut,
  [Artikel-all_avista_out].Variante,
  [Artikel-all_avista_out].[Bezeichnung 1],
  Mastertab.Auftragsdatum,
  [Baugruppen-avista_out].[LQ Art-Nr] AS BG,
  [Baugruppen-avista_out].[Bezeichnung 1] AS [BG Bezeichnung],
  Kunden.Kontaktperson,
  Mastertab.[Meldungsnummer Kunde],
  RSP_ReklTag.Tag,
  [Baugruppen-avista_out].[Kunden-Nr] AS BGKD,
  [Artikel-all_avista_out].[Kd Mat-Nr] AS ArtKD,
  [LQArtOut] & "-" & [Variante] AS LQArt
FROM
  (
    (
      (
        (
          Mastertab
          LEFT JOIN [Artikel-all_avista_out] ON Mastertab.[Artikel-ID 1] = [Artikel-all_avista_out].[Artikel-ID]
        )
        LEFT JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]
      )
      LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
    )
    LEFT JOIN [Baugruppen-avista_out] ON Mastertab.[Baugruppen-ID] = [Baugruppen-avista_out].[Baugruppen-ID]
  )
  LEFT JOIN (
    RSP_ReklTag
    RIGHT JOIN ReklIDtoTagID ON RSP_ReklTag.TagID = ReklIDtoTagID.TagID
  ) ON Mastertab.ID = ReklIDtoTagID.ReklID
WHERE
  (
    (
      (Mastertab.extern) = True
    )
  )
GROUP BY
  Mastertab.ID,
  tblFirma.KndNr_Avista,
  tblFirma.Firma_Name,
  [Artikel-all_avista_out].[LQ Art-Nr],
  [Artikel-all_avista_out].Variante,
  [Artikel-all_avista_out].[Bezeichnung 1],
  Mastertab.Auftragsdatum,
  [Baugruppen-avista_out].[LQ Art-Nr],
  [Baugruppen-avista_out].[Bezeichnung 1],
  Kunden.Kontaktperson,
  Mastertab.[Meldungsnummer Kunde],
  RSP_ReklTag.Tag,
  [Baugruppen-avista_out].[Kunden-Nr],
  [Artikel-all_avista_out].[Kd Mat-Nr]
HAVING
  (
    (
      (tblFirma.KndNr_Avista) Is Not Null
    )
  );
