SELECT
  Mastertab.ID,
  [Baugruppen-avista_out].[LQ Art-Nr],
  [Baugruppen-avista_out].[Var Nr],
  [Artikel-all_avista_out].[LQ Art-Nr],
  [Artikel-all_avista_out].Variante,
  Mastertab.[LINK-Nr - UBG],
  [Baugruppen-avista_out].[LQ Art-Nr] & "-" & [Baugruppen-avista_out].[Var Nr] AS LQBGNr,
  [Artikel-all_avista_out].[LQ Art-Nr] & "-" & [Artikel-all_avista_out].Variante AS LQArtNr
FROM
  (
    Mastertab
    LEFT JOIN [Baugruppen-avista_out] ON Mastertab.[Baugruppen-ID] = [Baugruppen-avista_out].[Baugruppen-ID]
  )
  LEFT JOIN [Artikel-all_avista_out] ON Mastertab.[Artikel-ID 1] = [Artikel-all_avista_out].[Artikel-ID];
