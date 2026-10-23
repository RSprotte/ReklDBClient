SELECT
  [Artikel-all_avista_out].*,
  [Artikel-all_avista_out].[Artikel-ID],
  [Artikel-all_avista_out].[LQ Art-Nr],
  [Artikel-all_avista_out].Variante,
  [Artikel-all_avista_out].[Kd Mat-Nr],
  [Artikel-all_avista_out].[Bezeichnung 1],
  [LQ Art-Nr] & "-" & [Variante] AS ArtNr
FROM
  [Artikel-all_avista_out]
WHERE
  (
    (
      ([Artikel-all_avista_out].QZV) = True
    )
  );
