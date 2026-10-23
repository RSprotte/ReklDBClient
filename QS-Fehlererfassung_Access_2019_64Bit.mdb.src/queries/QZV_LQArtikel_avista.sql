SELECT
  [Artikel-all_avista_out].*,
  [LQ Art-Nr] & "-" & [Variante] AS ArtNrVar
FROM
  [Artikel-all_avista_out]
WHERE
  (
    (
      ([Artikel-all_avista_out].QZV) = True
    )
  );
