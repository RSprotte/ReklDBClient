SELECT
  [Baugruppen-avista_out].*,
  [LQ Art-Nr] & "-" & [Var Nr] AS ArtNrVar
FROM
  [Baugruppen-avista_out]
WHERE
  (
    (
      ([Baugruppen-avista_out].QZV) = True
    )
  );
