SELECT
  [Baugruppen-avista_out].[Baugruppen-ID],
  [Baugruppen-avista_out].[LQ Art-Nr],
  [Baugruppen-avista_out].[Var Nr],
  [Baugruppen-avista_out].[Kunden-Nr],
  [Baugruppen-avista_out].[Bezeichnung 1],
  [Baugruppen-avista_out].Produktgruppe,
  [LQ Art-Nr] & "-" & [Var Nr] AS ArtNr
FROM
  [Baugruppen-avista_out]
WHERE
  (
    (
      ([Baugruppen-avista_out].QZV) = True
    )
  );
