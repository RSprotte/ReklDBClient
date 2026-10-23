SELECT
  LieferantUndQSVMahnwesen.[Mahnung-ID],
  Mahnstufen.Mahnstufe,
  Mahnstufen.Anzeige
FROM
  LieferantUndQSVMahnwesen
  INNER JOIN Mahnstufen ON LieferantUndQSVMahnwesen.[Mahnustufe-ID] = Mahnstufen.[MS-ID];
