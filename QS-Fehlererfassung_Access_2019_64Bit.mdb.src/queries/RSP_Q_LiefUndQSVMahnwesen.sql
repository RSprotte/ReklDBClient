SELECT
  LieferantUndQSVMahnwesen.[Mahnung-ID],
  LieferantUndQSVMahnwesen.[R-ID],
  Mahnstufen.Mahnstufe,
  LieferantUndQSVMahnwesen.Datum,
  LieferantUndQSVMahnwesen.Typ
FROM
  LieferantUndQSVMahnwesen
  INNER JOIN Mahnstufen ON LieferantUndQSVMahnwesen.[Mahnustufe-ID] = Mahnstufen.[MS-ID];
