SELECT
  qryServiceEinsatzKosten.ReklID,
  qryServiceEinsatzKosten.Datum,
  qryServiceEinsatzKosten.Typ,
  qryServiceEinsatzKosten.Info,
  qryServiceEinsatzKosten.Kosten
FROM
  qryServiceEinsatzKosten
ORDER BY
  qryServiceEinsatzKosten.ReklID DESC;
