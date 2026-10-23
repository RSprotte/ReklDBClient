SELECT
  tblKostenaufstellung.ReklID,
  tblKostenaufstellung.Pos,
  tblKostenaufstellung.Datum,
  tblKostenTyp.[Anzeige] AS Typ,
  tblKostenaufstellung.[InfoKostenTyp] AS Info,
  tblKostenaufstellung.Kosten
FROM
  tblKostenaufstellung
  LEFT JOIN tblKostenTyp ON tblKostenaufstellung.KostentypID = tblKostenTyp.KostenTypID
WHERE
  (
    (
      (
        tblKostenaufstellung.KostentypID
      ) = 14
      Or (
        tblKostenaufstellung.KostentypID
      ) = 5
    )
  )
GROUP BY
  tblKostenaufstellung.ReklID,
  tblKostenaufstellung.Pos,
  tblKostenaufstellung.Datum,
  tblKostenTyp.[Anzeige],
  tblKostenaufstellung.[InfoKostenTyp],
  tblKostenaufstellung.Kosten;
