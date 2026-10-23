SELECT
  tblKostenaufstellung.ReklID,
  Sum(tblKostenaufstellung.Kosten) AS SummevonKosten
FROM
  tblKostenaufstellung
  LEFT JOIN tblKostenTyp ON tblKostenaufstellung.KostentypID = tblKostenTyp.KostenTypID
WHERE
  (
    (
      (tblKostenTyp.Ausgaben) = True
    )
  )
GROUP BY
  tblKostenaufstellung.ReklID;
