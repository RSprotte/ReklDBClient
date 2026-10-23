SELECT
  tblKostenaufstellung.ReklID,
  tblKostenaufstellung.Datum,
  tblKostenTyp.Anzeige AS Typ,
  tblKostenaufstellung.InfoKostenTyp AS Info,
  tblKostenaufstellung.Kosten
FROM
  Mastertab
  RIGHT JOIN (
    tblKostenaufstellung
    LEFT JOIN tblKostenTyp ON tblKostenaufstellung.KostentypID = tblKostenTyp.KostenTypID
  ) ON Mastertab.ID = tblKostenaufstellung.ReklID
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
    AND (
      (Mastertab.[Einsatz vor Ort]) = True
    )
  )
GROUP BY
  tblKostenaufstellung.ReklID,
  tblKostenaufstellung.Datum,
  tblKostenTyp.Anzeige,
  tblKostenaufstellung.InfoKostenTyp,
  tblKostenaufstellung.Kosten,
  tblKostenaufstellung.Pos
UNION
SELECT
  AufwandVorgangPositionen.MID AS ReklID,
  AufwandVorgangPositionen.Datum,
  [Vorgang-Text].[Vorgang-Beschreibung] AS Typ,
  AufwandVorgangPositionen.InfoPos AS Info,
  AufwandVorgangPositionen.Aufwand AS Kosten
FROM
  (
    AufwandVorgangPositionen
    LEFT JOIN [Vorgang-Text] ON AufwandVorgangPositionen.Vorgang = [Vorgang-Text].[Vorgang-ID]
  )
  LEFT JOIN Mastertab ON AufwandVorgangPositionen.MID = Mastertab.ID
WHERE
  (
    (
      (
        AufwandVorgangPositionen.Vorgang
      ) = 21
      Or (
        AufwandVorgangPositionen.Vorgang
      ) = 141
      Or (
        AufwandVorgangPositionen.Vorgang
      ) = 142
    )
    AND (
      (Mastertab.[Einsatz vor Ort]) = True
    )
  )
GROUP BY
  AufwandVorgangPositionen.MID,
  AufwandVorgangPositionen.Datum,
  [Vorgang-Text].[Vorgang-Beschreibung],
  AufwandVorgangPositionen.InfoPos,
  AufwandVorgangPositionen.Aufwand,
  AufwandVorgangPositionen.Pos;
