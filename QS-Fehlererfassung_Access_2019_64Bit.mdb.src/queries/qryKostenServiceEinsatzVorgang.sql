SELECT
  AufwandVorgangPositionen.[MID] AS ReklID,
  AufwandVorgangPositionen.Pos,
  AufwandVorgangPositionen.Datum,
  [Vorgang-Text].[Vorgang-Beschreibung] AS Typ,
  AufwandVorgangPositionen.[InfoPos] AS Info,
  AufwandVorgangPositionen.[Aufwand] AS Kosten
FROM
  AufwandVorgangPositionen
  LEFT JOIN [Vorgang-Text] ON AufwandVorgangPositionen.Vorgang = [Vorgang-Text].[Vorgang-ID]
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
  )
GROUP BY
  AufwandVorgangPositionen.[MID],
  AufwandVorgangPositionen.Pos,
  AufwandVorgangPositionen.Datum,
  [Vorgang-Text].[Vorgang-Beschreibung],
  AufwandVorgangPositionen.[InfoPos],
  AufwandVorgangPositionen.[Aufwand];
