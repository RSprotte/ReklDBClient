UPDATE
  AufwandVorgangPositionen
SET
  AufwandVorgangPositionen.Aufwand = [Zeit] * 45,
  AufwandVorgangPositionen.StdSatzID = 1,
  AufwandVorgangPositionen.DatumStdSatz = Date(),
  AufwandVorgangPositionen.StdSatz = 45
WHERE
  (
    (
      (
        AufwandVorgangPositionen.StdSatzID
      ) = 0
    )
  );
