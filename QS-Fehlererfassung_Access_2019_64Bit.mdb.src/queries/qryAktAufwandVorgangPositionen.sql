UPDATE
  AufwandVorgangPositionen
  INNER JOIN tblAufwandStdSaetze ON AufwandVorgangPositionen.StdSatzID = tblAufwandStdSaetze.StdSatzID
SET
  AufwandVorgangPositionen.DatumStdSatz = [Datum],
  AufwandVorgangPositionen.StdSatzID = 1,
  AufwandVorgangPositionen.StdSatz = [Betrag];
