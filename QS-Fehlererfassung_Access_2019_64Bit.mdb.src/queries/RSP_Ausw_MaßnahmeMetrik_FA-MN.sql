SELECT
  MaßnahmenFehlerarten.Fehlerart AS Ausdr1,
  Maßnahmen.*
FROM
  MaßnahmenFehlerarten
  LEFT JOIN Maßnahmen ON MaßnahmenFehlerarten.MaßnahmeID = Maßnahmen.MaßnahmeID;
