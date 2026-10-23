SELECT
  Maßnahmen.MaßnahmeID AS VorgangID,
  Sachbearbeiter.Kurzzeichen
FROM
  Maßnahmen
  INNER JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID];
