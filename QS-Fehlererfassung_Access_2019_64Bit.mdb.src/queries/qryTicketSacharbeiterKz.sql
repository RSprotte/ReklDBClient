SELECT
  Maßnahmen.MaßnahmeID AS VorgangID,
  Sachbearbeiter.Kurzzeichen
FROM
  Maßnahmen
  INNER JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]
UNION
SELECT
  Mastertab.ID,
  Sachbearbeiter.Kurzzeichen
FROM
  Mastertab
  INNER JOIN Sachbearbeiter ON Mastertab.Sachbearbeiter = Sachbearbeiter.Name;
