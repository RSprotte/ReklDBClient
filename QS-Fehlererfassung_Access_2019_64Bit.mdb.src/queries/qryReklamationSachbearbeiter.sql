SELECT
  Mastertab.ID,
  Sachbearbeiter.Kurzzeichen,
  Sachbearbeiter.EmailName
FROM
  Mastertab
  INNER JOIN Sachbearbeiter ON Mastertab.Sachbearbeiter = Sachbearbeiter.Name;
