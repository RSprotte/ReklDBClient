SELECT
  LieferantAnsprechpartner.*,
  [Name] & ", " & [Vorname] AS GName
FROM
  LieferantAnsprechpartner;
