SELECT
  LieferantAnsprechpartner.AnsprechpartnerID,
  LieferantAnsprechpartner.Name & ", " & LieferantAnsprechpartner.Vorname AS NameAP,
  LieferantAnsprechpartner.Email,
  Lieferanten.Lieferantenname,
  Sachbearbeiter.Name,
  Sachbearbeiter.EmailName
FROM
  (
    LieferantAnsprechpartner
    LEFT JOIN Lieferanten ON LieferantAnsprechpartner.LieferantID = Lieferanten.[Lieferant-ID]
  )
  LEFT JOIN Sachbearbeiter ON Lieferanten.VerantwortlichEinkauf = Sachbearbeiter.Kurzzeichen;
