SELECT
  Maßnahmen.MaßnahmeID,
  Sachbearbeiter.Kurzzeichen,
  Maßnahmen.Maßnahme,
  Maßnahmen.Beschreibung,
  Maßnahmen.[Ziel der Maßnahme],
  Maßnahmen.[Ergebnis Maßnahme],
  Maßnahmen.StartUmsetzung,
  Maßnahmen.EndeUmsetzung,
  Maßnahmen.StatusUmsetzung,
  Maßnahmen.ArtDerMaßnahme,
  Maßnahmen.LinkDokument,
  Maßnahmen.Verantwortlich,
  Maßnahmen.Fortschritt,
  Maßnahmen.Ursprung,
  Maßnahmen.BearbeitungsOrt
FROM
  Maßnahmen
  LEFT JOIN Sachbearbeiter ON Maßnahmen.Verantwortlich = Sachbearbeiter.[SB-ID]
ORDER BY
  Maßnahmen.MaßnahmeID;
