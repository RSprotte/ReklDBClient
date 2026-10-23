SELECT
  Mastertab.ID,
  ReklMemo.AktuellerStundensatz,
  Mastertab.Zeitbedarf,
  Mastertab.[Material-Kosten],
  Mastertab.[Kategorie-Status],
  Mastertab.[Fracht-Kosten],
  Mastertab.Prüfkosten,
  Mastertab.[Belastungsbetrag Kunde],
  Mastertab.[Belastungsbetrag LQ],
  Mastertab.Reparaturbetrag,
  Mastertab.[Kategorie-Status],
  Mastertab.[Meldungsnummer Kunde],
  ReklMemo.VorgangstypID,
  Mastertab.Auftragsdatum,
  Mastertab.intern,
  Mastertab.extern,
  Mastertab.[LINK-Nr - UBG],
  [Artikel-all_avista_out].Variante,
  Mastertab.[Artikel-ID 1],
  Mastertab.[Baugruppen-ID]
FROM
  (
    (
      Mastertab
      LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID
    )
    LEFT JOIN [Artikel-all_avista_out] ON Mastertab.[Artikel-ID 1] = [Artikel-all_avista_out].[Artikel-ID]
  )
  LEFT JOIN [Baugruppen-avista_out] ON Mastertab.[Baugruppen-ID] = [Baugruppen-avista_out].[Baugruppen-ID];
