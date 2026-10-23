SELECT
  Mastertab.ID,
  Mastertab.[Baugruppen-ID],
  Mastertab.Unterbaugruppe,
  Mastertab.[Artikel-ID 1],
  Mastertab.[LINK-Nr - UBG],
  Mastertab.[Herstellungsdatum - BG],
  Mastertab.[Herstellernummer - BG],
  Mastertab.[Serial-Nr - BG],
  Mastertab.[Herstellernummer - UBG],
  Mastertab.[Herstellungsdatum - UBG],
  Mastertab.[Serial-Nr - UBG],
  Mastertab.Seriennummer,
  Mastertab.[Menge betroffen - Kunde],
  Mastertab.[Menge betroffen - LQ],
  Mastertab.[KW/JJ],
  Mastertab.[Kategorie-Status],
  [Artikel-all_avista_out].AvistaArtikelTyp,
  [Artikel-all_avista_out].Variante,
  [Baugruppen-avista_out].[Var Nr],
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.Auftragsdatum,
  ReklMemo.VorgangstypID,
  Mastertab.intern,
  Mastertab.extern
FROM
  (
    (
      Mastertab
      LEFT JOIN [Artikel-all_avista_out] ON Mastertab.[Artikel-ID 1] = [Artikel-all_avista_out].[Artikel-ID]
    )
    LEFT JOIN [Baugruppen-avista_out] ON Mastertab.[Baugruppen-ID] = [Baugruppen-avista_out].[Baugruppen-ID]
  )
  LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID;
