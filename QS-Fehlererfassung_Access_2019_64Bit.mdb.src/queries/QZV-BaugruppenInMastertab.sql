SELECT
  RSP_Ausw_Basisdaten_1.ID,
  RSP_Ausw_Basisdaten_1.Datum,
  RSP_Ausw_Basisdaten_1.Auftragsdatum,
  RSP_Ausw_Basisdaten_1.BG,
  RSP_Ausw_Basisdaten_1.ArtNr,
  RSP_Ausw_Basisdaten_1.[LINK-Nr - UBG],
  RSP_Ausw_Basisdaten_1.[Menge betroffen - Kunde],
  RSP_Ausw_Basisdaten_1.[Menge betroffen - LQ],
  RSP_Ausw_Basisdaten_1.Fehlerpositionen
FROM
  RSP_Ausw_Basisdaten_1
  INNER JOIN [QZV-Baugruppen] ON RSP_Ausw_Basisdaten_1.BG = [QZV-Baugruppen].ArtNr;
