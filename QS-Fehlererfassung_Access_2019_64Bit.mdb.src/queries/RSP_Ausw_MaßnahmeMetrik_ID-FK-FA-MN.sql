SELECT
  RSP_Ausw_Basisdaten_1.ID,
  [RSP_Ausw_MaßnahmeMetrik_FK-FA-MN].*,
  RSP_Ausw_Basisdaten_1.BG,
  RSP_Ausw_Basisdaten_1.[LINK-Nr - UBG],
  RSP_Ausw_Basisdaten_1.ArtNr,
  RSP_Ausw_Basisdaten_1.LQEinheit,
  RSP_Ausw_Basisdaten_1.[Menge betroffen - Kunde],
  RSP_Ausw_Basisdaten_1.Fehlerpositionen
FROM
  RSP_Ausw_Basisdaten_1,
  [RSP_Ausw_MaßnahmeMetrik_FK-FA-MN]
WHERE
  (
    (
      (
        RSP_Ausw_Basisdaten_1.LQEinheit
      ) Is Not Null
    )
  );
