SELECT
  FehlerartKlassifizierung.Klassifizierung AS Ausdr2,
  [RSP_Ausw_MaßnahmeMetrik_FA-MN].*
FROM
  FehlerartKlassifizierung,
  [RSP_Ausw_MaßnahmeMetrik_FA-MN];
