SELECT
  RSP_Ausw_Basisdaten_1.ID,
  Format([Auftragsdatum], "yyyy") AS [A-DatumJahr],
  RSP_Ausw_Basisdaten_1.[Kategorie-Problem],
  RSP_Ausw_Basisdaten_1.extern,
  RSP_Ausw_Basisdaten_1.[Fracht-Kosten],
  RSP_Ausw_Basisdaten_1.[Material-Kosten],
  RSP_Ausw_Basisdaten_1.Zeitbedarf,
  [Zeitbedarf] * 45 + [Fracht-Kosten] + [Material-Kosten] AS GesamtKosten,
  RSP_Ausw_Basisdaten_1.[TRUMPF?]
FROM
  RSP_Ausw_Basisdaten_1
WHERE
  (
    (
      (
        RSP_Ausw_Basisdaten_1.[Kategorie-Problem]
      ) = "berechtigt"
      Or (
        RSP_Ausw_Basisdaten_1.[Kategorie-Problem]
      ) = "nicht berechtigt"
    )
    AND (
      (RSP_Ausw_Basisdaten_1.extern) = True
    )
  )
  OR (
    (
      (
        RSP_Ausw_Basisdaten_1.[Kategorie-Problem]
      ) = "berechtigt"
      Or (
        RSP_Ausw_Basisdaten_1.[Kategorie-Problem]
      ) = "nicht berechtigt"
    )
    AND (
      (
        RSP_Ausw_Basisdaten_1.[TRUMPF?]
      ) = True
    )
  );
