SELECT
  RSP_Ausw_Basisdaten_1.ID,
  Format([Auftragsdatum], "yyyy") AS ADatum,
  RSP_Ausw_Basisdaten_1.[TRUMPF?],
  RSP_Ausw_Basisdaten_1.extern,
  RSP_Ausw_Basisdaten_1.[Kategorie-Problem],
  RSP_Ausw_Basisdaten_1.[WE-Datum],
  RSP_Ausw_Basisdaten_1.[Verursacher Lieferant]
FROM
  RSP_Ausw_Basisdaten_1
WHERE
  (
    (
      (
        Format([Auftragsdatum], "yyyy")
      ) = "2012"
    )
    AND (
      (
        RSP_Ausw_Basisdaten_1.[TRUMPF?]
      ) = True
    )
    AND (
      (
        RSP_Ausw_Basisdaten_1.[Kategorie-Problem]
      ) = "berechtigt"
      Or (
        RSP_Ausw_Basisdaten_1.[Kategorie-Problem]
      ) = "nicht berechtigt"
    )
    AND (
      (
        RSP_Ausw_Basisdaten_1.[WE-Datum]
      ) Is Not Null
    )
  )
  OR (
    (
      (
        Format([Auftragsdatum], "yyyy")
      ) = "2012"
    )
    AND (
      (RSP_Ausw_Basisdaten_1.extern) = True
    )
    AND (
      (
        RSP_Ausw_Basisdaten_1.[Kategorie-Problem]
      ) = "berechtigt"
      Or (
        RSP_Ausw_Basisdaten_1.[Kategorie-Problem]
      ) = "nicht berechtigt"
    )
    AND (
      (
        RSP_Ausw_Basisdaten_1.[WE-Datum]
      ) Is Not Null
    )
  );
