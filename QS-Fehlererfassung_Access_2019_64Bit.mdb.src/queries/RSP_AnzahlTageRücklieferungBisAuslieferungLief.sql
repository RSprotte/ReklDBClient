SELECT
  RSP_Ausw_Basisdaten_1.ID,
  Format(
    [RSP_Ausw_Basisdaten_1].[Auftragsdatum],
    "yyyy\.mm"
  ) AS ADATUM,
  RSP_Ausw_Basisdaten_1.[WE-Datum],
  RSP_Ausw_Basisdaten_1.[Vorgang - Pos20] AS Ausdr1,
  RSP_Ausw_Basisdaten_1.[Vorgang - Pos30] AS Ausdr2,
  RSP_Ausw_Basisdaten_1.[Vorgang - Pos40] AS Ausdr3,
  RSP_Ausw_Basisdaten_1.[Vorgang - Pos50] AS Ausdr4,
  RSP_Ausw_Basisdaten_1.[Vorgang - Pos60] AS Ausdr5,
  RSP_Ausw_Basisdaten_1.[Vorgang - Pos70] AS Ausdr6,
  RSP_Ausw_Basisdaten_1.[Vorgang - Pos80] AS Ausdr7,
  Mastertab.[Datum - Pos20] AS Ausdr8,
  Mastertab.[Datum - Pos30] AS Ausdr9,
  Mastertab.[Datum - Pos40] AS Ausdr10,
  Mastertab.[Datum - Pos50] AS Ausdr11,
  Mastertab.[Datum - Pos60] AS Ausdr12,
  Mastertab.[Datum - Pos70] AS Ausdr13,
  Mastertab.[Datum - Pos80] AS Ausdr14,
  getTageQuery([RSP_Ausw_Basisdaten_1].[ID]) AS TageRücklieferung,
  RSP_Ausw_Basisdaten_1.Sachbearbeiter
FROM
  Mastertab
  INNER JOIN RSP_Ausw_Basisdaten_1 ON Mastertab.ID = RSP_Ausw_Basisdaten_1.ID
WHERE
  (
    (
      (
        RSP_Ausw_Basisdaten_1.[WE-Datum]
      ) Is Not Null
    )
    AND (
      (
        [RSP_Ausw_Basisdaten_1].[Vorgang - Pos20]
      ) = "26"
    )
    AND (
      (
        RSP_Ausw_Basisdaten_1.[TRUMPF?]
      ) = True
    )
  )
  OR (
    (
      (
        RSP_Ausw_Basisdaten_1.[WE-Datum]
      ) Is Not Null
    )
    AND (
      (
        RSP_Ausw_Basisdaten_1.[TRUMPF?]
      ) = True
    )
    AND (
      (
        [RSP_Ausw_Basisdaten_1].[Vorgang - Pos30]
      ) = "26"
    )
  )
  OR (
    (
      (
        RSP_Ausw_Basisdaten_1.[WE-Datum]
      ) Is Not Null
    )
    AND (
      (
        RSP_Ausw_Basisdaten_1.[TRUMPF?]
      ) = True
    )
    AND (
      (
        [RSP_Ausw_Basisdaten_1].[Vorgang - Pos40]
      ) = "26"
    )
  )
  OR (
    (
      (
        RSP_Ausw_Basisdaten_1.[WE-Datum]
      ) Is Not Null
    )
    AND (
      (
        RSP_Ausw_Basisdaten_1.[TRUMPF?]
      ) = True
    )
    AND (
      (
        [RSP_Ausw_Basisdaten_1].[Vorgang - Pos50]
      ) = "26"
    )
  )
  OR (
    (
      (
        RSP_Ausw_Basisdaten_1.[WE-Datum]
      ) Is Not Null
    )
    AND (
      (
        RSP_Ausw_Basisdaten_1.[TRUMPF?]
      ) = True
    )
    AND (
      (
        [RSP_Ausw_Basisdaten_1].[Vorgang - Pos60]
      ) = "26"
    )
  )
  OR (
    (
      (
        RSP_Ausw_Basisdaten_1.[WE-Datum]
      ) Is Not Null
    )
    AND (
      (
        RSP_Ausw_Basisdaten_1.[TRUMPF?]
      ) = True
    )
    AND (
      (
        [RSP_Ausw_Basisdaten_1].[Vorgang - Pos70]
      ) = "26"
    )
  )
  OR (
    (
      (
        RSP_Ausw_Basisdaten_1.[WE-Datum]
      ) Is Not Null
    )
    AND (
      (
        RSP_Ausw_Basisdaten_1.[TRUMPF?]
      ) = True
    )
    AND (
      (
        [RSP_Ausw_Basisdaten_1].[Vorgang - Pos80]
      ) = "26"
    )
  )
  OR (
    (
      (
        RSP_Ausw_Basisdaten_1.[WE-Datum]
      ) Is Not Null
    )
    AND (
      (
        [RSP_Ausw_Basisdaten_1].[Vorgang - Pos20]
      ) = "26"
    )
    AND (
      (RSP_Ausw_Basisdaten_1.extern) = True
    )
  )
  OR (
    (
      (
        RSP_Ausw_Basisdaten_1.[WE-Datum]
      ) Is Not Null
    )
    AND (
      (
        [RSP_Ausw_Basisdaten_1].[Vorgang - Pos30]
      ) = "26"
    )
    AND (
      (RSP_Ausw_Basisdaten_1.extern) = True
    )
  )
  OR (
    (
      (
        RSP_Ausw_Basisdaten_1.[WE-Datum]
      ) Is Not Null
    )
    AND (
      (
        [RSP_Ausw_Basisdaten_1].[Vorgang - Pos40]
      ) = "26"
    )
    AND (
      (RSP_Ausw_Basisdaten_1.extern) = True
    )
  )
  OR (
    (
      (
        RSP_Ausw_Basisdaten_1.[WE-Datum]
      ) Is Not Null
    )
    AND (
      (
        [RSP_Ausw_Basisdaten_1].[Vorgang - Pos50]
      ) = "26"
    )
    AND (
      (RSP_Ausw_Basisdaten_1.extern) = True
    )
  )
  OR (
    (
      (
        RSP_Ausw_Basisdaten_1.[WE-Datum]
      ) Is Not Null
    )
    AND (
      (
        [RSP_Ausw_Basisdaten_1].[Vorgang - Pos60]
      ) = "26"
    )
    AND (
      (RSP_Ausw_Basisdaten_1.extern) = True
    )
  )
  OR (
    (
      (
        RSP_Ausw_Basisdaten_1.[WE-Datum]
      ) Is Not Null
    )
    AND (
      (
        [RSP_Ausw_Basisdaten_1].[Vorgang - Pos70]
      ) = "26"
    )
    AND (
      (RSP_Ausw_Basisdaten_1.extern) = True
    )
  )
  OR (
    (
      (
        RSP_Ausw_Basisdaten_1.[WE-Datum]
      ) Is Not Null
    )
    AND (
      (
        [RSP_Ausw_Basisdaten_1].[Vorgang - Pos80]
      ) = "26"
    )
    AND (
      (RSP_Ausw_Basisdaten_1.extern) = True
    )
  )
ORDER BY
  RSP_Ausw_Basisdaten_1.ID DESC;
