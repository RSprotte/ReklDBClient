SELECT
  RSP_Ausw_Basisdaten_1.ID,
  Format(
    RSP_Ausw_Basisdaten_1.Auftragsdatum,
    "yyyy\.mm"
  ) AS ADATUM,
  RSP_Ausw_Basisdaten_1.[WE-Datum],
  AufwandVorgangPositionen.Vorgang,
  getTageQuery(RSP_Ausw_Basisdaten_1.ID) AS TageRücklieferung,
  RSP_Ausw_Basisdaten_1.Sachbearbeiter
FROM
  (
    Mastertab
    INNER JOIN RSP_Ausw_Basisdaten_1 ON Mastertab.ID = RSP_Ausw_Basisdaten_1.ID
  )
  INNER JOIN AufwandVorgangPositionen ON Mastertab.ID = AufwandVorgangPositionen.MID
WHERE
  (
    (
      (
        RSP_Ausw_Basisdaten_1.[WE-Datum]
      ) Is Not Null
    )
    AND (
      (
        AufwandVorgangPositionen.Vorgang
      ) = 26
    )
    AND (
      (RSP_Ausw_Basisdaten_1.extern) = True
    )
  )
ORDER BY
  RSP_Ausw_Basisdaten_1.ID DESC;
