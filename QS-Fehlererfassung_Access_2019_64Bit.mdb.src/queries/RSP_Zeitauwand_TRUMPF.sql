SELECT
  Count(Mastertab.ID) AS AnzahlvonID,
  LQEinheit.LQEinheit,
  LQBereich.LQBereich,
  LQWorkstation.LQWorkstation,
  Sum(Mastertab.[Aufwand-TRUMPF]) AS [SummevonAufwand-TRUMPF],
  Mastertab.[Kategorie-Fehleranalyse],
  Mastertab.Auftragsdatum,
  tblFirma.Firma_Name AS Firma
FROM
  (
    (
      (
        (
          Mastertab
          LEFT JOIN LQEinheit ON Mastertab.LQEinheitID = LQEinheit.ID
        )
        LEFT JOIN LQBereich ON Mastertab.LQBereichID = LQBereich.ID
      )
      LEFT JOIN LQWorkstation ON Mastertab.LQWorkstationErfasserID = LQWorkstation.ID
    )
    INNER JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]
  )
  LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
GROUP BY
  LQEinheit.LQEinheit,
  LQBereich.LQBereich,
  LQWorkstation.LQWorkstation,
  Mastertab.[Kategorie-Fehleranalyse],
  Mastertab.Auftragsdatum,
  tblFirma.Firma_Name
HAVING
  (
    (
      (LQEinheit.LQEinheit) Is Not Null
    )
  )
ORDER BY
  LQEinheit.LQEinheit;
