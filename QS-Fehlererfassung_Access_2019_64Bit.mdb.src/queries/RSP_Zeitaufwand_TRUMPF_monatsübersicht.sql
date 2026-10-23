SELECT
  RSP_Zeitauwand_TRUMPF.LQEinheit,
  RSP_Zeitauwand_TRUMPF.LQBereich,
  RSP_Zeitauwand_TRUMPF.LQWorkstation,
  Format([Auftragsdatum], "mm\.yy") AS Datum,
  RSP_Zeitauwand_TRUMPF.[SummevonAufwand-TRUMPF],
  Sum(
    RSP_Zeitauwand_TRUMPF.AnzahlvonID
  ) AS [Anzahl Fehler],
  RSP_Zeitauwand_TRUMPF.[Kategorie-Fehleranalyse],
  RSP_Zeitauwand_TRUMPF.Firma
FROM
  RSP_Zeitauwand_TRUMPF
GROUP BY
  RSP_Zeitauwand_TRUMPF.LQEinheit,
  RSP_Zeitauwand_TRUMPF.LQBereich,
  RSP_Zeitauwand_TRUMPF.LQWorkstation,
  Format([Auftragsdatum], "mm\.yy"),
  RSP_Zeitauwand_TRUMPF.[SummevonAufwand-TRUMPF],
  RSP_Zeitauwand_TRUMPF.[Kategorie-Fehleranalyse],
  RSP_Zeitauwand_TRUMPF.Firma,
  RSP_Zeitauwand_TRUMPF.Auftragsdatum
HAVING
  (
    (
      (
        RSP_Zeitauwand_TRUMPF.Auftragsdatum
      ) >= #1/1/2013#
    )
    AND (
      (
        RSP_Zeitauwand_TRUMPF.[SummevonAufwand-TRUMPF]
      ) > 0
    )
  )
ORDER BY
  RSP_Zeitauwand_TRUMPF.Auftragsdatum DESC,
  RSP_Zeitauwand_TRUMPF.[SummevonAufwand-TRUMPF] DESC;
