SELECT
  Format$(Mastertab.Auftragsdatum, 'yyyy') AS Jahr,
  Sum(Mastertab.Zeitbedarf) AS [Zeitaufwand LQ],
  Sum(Mastertab.[Material-Kosten]) AS [Materialkosten LQ],
  Sum(Mastertab.[Fracht-Kosten]) AS [Frachtkosten LQ],
  Sum(Mastertab.Prüfkosten) AS [Prüfkosten LQ],
  Sum(
    Mastertab.[Belastungsbetrag Kunde]
  ) AS [Belastungen an LQ],
  Sum(
    [Zeitbedarf] * 45 + [Material-Kosten] + [Fracht-Kosten] + [Prüfkosten] + [Belastungsbetrag Kunde]
  ) AS [Kosten-Aufwand],
  Sum(Mastertab.Reparaturbetrag) AS Reparaturkosten,
  Sum(
    Mastertab.[Belastungsbetrag LQ]
  ) AS [Belastungen von LQ],
  Count(Mastertab.ID) AS [QS-Vorgänge]
FROM
  Mastertab
WHERE
  (
    (
      (Mastertab.Auftragsdatum) > #12/31/2007#
    )
  )
GROUP BY
  Format$(Mastertab.Auftragsdatum, 'yyyy');
