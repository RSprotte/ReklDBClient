PARAMETERS [Bitte geben Sie das Jahr an] Text (255);
SELECT
  Mastertab.Auftragsdatum,
  Mastertab.Zeitbedarf,
  Mastertab.[Material-Kosten],
  Mastertab.[Fracht-Kosten],
  Mastertab.Prüfkosten,
  Mastertab.[Belastungsbetrag Kunde],
  [Zeitbedarf] * 45 + [Material-Kosten] + [Fracht-Kosten] + [Prüfkosten] + [Belastungsbetrag Kunde] AS [Kosten-Aufwand],
  Mastertab.Reparaturbetrag,
  Mastertab.[Belastungsbetrag LQ],
  Mastertab.ID,
  Format$(Mastertab.Auftragsdatum, 'yyyy') AS Jahr
FROM
  Mastertab
WHERE
  (
    (
      (Mastertab.Auftragsdatum) > #12/31/2007#
    )
    And (
      (
        Format$(Mastertab.Auftragsdatum, 'yyyy')
      ) = [Bitte geben Sie das Jahr an]
    )
  )
ORDER BY
  Mastertab.Auftragsdatum;
