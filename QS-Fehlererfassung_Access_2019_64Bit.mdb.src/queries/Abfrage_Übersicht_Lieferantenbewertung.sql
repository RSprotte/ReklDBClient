SELECT
  Lieferanten.Lieferantenname,
  Lieferanten.[L-Nummer/AVISTA] AS [L-Nr avista],
  Count(Mastertab.ID) AS Reklamationen,
  Sum(
    [Menge betroffen - Kunde] + [Menge betroffen - LQ]
  ) AS [Teile reklamiert],
  [Artikel-all_avista_out].[LQ Art-Nr] AS Artikel,
  [Baugruppen-avista_out].[LQ Art-Nr] AS Baugruppe,
  Format$(
    [Mastertab].[Auftragsdatum], "mmm"". ""yyyy"
  ) AS Monat,
  Mastertab.Rückmeldung_Lieferant_Soll AS [Antwort-Soll],
  Mastertab.Rückmeldung_Lieferant_Ist AS [Antwort-Ist],
  IIf(
    [Rückmeldung_Lieferant_Ist] Is Null,
    DateDiff(
      "y",
      (
        Now()
      ),
      [Rückmeldung_Lieferant_Soll]
    ),
    DateDiff(
      "y", [Rückmeldung_Lieferant_Ist],
      [Rückmeldung_Lieferant_Soll]
    )
  ) AS Differenz,
  Mastertab.ID,
  Lieferanten.[L-Kategorie],
  [Zeitbedarf] * 45 + [Material-Kosten] + [Fracht-Kosten] + [Prüfkosten] AS [LQ Aufwand]
FROM
  Lieferanten
  INNER JOIN (
    [Kategorien-Problemstatus]
    RIGHT JOIN (
      [Baugruppen-avista_out]
      RIGHT JOIN (
        [Artikel-all_avista_out]
        RIGHT JOIN Mastertab ON [Artikel-all_avista_out].[Artikel-ID] = Mastertab.[Artikel-ID 1]
      ) ON [Baugruppen-avista_out].[Baugruppen-ID] = Mastertab.[Baugruppen-ID]
    ) ON [Kategorien-Problemstatus].[Kategorie-Problemstatus] = Mastertab.[Kategorie-Problem]
  ) ON Lieferanten.[Lieferant-ID] = Mastertab.[Lieferanten-Code 1]
WHERE
  (
    (
      (Mastertab.Auftragsdatum) > #12/31/2009#
    )
    AND (
      (Mastertab.Verursacher) = "Lieferant"
    )
  )
GROUP BY
  Lieferanten.Lieferantenname,
  Lieferanten.[L-Nummer/AVISTA],
  [Artikel-all_avista_out].[LQ Art-Nr],
  [Baugruppen-avista_out].[LQ Art-Nr],
  Format$(
    [Mastertab].[Auftragsdatum], "mmm"". ""yyyy"
  ),
  Mastertab.Rückmeldung_Lieferant_Soll,
  Mastertab.Rückmeldung_Lieferant_Ist,
  IIf(
    [Rückmeldung_Lieferant_Ist] Is Null,
    DateDiff(
      "y",
      (
        Now()
      ),
      [Rückmeldung_Lieferant_Soll]
    ),
    DateDiff(
      "y", [Rückmeldung_Lieferant_Ist],
      [Rückmeldung_Lieferant_Soll]
    )
  ),
  Mastertab.ID,
  Lieferanten.[L-Kategorie],
  [Zeitbedarf] * 45 + [Material-Kosten] + [Fracht-Kosten] + [Prüfkosten],
  Lieferanten.[Anzahl gelieferter Teile],
  [Kategorien-Problemstatus].[Kategorie-Problemstatus]
HAVING
  (
    (
      (
        [Kategorien-Problemstatus].[Kategorie-Problemstatus]
      ) = "berechtigt"
    )
  )
ORDER BY
  Lieferanten.Lieferantenname,
  Count(Mastertab.ID) DESC;
