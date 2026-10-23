SELECT
  Format$(
    [Mastertab].[Auftragsdatum], 'yyyy'
  ) AS Jahr,
  Count(Mastertab.ID) AS AnzahlvonID,
  Sum(Mastertab.Zeitbedarf) AS SummevonZeitbedarf,
  Sum(Mastertab.[PE-Bearbeitung]) AS [SummevonPE-Bearbeitung],
  Sum(Mastertab.[QS-Bearbeitung]) AS [SummevonQS-Bearbeitung],
  Sum(
    [Zeitbedarf] + [QS-Bearbeitung] + [PE-Bearbeitung]
  ) AS [Gesamt-Zeitbedarf],
  Sum(
    (
      [Zeitbedarf] + [QS-Bearbeitung] + [PE-Bearbeitung]
    ) * 45
  ) AS Kosten,
  Taetigkeiten.Taetigkeit,
  Mastertab.intern,
  Mastertab.extern
FROM
  Mastertab
  LEFT JOIN Taetigkeiten ON Mastertab.TaetigkeitID = Taetigkeiten.ID
WHERE
  (
    (
      (Mastertab.[Kategorie-Problem]) = "berechtigt"
    )
    AND (
      (Mastertab.[Kategorie-Status]) = "Abgeschlossen"
    )
    AND (
      (Mastertab.Auftragsdatum) >= #1/1/2010#
    )
  )
GROUP BY
  Format$(
    [Mastertab].[Auftragsdatum], 'yyyy'
  ),
  Taetigkeiten.Taetigkeit,
  Mastertab.intern,
  Mastertab.extern
HAVING
  (
    (
      (
        Format$(
          [Mastertab].[Auftragsdatum], 'yyyy'
        )
      ) = [Bitte geben Sie das gewünschte Jahr an]
    )
  );
