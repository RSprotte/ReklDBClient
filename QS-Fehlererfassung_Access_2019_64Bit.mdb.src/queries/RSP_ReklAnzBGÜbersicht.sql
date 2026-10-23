SELECT
  [Baugruppen-avista_out].[LQ Art-Nr],
  [Baugruppen-avista_out].[Kunden-Nr],
  [Baugruppen-avista_out].[Bezeichnung 1],
  RSP_ReklAnzBGAktGJ.[Anzahl Reklamationen] AS 12M,
  RSP_ReklAnzBGAktGJ_M7_12.[Anzahl Reklamationen] AS 6M,
  RSP_ReklAnzBGAktGJ_M10_12.[Anzahl Reklamationen] AS 3M
FROM
  (
    (
      [Baugruppen-avista_out]
      LEFT JOIN RSP_ReklAnzBGAktGJ ON [Baugruppen-avista_out].[Baugruppen-ID] = RSP_ReklAnzBGAktGJ.[Baugruppen-ID]
    )
    LEFT JOIN RSP_ReklAnzBGAktGJ_M10_12 ON [Baugruppen-avista_out].[Baugruppen-ID] = RSP_ReklAnzBGAktGJ_M10_12.[Baugruppen-ID]
  )
  LEFT JOIN RSP_ReklAnzBGAktGJ_M7_12 ON [Baugruppen-avista_out].[Baugruppen-ID] = RSP_ReklAnzBGAktGJ_M7_12.[Baugruppen-ID]
ORDER BY
  RSP_ReklAnzBGAktGJ.[Anzahl Reklamationen] DESC;
