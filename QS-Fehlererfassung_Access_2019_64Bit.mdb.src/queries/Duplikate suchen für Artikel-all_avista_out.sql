SELECT
  [Artikel-all_avista_out].[Artikel-ID],
  [Artikel-all_avista_out].[LQ Art-Nr],
  [Artikel-all_avista_out].Variante,
  [Artikel-all_avista_out].[Kd Mat-Nr],
  [Artikel-all_avista_out].[Hersteller-Teilenummer]
FROM
  [Artikel-all_avista_out]
WHERE
  (
    (
      (
        [Artikel-all_avista_out].[LQ Art-Nr]
      ) In (
        SELECT
          [LQ Art-Nr]
        FROM
          [Artikel-all_avista_out] As Tmp
        GROUP BY
          [LQ Art-Nr],
          [Variante],
          [Kd Mat-Nr]
        HAVING
          Count(*) > 1
          And [Variante] = [Artikel-all_avista_out].[Variante]
          And [Kd Mat-Nr] = [Artikel-all_avista_out].[Kd Mat-Nr]
      )
    )
  )
ORDER BY
  [Artikel-all_avista_out].[LQ Art-Nr],
  [Artikel-all_avista_out].Variante,
  [Artikel-all_avista_out].[Kd Mat-Nr];
