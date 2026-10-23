SELECT
  RSP_Kostenaufstellung.[A-DatumJahr],
  RSP_Kostenaufstellung.[Kategorie-Problem],
  Sum(
    RSP_Kostenaufstellung.[Fracht-Kosten]
  ) AS [SummevonFracht-Kosten],
  Sum(
    RSP_Kostenaufstellung.[Material-Kosten]
  ) AS [SummevonMaterial-Kosten],
  Sum(
    RSP_Kostenaufstellung.GesamtKosten
  ) AS SummevonGesamtKosten,
  RSP_Kostenaufstellung.[TRUMPF?]
FROM
  RSP_Kostenaufstellung
GROUP BY
  RSP_Kostenaufstellung.[A-DatumJahr],
  RSP_Kostenaufstellung.[Kategorie-Problem],
  RSP_Kostenaufstellung.[TRUMPF?]
HAVING
  (
    (
      (
        RSP_Kostenaufstellung.[TRUMPF?]
      ) = True
    )
  );
