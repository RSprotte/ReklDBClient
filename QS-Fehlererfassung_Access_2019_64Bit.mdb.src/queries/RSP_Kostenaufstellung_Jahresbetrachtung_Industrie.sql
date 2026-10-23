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
  RSP_Kostenaufstellung.[TRUMPF?],
  RSP_Kostenaufstellung.extern
FROM
  RSP_Kostenaufstellung
GROUP BY
  RSP_Kostenaufstellung.[A-DatumJahr],
  RSP_Kostenaufstellung.[Kategorie-Problem],
  RSP_Kostenaufstellung.[TRUMPF?],
  RSP_Kostenaufstellung.extern
HAVING
  (
    (
      (
        RSP_Kostenaufstellung.[TRUMPF?]
      ) = False
    )
    AND (
      (RSP_Kostenaufstellung.extern) = True
    )
  );
