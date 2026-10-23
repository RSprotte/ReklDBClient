SELECT
  Sachbearbeiter.Kurzzeichen,
  Sachbearbeiter.Einheit
FROM
  Sachbearbeiter
WHERE
  (
    (
      (Sachbearbeiter.Einheit) = "Einkauf"
    )
  );
