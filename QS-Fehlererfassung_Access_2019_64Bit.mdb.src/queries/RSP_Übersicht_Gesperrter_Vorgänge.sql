SELECT
  [Mastertab-DSBerechtigung].*
FROM
  [Mastertab-DSBerechtigung]
WHERE
  (
    (
      (
        [Mastertab-DSBerechtigung].Kürzel
      ) = "AK"
    )
    AND (
      (
        [Mastertab-DSBerechtigung].Zustand
      ) = "Gesperrt"
    )
  );
