SELECT
  [Mastertab-DSBerechtigung].ID,
  [Mastertab-DSBerechtigung].Zustand,
  [Mastertab-DSBerechtigung].Kürzel,
  [Mastertab-DSBerechtigung].Datum,
  Int(
    DateDiff("n", [Uhrzeit_Neu], [Jetzt]) / 60
  )& ":" & Format(
    (
      Abs(
        DateDiff("n", [Uhrzeit_Neu], [Jetzt]) Mod 60
      )
    ),
    "00"
  ) AS Seit,
  Format(
    Now(),
    "hh:nn:ss"
  ) AS Jetzt,
  set_SplittenAktuellerEintrag([Aktueller Eintrag]) AS Uhrzeit,
  TimeValue([Uhrzeit]) AS Uhrzeit_Neu,
  [Mastertab-DSBerechtigung].Client
FROM
  [Mastertab-DSBerechtigung]
  INNER JOIN ReklMemo ON [Mastertab-DSBerechtigung].ID = ReklMemo.ID
WHERE
  (
    (
      (
        [Mastertab-DSBerechtigung].Zustand
      ) = "Gesperrt"
    )
  );
