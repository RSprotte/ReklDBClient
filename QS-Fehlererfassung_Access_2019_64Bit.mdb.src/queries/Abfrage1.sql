SELECT
  Mastertab.ID,
  Mastertab.extern,
  Mastertab.[Kategorie-Status],
  Kunden.Kontaktperson,
  Mastertab.[Wie Warum passierte der Fehler],
  Mastertab.LQEinheitID,
  Mastertab.LQBereichID,
  Mastertab.LQWorkstationID,
  MIDZuReklID.ReklID,
  MIDZuReklID.MID
FROM
  (
    (
      (
        (
          Mastertab
          LEFT JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]
        )
        LEFT JOIN LQEinheit ON Mastertab.LQEinheitID = LQEinheit.ID
      )
      LEFT JOIN LQBereich ON Mastertab.LQBereichID = LQBereich.ID
    )
    LEFT JOIN LQWorkstation ON Mastertab.LQWorkstationID = LQWorkstation.ID
  )
  LEFT JOIN MIDZuReklID ON Mastertab.ID = MIDZuReklID.ReklID
WHERE
  (
    (
      (Mastertab.ID) > 83300
    )
    AND (
      (Mastertab.extern) = True
    )
    AND (
      (Mastertab.[Kategorie-Status]) <> "Stellungnahme Klärung Kunde"
    )
    AND (
      (Kunden.Kontaktperson) = "Kraus, Fabian"
    )
  );
