SELECT
  Mastertab.ID,
  Sachbearbeiter.[SB-ID],
  Sachbearbeiter.Kurzzeichen,
  LQEinheit.LQEinheit,
  ReklMemo.QSWeiterleitung,
  ReklMemo.QSBearbeitungBis_Soll,
  ReklMemo.QSBearbeitungBis_Verl,
  Sachbearbeiter_1.EmailName AS MailVon,
  LQEinheit.MailTOLQEinheit AS MailTo,
  LQEinheit.MailCCLQEinheit AS MailCC,
  LQEinheit.MailEskalationLQEinheit AS MailBCC,
  Mastertab.[Kategorie-Status]
FROM
  (
    (
      (
        Mastertab
        LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID
      )
      INNER JOIN Sachbearbeiter ON Mastertab.Sachbearbeiter = Sachbearbeiter.Name
    )
    INNER JOIN LQEinheit ON Sachbearbeiter.Einheit = LQEinheit.LQEinheit
  )
  INNER JOIN Sachbearbeiter AS Sachbearbeiter_1 ON Mastertab.[Zuständiger Service-MA] = Sachbearbeiter_1.Name
WHERE
  (
    (
      (ReklMemo.Rückmeldung_QSV_Ist) Is Null
    )
    AND (
      (Mastertab.Auftragsdatum) Between #1/1/2017#
      And Date()
    )
    AND (
      (Mastertab.extern) = True
    )
  )
GROUP BY
  Mastertab.ID,
  Sachbearbeiter.[SB-ID],
  Sachbearbeiter.Kurzzeichen,
  LQEinheit.LQEinheit,
  ReklMemo.QSWeiterleitung,
  ReklMemo.QSBearbeitungBis_Soll,
  ReklMemo.QSBearbeitungBis_Verl,
  Sachbearbeiter_1.EmailName,
  LQEinheit.MailTOLQEinheit,
  LQEinheit.MailCCLQEinheit,
  LQEinheit.MailEskalationLQEinheit,
  Mastertab.[Kategorie-Status]
HAVING
  (
    (
      (ReklMemo.QSWeiterleitung) Is Not Null
    )
    AND (
      (ReklMemo.QSBearbeitungBis_Soll) Is Not Null
    )
    AND (
      (Mastertab.[Kategorie-Status]) = "weiter an: QS-Bearbeitung"
    )
  )
  OR (
    (
      (Mastertab.[Kategorie-Status]) = "Bearbeitung: QS"
    )
  )
  OR (
    (
      (Mastertab.[Kategorie-Status]) = "Klärung: Lieferant/QS"
    )
  )
  OR (
    (
      (Mastertab.[Kategorie-Status]) = "Klärung: Lieferant/Einkauf"
    )
  );
