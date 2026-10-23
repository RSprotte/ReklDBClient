SELECT
  Mastertab.ID,
  Sachbearbeiter.[SB-ID],
  Mastertab.Rückmeldung_Lieferant_Soll,
  Mastertab.Rückmeldung_Lieferant_Verlängertl,
  LQEinheit.MailTOLQEinheit AS MailVon,
  RSP_LieferantAnsprechpartner.Email AS MailTo,
  LQEinheit.MailTOLQEinheit AS MailCC,
  Mastertab.[Meldungsnummer Kunde],
  Sachbearbeiter.TelefonBeruflich,
  LQEinheit.MailTOLQEinheit,
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
  INNER JOIN RSP_LieferantAnsprechpartner ON Mastertab.[LiefAnsprechp-ID] = RSP_LieferantAnsprechpartner.AnsprechpartnerID
WHERE
  (
    (
      (
        Mastertab.Rückmeldung_Lieferant_Ist
      ) Is Null
    )
    AND (
      (Mastertab.Auftragsdatum) Between #1/1/2017#
      And Date()
    )
  )
GROUP BY
  Mastertab.ID,
  Sachbearbeiter.[SB-ID],
  Mastertab.Rückmeldung_Lieferant_Soll,
  Mastertab.Rückmeldung_Lieferant_Verlängertl,
  RSP_LieferantAnsprechpartner.Email,
  Mastertab.[Meldungsnummer Kunde],
  Sachbearbeiter.TelefonBeruflich,
  LQEinheit.MailTOLQEinheit,
  Mastertab.[Kategorie-Status],
  LQEinheit.MailTOLQEinheit
HAVING
  (
    (
      (Mastertab.[Kategorie-Status]) = "Klärung: Lieferant/Einkauf"
      Or (Mastertab.[Kategorie-Status]) = "Klärung: Lieferant/QS"
    )
  );
