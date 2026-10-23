SELECT
  tblTicketListe.ReklID,
  tblTicketListe.MailDatum,
  tblTicketListe.BearbeitungBis,
  tblTicketListe.VerlängertBis,
  tblTicketTyp.Anzeige,
  tblTicketListe.MailBetreff,
  tblTicketMahnungStatus.Anzeige,
  tblTicketMahnwesen.Datum
FROM
  (
    tblTicketMahnungStatus
    INNER JOIN (
      tblTicketListe
      INNER JOIN tblTicketMahnwesen ON tblTicketListe.TicketID = tblTicketMahnwesen.TicketID
    ) ON tblTicketMahnungStatus.TicketMahnungStatusID = tblTicketMahnwesen.TicketMahnungStatusID
  )
  INNER JOIN tblTicketTyp ON tblTicketListe.TicketTypID = tblTicketTyp.TicketTypID
WHERE
  (
    (
      (tblTicketTyp.TicketTypID) = 2
    )
  )
GROUP BY
  tblTicketListe.ReklID,
  tblTicketListe.MailDatum,
  tblTicketListe.BearbeitungBis,
  tblTicketListe.VerlängertBis,
  tblTicketTyp.Anzeige,
  tblTicketListe.MailBetreff,
  tblTicketMahnungStatus.Anzeige,
  tblTicketMahnwesen.Datum
ORDER BY
  tblTicketMahnwesen.Datum DESC;
