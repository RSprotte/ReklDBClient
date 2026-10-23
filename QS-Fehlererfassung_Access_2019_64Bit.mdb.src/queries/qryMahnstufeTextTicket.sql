SELECT
  tblTicketMahnwesen.TicketID,
  tblTicketMahnungStatus.Anzeige,
  tblTicketListe.ReklID,
  tblTicketMahnwesen.Datum,
  tblTicketMahnwesen.TicketMahnungStatusID
FROM
  (
    tblTicketMahnwesen
    INNER JOIN tblTicketMahnungStatus ON tblTicketMahnwesen.TicketMahnungStatusID = tblTicketMahnungStatus.TicketMahnungStatusID
  )
  LEFT JOIN tblTicketListe ON tblTicketMahnwesen.TicketID = tblTicketListe.TicketID
GROUP BY
  tblTicketMahnwesen.TicketID,
  tblTicketMahnungStatus.Anzeige,
  tblTicketListe.ReklID,
  tblTicketMahnwesen.Datum,
  tblTicketMahnwesen.TicketMahnungStatusID
ORDER BY
  tblTicketMahnwesen.Datum DESC;
