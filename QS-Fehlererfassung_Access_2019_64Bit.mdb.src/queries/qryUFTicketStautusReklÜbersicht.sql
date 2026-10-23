SELECT
  tblTicketListe.ReklID,
  tblTicketTyp.Anzeige,
  tblTicketListe.BearbeitungBis,
  tblTicketListe.VerlängertBis,
  tblTicketMahnungStatus.Anzeige AS Mahnstufe
FROM
  (
    (
      (
        tblTicketListe
        LEFT JOIN tblTicketTyp ON tblTicketListe.TicketTypID = tblTicketTyp.TicketTypID
      )
      LEFT JOIN tblTicketStatus ON tblTicketListe.TicketStatusID = tblTicketStatus.TicketStatusID
    )
    LEFT JOIN tblTicketMahnwesen ON tblTicketListe.TicketMahnwesenID = tblTicketMahnwesen.TicketMahnungID
  )
  LEFT JOIN tblTicketMahnungStatus ON tblTicketMahnwesen.TicketMahnungStatusID = tblTicketMahnungStatus.TicketMahnungStatusID
WHERE
  (
    (
      (tblTicketStatus.Anzeige) Not Like "Erstellt"
    )
    AND (
      (tblTicketListe.ErledigtAm) Is Null
    )
  )
GROUP BY
  tblTicketListe.ReklID,
  tblTicketTyp.Anzeige,
  tblTicketListe.BearbeitungBis,
  tblTicketListe.VerlängertBis,
  tblTicketMahnungStatus.Anzeige;
