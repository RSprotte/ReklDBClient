SELECT
  tblTicketListe.TicketID,
  tblTicketListe.ReklID,
  tblTicketListe.TicketTypID,
  tblTicketListe.VerlängertBis,
  getDelayTwoDates(
    Date(),
    [VerlängertBis]
  ) AS Tage,
  tblTicketListe.TicketMahnwesenID,
  tblTicketTyp.ErinnerungTage,
  tblTicketTyp.Mahnung1Tage,
  tblTicketTyp.Mahnung2Tage,
  tblTicketListe.TicketStatusID,
  tblTicketTyp.MahnwesenAktiv
FROM
  (
    tblTicketListe
    LEFT JOIN tblTicketTyp ON tblTicketListe.TicketTypID = tblTicketTyp.TicketTypID
  )
  LEFT JOIN tblTicketStatus ON tblTicketListe.TicketStatusID = tblTicketStatus.TicketStatusID
WHERE
  (
    (
      (tblTicketListe.VerlängertBis) <= Date()
    )
    AND (
      (
        tblTicketListe.TicketMahnwesenID
      ) > 0
    )
    AND (
      (tblTicketListe.TicketStatusID) <> 3
    )
    AND (
      (tblTicketTyp.MahnwesenAktiv) = True
    )
  );
