SELECT
  tblTicketListe.*,
  tblTicketStatus.TicketStatusID,
  tblTicketTyp.Anzeige,
  tblTicketStatus.Anzeige,
  LQEinheit.LQEinheit,
  tblTicketListe.VerlängertBis,
  qryTicketSacharbeiterKz.Kurzzeichen,
  tblTicketStatus.Anzeige,
  [Kategorien-Vorgangstatus].Anzeige
FROM
  (
    (
      (
        Mastertab
        LEFT JOIN LQEinheit ON Mastertab.LQEinheitID = LQEinheit.ID
      )
      RIGHT JOIN (
        (
          tblTicketListe
          LEFT JOIN tblTicketStatus ON tblTicketListe.TicketStatusID = tblTicketStatus.TicketStatusID
        )
        LEFT JOIN tblTicketTyp ON tblTicketListe.TicketTypID = tblTicketTyp.TicketTypID
      ) ON Mastertab.ID = tblTicketListe.ReklID
    )
    LEFT JOIN qryTicketSacharbeiterKz ON tblTicketListe.ReklID = qryTicketSacharbeiterKz.VorgangID
  )
  LEFT JOIN [Kategorien-Vorgangstatus] ON Mastertab.[Kategorie-Status] = [Kategorien-Vorgangstatus].[Kategorie-Vorgangstatus]
ORDER BY
  tblTicketListe.VerlängertBis;
