SELECT
  Mastertab.ID,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.Rückmeldung_Kunde_Verschiebung,
  tblTicketListe.VerlängertBis,
  getDateAddTage([VerlängertBis], 10) AS TerminKund,
  getDateAddTage([VerlängertBis], 5) AS TerminQSLQ,
  Mastertab.Rückmeldung_Kunde_Ist
FROM
  Mastertab
  LEFT JOIN (
    tblTicketListe
    LEFT JOIN tblTicketTyp ON tblTicketListe.TicketTypID = tblTicketTyp.TicketTypID
  ) ON Mastertab.ID = tblTicketListe.ReklID
WHERE
  (
    (
      (tblTicketListe.ErledigtAm) Is Null
    )
    AND (
      (tblTicketTyp.TicketTypID) = 2
    )
    AND (
      (Mastertab.extern) = True
    )
  )
GROUP BY
  Mastertab.ID,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.Rückmeldung_Kunde_Verschiebung,
  tblTicketListe.VerlängertBis,
  getDateAddTage([VerlängertBis], 10),
  getDateAddTage([VerlängertBis], 5),
  Mastertab.Rückmeldung_Kunde_Ist
HAVING
  (
    (
      (
        Mastertab.Rückmeldung_Kunde_Verschiebung
      ) <= tblTicketListe!VerlängertBis
    )
    And (
      (
        Mastertab.Rückmeldung_Kunde_Ist
      ) Is Null
    )
  );
