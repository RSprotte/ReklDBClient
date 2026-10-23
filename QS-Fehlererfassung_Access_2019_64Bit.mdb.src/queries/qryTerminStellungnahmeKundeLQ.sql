SELECT
  Mastertab.ID,
  tblFirma.Firma_Name,
  Mastertab.Rückmeldung_Kunde_Verschiebung AS Soll_Kunde,
  tblTicketListe.VerlängertBis AS Ist_QS
FROM
  (
    (
      Mastertab
      LEFT JOIN tblTicketListe ON Mastertab.ID = tblTicketListe.ReklID
    )
    LEFT JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]
  )
  LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
WHERE
  (
    (
      (
        Mastertab.Rückmeldung_Kunde_Ist
      ) Is Null
    )
    AND (
      (Mastertab.[Kategorie-Status]) <> "Abgeschlossen"
    )
    AND (
      (Mastertab.extern) = True
    )
    AND (
      (tblTicketListe.TicketTypID) = 1
    )
    AND (
      (tblTicketListe.VerlängertBis) > [Rückmeldung_Kunde_Verschiebung]
    )
    AND (
      (tblTicketListe.TicketStatusID) <> 3
    )
  )
GROUP BY
  Mastertab.ID,
  tblFirma.Firma_Name,
  Mastertab.Rückmeldung_Kunde_Verschiebung,
  tblTicketListe.VerlängertBis;
