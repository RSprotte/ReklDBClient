SELECT
  Mastertab.ID,
  [LQ Art-Nr] & ("-")& [Var Nr] AS BG,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.[Bestellnummer Kunde],
  tblFirma.KndNr_Avista,
  Mastertab.Rückmeldung_Kunde_Soll,
  Bearbeiter.Kurzzeichen,
  ReklMemo.BestellnrRepAuftrag,
  ReklMemo.AnlieferterminRep,
  Mastertab.[Menge betroffen - Kunde]
FROM
  (
    (
      (
        (
          Mastertab
          LEFT JOIN [Baugruppen-avista_out] ON Mastertab.[Baugruppen-ID] = [Baugruppen-avista_out].[Baugruppen-ID]
        )
        LEFT JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]
      )
      LEFT JOIN Bearbeiter ON Mastertab.[Zuständiger Service-MA] = Bearbeiter.Name
    )
    LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
  )
  INNER JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID;
