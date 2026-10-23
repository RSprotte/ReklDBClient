SELECT
  Mastertab.ID,
  Mastertab.Auftragsdatum,
  Mastertab.[Meldungsnummer Kunde],
  tblFirma.Firma_Name AS Firma,
  Bearbeiter.Kurzzeichen AS [KZ-Service],
  Sachbearbeiter.Kurzzeichen AS [KZ-Sachbearbeiter],
  Mastertab.[Kategorie-Status],
  Taetigkeiten.Anzeige AS Taetigkeit,
  Mastertab.TaetigkeitID,
  ZuordnungDefTaetigkeit.Anzeige AS DefTaetigkeit,
  Mastertab.DefTaetigkeitID,
  Mastertab.[Kategorie-Problem],
  Verursacher.Anzeige AS Verursacher,
  Mastertab.[Lieferanten-Code 1],
  Lieferanten.Lieferantenname,
  Mastertab.ErfasserIntern,
  Sachbearbeiter_1.Kurzzeichen,
  Mastertab.intern,
  Mastertab.extern,
  ReklMemo.VorgangstypID,
  Mastertab.[Bestellnummer Kunde],
  Mastertab.LQEinheitID,
  Mastertab.LQBereichID,
  Mastertab.LQWorkstationID,
  LQWorkstation.[ID-LQFunktion],
  Kunden.Kontaktperson,
  Mastertab.Sachbearbeiter,
  Mastertab.[Zuständiger Service-MA],
  Mastertab.[Artikel-ID 1],
  Mastertab.[Baugruppen-ID],
  [Artikel-all_avista_out].AvistaArtikelTyp,
  Mastertab.LQEinheitErfasserID,
  Mastertab.LQBereichErfasserID,
  Mastertab.LQWorkstationErfasserID,
  LQWorkstation_1.[ID-LQFunktion],
  ReklIDtoTagID.TagID,
  Mastertab.[LINK-Nr - UBG],
  [Artikel-all_avista_out].[LQ Art-Nr],
  Lieferanten.VerantwortlichEinkauf,
  [Baugruppen-avista_out].[LQ Art-Nr],
  tblZuordnung.Zuordnung,
  tblZuordnung.Zuordnung_ID
FROM
  (
    (
      (
        (
          (
            (
              (
                (
                  (
                    (
                      (
                        (
                          (
                            (
                              (
                                Mastertab
                                LEFT JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]
                              )
                              LEFT JOIN Bearbeiter ON Mastertab.[Zuständiger Service-MA] = Bearbeiter.Name
                            )
                            LEFT JOIN Sachbearbeiter ON Mastertab.Sachbearbeiter = Sachbearbeiter.Name
                          )
                          LEFT JOIN Taetigkeiten ON Mastertab.TaetigkeitID = Taetigkeiten.ID
                        )
                        LEFT JOIN ZuordnungDefTaetigkeit ON Mastertab.DefTaetigkeitID = ZuordnungDefTaetigkeit.ID
                      )
                      LEFT JOIN Lieferanten ON Mastertab.[Lieferanten-Code 1] = Lieferanten.[Lieferant-ID]
                    )
                    LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID
                  )
                  LEFT JOIN LQWorkstation ON Mastertab.LQWorkstationID = LQWorkstation.ID
                )
                LEFT JOIN [Artikel-all_avista_out] ON Mastertab.[Artikel-ID 1] = [Artikel-all_avista_out].[Artikel-ID]
              )
              LEFT JOIN LQWorkstation AS LQWorkstation_1 ON Mastertab.LQWorkstationErfasserID = LQWorkstation_1.ID
            )
            LEFT JOIN ReklIDtoTagID ON Mastertab.ID = ReklIDtoTagID.ReklID
          )
          LEFT JOIN [Baugruppen-avista_out] ON Mastertab.[Baugruppen-ID] = [Baugruppen-avista_out].[Baugruppen-ID]
        )
        LEFT JOIN Sachbearbeiter AS Sachbearbeiter_1 ON Mastertab.ErfasserIntern = Sachbearbeiter_1.[SB-ID]
      )
      INNER JOIN Verursacher ON Mastertab.Verursacher = Verursacher.Verursacher
    )
    LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
  )
  LEFT JOIN tblZuordnung ON tblFirma.Zurodnung_ID = tblZuordnung.Zuordnung_ID
WHERE
  (
    (
      (Mastertab.Auftragsdatum) Between #5/8/2015#
      And #6/27/2016#
    )
    AND (
      (Mastertab.extern) = True
    )
  )
ORDER BY
  Mastertab.ID DESC;
