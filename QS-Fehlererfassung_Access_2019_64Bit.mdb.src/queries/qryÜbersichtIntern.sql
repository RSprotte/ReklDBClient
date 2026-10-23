SELECT
  Mastertab.Auftragsdatum,
  Mastertab.ID,
  Mastertab.[Meldungsnummer Kunde],
  tblFirma.Firma_Name,
  Bearbeiter.Kurzzeichen AS [KZ-Service],
  Sachbearbeiter.Kurzzeichen AS [KZ-Sachbearbeiter],
  Mastertab.[Kategorie-Status],
  [Kategorien-Vorgangstatus].Anzeige AS [Vorgang-Status],
  Mastertab.Rückmeldung_Kunde_Verschiebung,
  Taetigkeiten.Anzeige AS [Taetigkeit-Anzeige],
  Mastertab.TaetigkeitID,
  ZuordnungDefTaetigkeit.Anzeige AS [Fehlerbild-Anzeige],
  Mastertab.DefTaetigkeitID,
  Mastertab.[Kategorie-Problem],
  Mastertab.[Aufwand-TRUMPF],
  Mastertab.Zeitbedarf,
  Verursacher.Anzeige AS [Verursacher-Anzeige],
  Mastertab.[Lieferanten-Code 1],
  Lieferanten.Lieferantenname,
  Mastertab.Rückmeldung_Kunde_Ist,
  Mastertab.ErfasserIntern,
  Mastertab.intern,
  Mastertab.extern,
  Mastertab.Rückmeldung_Kunde_Soll,
  ReklMemo.VorgangstypID,
  Mastertab.[Bestellnummer Kunde],
  Mastertab.LQEinheitID,
  Mastertab.LQBereichID,
  Mastertab.LQWorkstationID,
  LQWorkstation.[ID-LQFunktion],
  Mastertab.Sachbearbeiter,
  Mastertab.[Zuständiger Service-MA],
  Mastertab.[Artikel-ID 1],
  Mastertab.[Baugruppen-ID],
  Mastertab.[LINK-Nr - UBG],
  [Artikel-all_avista_out].AvistaArtikelTyp,
  Mastertab.LQEinheitErfasserID,
  Mastertab.LQBereichErfasserID,
  Mastertab.LQWorkstationErfasserID,
  LQWorkstation_1.[ID-LQFunktion],
  ReklIDtoTagID.TagID,
  Mastertab.[LINK-Nr - UBG],
  [Artikel-all_avista_out].[LQ Art-Nr],
  [Baugruppen-avista_out].[LQ Art-Nr],
  Lieferanten.VerantwortlichEinkauf,
  Kunden.Kurzzeichen AS [KZ-Melder],
  Kunden_1.Kurzzeichen AS [KZ-Erfasser],
  Mastertab.[Welcher Fehler tritt auf],
  Mastertab.Zuständig,
  Verursacher.Verursacher,
  Mastertab.TaetigkeitErfasserID,
  Mastertab.DefTaetigkeitErfasserID,
  LQEinheit.LQEinheit,
  Mastertab.[Fertigungsauftrag-FA],
  Sachbearbeiter.Name,
  Mastertab.[Fertigungsauftrag-FA]
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
                                (
                                  Mastertab
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
            LEFT JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]
          )
          LEFT JOIN [Kategorien-Vorgangstatus] ON Mastertab.[Kategorie-Status] = [Kategorien-Vorgangstatus].[Kategorie-Vorgangstatus]
        )
        LEFT JOIN Verursacher ON Mastertab.Verursacher = Verursacher.Verursacher
      )
      LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
    )
    LEFT JOIN LQEinheit ON Mastertab.LQEinheitID = LQEinheit.ID
  )
  LEFT JOIN Kunden AS Kunden_1 ON Mastertab.ErfasserIntern = Kunden_1.[Kunden-ID]
WHERE
  (
    (
      (Mastertab.intern) = True
    )
  )
ORDER BY
  Mastertab.ID DESC;
