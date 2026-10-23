SELECT
  Mastertab.Auftragsdatum,
  Mastertab.ID,
  Mastertab.[Meldungsnummer Kunde],
  tblFirma.Firma_Name,
  Bearbeiter.Kurzzeichen AS [KZ-Service],
  Sachbearbeiter.Kurzzeichen AS [KZ-Sachbearbeiter],
  Mastertab.[Kategorie-Status],
  [Kategorien-Vorgangstatus].Anzeige AS Status_Vorgang,
  Mastertab.Rückmeldung_Kunde_Verschiebung,
  Taetigkeiten.Anzeige AS [Taetigkeit-Anzeige],
  Mastertab.TaetigkeitID,
  ZuordnungDefTaetigkeit.Anzeige AS [Fehlerbild-Anzeige],
  Mastertab.DefTaetigkeitID,
  Mastertab.[Kategorie-Problem],
  [Kategorien-Problemstatus].Anzeige AS Qualifizierung,
  Mastertab.[Aufwand-TRUMPF],
  Mastertab.Zeitbedarf,
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
  Kunden.Kontaktperson,
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
  Lieferanten.VerantwortlichEinkauf,
  [Baugruppen-avista_out].[LQ Art-Nr],
  ReklMemo.QSBearbeitungBis_Verl,
  Verursacher.Anzeige AS [Verursacher-Anzeige],
  Verursacher.Verursacher,
  Mastertab.Zuständig,
  Mastertab.TaetigkeitErfasserID,
  Mastertab.DefTaetigkeitErfasserID,
  tblZuordnung.Zuordnung,
  Lieferanten.LieferantenverantwortungLQ,
  tblZuordnung.Zuordnung_ID,
  tblFirma.Firma_ID,
  ReklMemo.Rückmeldung_QSV_Ist,
  Mastertab.[Fertigungsauftrag-FA],
  ReklMemo.BestellnrRepAuftrag,
  [Kategorien-Vorgangstatus].Anzeige,
  Mastertab.Maschinennummer
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
          LEFT JOIN [Kategorien-Vorgangstatus] ON Mastertab.[Kategorie-Status] = [Kategorien-Vorgangstatus].[Kategorie-Vorgangstatus]
        )
        LEFT JOIN Verursacher ON Mastertab.Verursacher = Verursacher.Verursacher
      )
      LEFT JOIN [Kategorien-Problemstatus] ON Mastertab.[Kategorie-Problem] = [Kategorien-Problemstatus].[Kategorie-Problemstatus]
    )
    LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
  )
  LEFT JOIN tblZuordnung ON tblFirma.Zurodnung_ID = tblZuordnung.Zuordnung_ID
ORDER BY
  Mastertab.ID DESC;
