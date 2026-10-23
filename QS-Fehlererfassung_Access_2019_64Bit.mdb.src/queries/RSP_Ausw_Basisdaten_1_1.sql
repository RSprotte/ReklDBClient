SELECT
  Mastertab.ID,
  Mastertab.Auftragsdatum,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.[Bestellnummer Kunde],
  Mastertab.[Welcher Fehler tritt auf],
  Mastertab.intern,
  Mastertab.extern,
  Verursacher.Anzeige AS Verursacher,
  LQEinheit.LQEinheit,
  LQBereich.LQBereich,
  LQWorkstation.LQWorkstation,
  Lieferanten.Lieferantenname,
  Mastertab.[Kategorie-Problem],
  Mastertab.[Kategorie-Status],
  [Artikel-all_avista_out].[LQ Art-Nr] & "-" & [Artikel-all_avista_out].Variante AS ArtNr,
  [Baugruppen-avista_out].[LQ Art-Nr] & "-" & [Baugruppen-avista_out].[Var Nr] AS BG,
  tblFirma.Firma_Name AS Firma,
  tblFirma.KndNr_Avista AS Kundennummer,
  Kunden.Kontaktperson,
  Taetigkeiten.Anzeige AS Taetigkeit,
  ZuordnungDefTaetigkeit.Anzeige AS DefTaetigkeit,
  [Artikel-all_avista_out].AvistaArtikelTyp,
  Vorgangstyp.Vorgangstyp,
  Lieferanten.VerantwortlichEinkauf,
  Taetigkeiten_1.Anzeige AS TaetigkeitLieferant,
  ROS_Daten_Einkauf.Wiederholungsfehler,
  tblZuordnung.Zuordnung,
  Mastertab.ErfasserIntern,
  Mastertab.[Lieferanten-Code 1],
  tblZuordnung.Zuordnung_ID,
  tblFirma.Firma_Name,
  Mastertab.[Fertigungsauftrag-FA],
  Mastertab.Sachbearbeiter,
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
                            Lieferanten
                            RIGHT JOIN (
                              Kunden
                              RIGHT JOIN (
                                [Artikel-all_avista_out]
                                RIGHT JOIN (
                                  [Baugruppen-avista_out]
                                  RIGHT JOIN Mastertab ON [Baugruppen-avista_out].[Baugruppen-ID] = Mastertab.[Baugruppen-ID]
                                ) ON [Artikel-all_avista_out].[Artikel-ID] = Mastertab.[Artikel-ID 1]
                              ) ON Kunden.[Kunden-ID] = Mastertab.[Kunden-ID]
                            ) ON Lieferanten.[Lieferant-ID] = Mastertab.[Lieferanten-Code 1]
                          )
                          LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID
                        )
                        LEFT JOIN LQEinheit ON Mastertab.LQEinheitID = LQEinheit.ID
                      )
                      LEFT JOIN LQBereich ON Mastertab.LQBereichID = LQBereich.ID
                    )
                    LEFT JOIN LQWorkstation ON Mastertab.LQWorkstationID = LQWorkstation.ID
                  )
                  LEFT JOIN Taetigkeiten ON Mastertab.TaetigkeitID = Taetigkeiten.ID
                )
                LEFT JOIN ZuordnungDefTaetigkeit ON Mastertab.DefTaetigkeitID = ZuordnungDefTaetigkeit.ID
              )
              LEFT JOIN AvistaArtikelTyp ON [Artikel-all_avista_out].AvistaArtikelTyp = AvistaArtikelTyp.[AvistaArtikelTyp-ID]
            )
            LEFT JOIN Vorgangstyp ON ReklMemo.VorgangstypID = Vorgangstyp.VorgangstypID
          )
          LEFT JOIN Verursacher ON Mastertab.Verursacher = Verursacher.Verursacher
        )
        LEFT JOIN ROS_Daten_Einkauf ON Mastertab.ID = ROS_Daten_Einkauf.[ID-ReklNr]
      )
      LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
    )
    LEFT JOIN tblZuordnung ON tblFirma.Zurodnung_ID = tblZuordnung.Zuordnung_ID
  )
  LEFT JOIN Taetigkeiten AS Taetigkeiten_1 ON ROS_Daten_Einkauf.TaetigkeitLieferantID = Taetigkeiten_1.ID
ORDER BY
  Mastertab.ID DESC;
