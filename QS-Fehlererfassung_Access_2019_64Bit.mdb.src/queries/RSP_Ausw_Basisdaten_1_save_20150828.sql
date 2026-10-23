SELECT
  Mastertab.ID,
  Mastertab.Datum,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.[Bestellnummer Kunde],
  Mastertab.[Welcher Fehler tritt auf],
  Mastertab.[Bemerkungen - Maßnahmen],
  Mastertab.intern,
  Mastertab.extern,
  Mastertab.[TRUMPF?],
  Mastertab.ErfasserIntern,
  Kunden_1.Kontaktperson AS Erfasser,
  Kunden_1.Kurzzeichen AS [KZ-Erfasser],
  Mastertab.[Menge betroffen - Kunde],
  Mastertab.[Menge betroffen - LQ],
  Mastertab.Auftragsdatum,
  Mastertab.Rückmeldung_Kunde_Soll,
  Mastertab.Rückmeldung_Kunde_Verschiebung,
  Mastertab.[Wie Warum passierte der Fehler],
  Mastertab.Verursacher,
  LQEinheit.LQEinheit,
  LQBereich.LQBereich,
  LQWorkstation.LQWorkstation,
  LQFunktion.LQFunktion,
  Mastertab.[Fertigungsauftrag-FA],
  Mastertab.FALosgroesse,
  Mastertab.[Lieferanten-Code 1],
  Lieferanten.Lieferantenname,
  Mastertab.[Sofortmaßnahmen (Aktion)],
  Mastertab.[langfristige Maßnahmen],
  Mastertab.MaßnahmenEinkaufLieferant,
  Mastertab.[QS-Kunde],
  Mastertab.[Status: erledigt],
  Mastertab.Rückmeldung_Lieferant_Soll,
  Mastertab.Rückmeldung_Lieferant_Ist,
  Mastertab.Rückmeldung_Kunde_Ist,
  Mastertab.[Kategorie-Problem],
  Mastertab.[Kategorie-Status],
  Mastertab.Sachbearbeiter,
  Mastertab.[Kategorie-Bearbeitung II],
  Mastertab.[Kategorie-Bearbeitung I],
  Mastertab.Fehlerpositionen,
  [Artikel-all_avista_out].[LQ Art-Nr] & "-" & [Artikel-all_avista_out].Variante AS ArtNr,
  [Artikel-all_avista_out].[Hersteller-Teilenummer],
  Mastertab.Seriennummer,
  Mastertab.[KW/JJ],
  [Artikel-all_avista_out].[Kd Mat-Nr] AS [ArtNr-TRUMPF],
  [Baugruppen-avista_out].[LQ Art-Nr] & "-" & [Baugruppen-avista_out].[Var Nr] AS BG,
  [Baugruppen-avista_out].[Kunden-Nr] AS [BG-TRUMPF],
  Bearbeiter.Kurzzeichen AS [KZ-Service],
  Sachbearbeiter.Kurzzeichen AS [KZ-Sachbearbeiter],
  Mastertab.[Kategorie-Fehleranalyse],
  Mastertab.[Kunden-ID],
  Kunden.Kurzzeichen AS [KZ-Melder],
  Kunden.Firma,
  Kunden.Kundennummer,
  Mastertab.[Zuständiger Service-MA],
  Kunden.Kontaktperson,
  Mastertab.[LINK-Nr - UBG],
  Mastertab.Fehlergruppe,
  ReklMemo.QSWeiterleitung,
  ReklMemo.QSBearbeitungBis_Ist,
  ReklMemo.QSBearbeitungBis_Soll,
  ReklMemo.QSBearbeitungBis_Verl,
  Sachbearbeiter.Einheit,
  Sachbearbeiter.Bereich,
  Mastertab.[Mitarbeiter-Kurzzeichen],
  Mastertab.[Fracht-Kosten],
  Mastertab.[Material-Kosten],
  Mastertab.Zeitbedarf,
  Mastertab.[Kosten-Aufwand],
  Mastertab.[WE-Datum],
  Mastertab.[Aufwand-TRUMPF],
  Lieferanten.VerantwortlichEinkauf,
  Lieferanten.LieferantenverantwortungLQ,
  Sachbearbeiter.EmailName,
  Sachbearbeiter.Name,
  Sachbearbeiter.TelefonBeruflich,
  Sachbearbeiter.Strasse,
  Sachbearbeiter.PLZ,
  Sachbearbeiter.Ort,
  Mastertab.Maschinennummer,
  Mastertab.[Weitere Infos],
  Mastertab.[Rücklieferung an Lieferant],
  Mastertab.[Annahme mit Prüfausnahme],
  Mastertab.[Annahme mit Nacharbeit],
  Mastertab.[QS-Bearbeitung],
  Mastertab.[PE-Bearbeitung],
  Taetigkeiten.Taetigkeit,
  ZuordnungDefTaetigkeit.DefTaetigkeit,
  ZuordnungAnalyseFehlerbild.AnalyseFehlerbild,
  Mastertab.LQEinheitID,
  Mastertab.LQBereichID,
  Mastertab.LQWorkstationID,
  Mastertab.LQEinheitErfasserID,
  Mastertab.LQBereichErfasserID,
  Mastertab.LQWorkstationErfasserID,
  LQEinheit_1.LQEinheit AS [Erfasser-Einheit],
  LQBereich_1.LQBereich AS [Erfasser-Bereich],
  LQWorkstation_1.LQWorkstation AS [Erfasser-WS],
  LQFunktion_1.LQFunktion AS Erfasser_Funktion,
  Kunden.Zuordnung,
  [Artikel-all_avista_out].[Bezeichnung 1],
  [Baugruppen-avista_out].[Bezeichnung 1],
  ROS_Daten_Einkauf.TaetigkeitLieferant,
  Kunden.Zuordnung
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
                              Lieferanten
                              RIGHT JOIN (
                                Kunden
                                RIGHT JOIN (
                                  Bearbeiter
                                  RIGHT JOIN (
                                    Sachbearbeiter
                                    RIGHT JOIN (
                                      [Artikel-all_avista_out]
                                      RIGHT JOIN (
                                        [Baugruppen-avista_out]
                                        RIGHT JOIN Mastertab ON [Baugruppen-avista_out].[Baugruppen-ID] = Mastertab.[Baugruppen-ID]
                                      ) ON [Artikel-all_avista_out].[Artikel-ID] = Mastertab.[Artikel-ID 1]
                                    ) ON Sachbearbeiter.Name = Mastertab.Sachbearbeiter
                                  ) ON Bearbeiter.Name = Mastertab.[Zuständiger Service-MA]
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
                    LEFT JOIN Kunden AS Kunden_1 ON Mastertab.ErfasserIntern = Kunden_1.[Kunden-ID]
                  )
                  LEFT JOIN Taetigkeiten ON Mastertab.TaetigkeitID = Taetigkeiten.ID
                )
                LEFT JOIN ZuordnungDefTaetigkeit ON Mastertab.DefTaetigkeitID = ZuordnungDefTaetigkeit.ID
              )
              LEFT JOIN ZuordnungAnalyseFehlerbild ON Mastertab.DefTaetigkeitAnalyseID = ZuordnungAnalyseFehlerbild.ID
            )
            LEFT JOIN LQFunktion ON LQWorkstation.[ID-LQFunktion] = LQFunktion.ID
          )
          LEFT JOIN LQBereich AS LQBereich_1 ON Mastertab.LQBereichErfasserID = LQBereich_1.ID
        )
        LEFT JOIN LQWorkstation AS LQWorkstation_1 ON Mastertab.LQWorkstationErfasserID = LQWorkstation_1.ID
      )
      LEFT JOIN LQEinheit AS LQEinheit_1 ON Mastertab.LQEinheitErfasserID = LQEinheit_1.ID
    )
    LEFT JOIN LQFunktion AS LQFunktion_1 ON LQWorkstation_1.[ID-LQFunktion] = LQFunktion_1.ID
  )
  LEFT JOIN ROS_Daten_Einkauf ON Mastertab.ID = ROS_Daten_Einkauf.[ID-ReklNr]
ORDER BY
  Mastertab.ID DESC;
