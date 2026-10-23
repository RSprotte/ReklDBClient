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
  Mastertab.[Menge betroffen - Kunde],
  Mastertab.[Menge betroffen - LQ],
  Mastertab.Auftragsdatum,
  Mastertab.Rückmeldung_Kunde_Soll,
  Mastertab.Rückmeldung_Kunde_Verschiebung,
  Mastertab.[Wie Warum passierte der Fehler],
  Mastertab.[Vorgang - Pos20] AS Ausdr1,
  Mastertab.[Vorgang - Pos30] AS Ausdr2,
  Mastertab.[Vorgang - Pos40] AS Ausdr3,
  Mastertab.[Vorgang - Pos50] AS Ausdr4,
  Mastertab.[Vorgang - Pos60] AS Ausdr5,
  Mastertab.[Vorgang - Pos70] AS Ausdr6,
  Mastertab.[Vorgang - Pos80] AS Ausdr7,
  Mastertab.Verursacher,
  Mastertab.[Fertigungsauftrag-FA],
  Mastertab.[Lieferanten-Code 1],
  Lieferanten.Lieferantenname,
  Mastertab.[Sofortmaßnahmen (Aktion)],
  Mastertab.[langfristige Maßnahmen],
  Mastertab.MaßnahmenEinkaufLieferant,
  Mastertab.[QS-Lieferant] AS Ausdr8,
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
  [Artikel-all_avista_out].[Kd Mat-Nr] AS [ArtNr-TRUMPF],
  [Baugruppen-avista_out].[LQ Art-Nr] & "-" & [Baugruppen-avista_out].[Var Nr] AS BG,
  [Baugruppen-avista_out].[Kunden-Nr] AS [BG-TRUMPF],
  Bearbeiter.Kurzzeichen AS [KZ-Service],
  Sachbearbeiter.Kurzzeichen AS [KZ-Sachbearbeiter],
  Mastertab.[Kategorie-Fehleranalyse],
  tblFirma.Firma_Name,
  tblFirma.KndNr_Avista,
  Mastertab.[Zuständiger Service-MA],
  Kunden.Kontaktperson,
  Mastertab.[LINK-Nr - UBG],
  Mastertab.Fehlergruppe,
  ReklMemo.QSWeiterleitung,
  ReklMemo.QSBearbeitungBis_Ist,
  ReklMemo.QSBearbeitungBis_Soll,
  ReklMemo.QSBearbeitungBis_Verl,
  Mastertab.Fehlerart AS Ausdr9,
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
  Mastertab.Maschinennummer
FROM
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
  LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
ORDER BY
  Mastertab.ID DESC;
