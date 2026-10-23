SELECT
  tblFirma.Firma_Name,
  Mastertab.[Meldungsnummer Kunde] AS Meldung,
  [Baugruppen-avista_out].[Kunden-Nr] AS [BG-Nr],
  Mastertab.[Herstellungsdatum - BG] AS [BG-Datum],
  Mastertab.Unterbaugruppe AS UBG,
  Mastertab.[Herstellungsdatum - UBG] AS [UBG-Datum],
  [Artikel-all_avista_out].[Kd Mat-Nr] AS Artikel,
  Mastertab.[KW/JJ] AS [Art-Datum],
  Mastertab.[Menge betroffen - Kunde] AS Anzahl,
  Mastertab.Fehlerpositionen,
  Mastertab.[Welcher Fehler tritt auf] AS Fehlerbeschreibung,
  Mastertab.[Wie Warum passierte der Fehler] AS Fehlerursache,
  Mastertab.[Kategorie-Fehleranalyse] AS Fehleranalyse,
  Mastertab.Auftragsdatum,
  Mastertab.Rückmeldung_Kunde_Ist AS Antwort,
  Mastertab.ID AS [LQ-Vorgang],
  Mastertab.[langfristige Maßnahmen],
  Mastertab.[Umsetzung bis - IST] AS Umgesetzt,
  [Kategorien-Vorgangstatus].[Kategorie-Vorgangstatus] AS Status,
  Lieferanten.Lieferantenname AS Lieferant,
  "Qrtl " & Format(
    Mastertab!Auftragsdatum, "q / yyyy"
  ) AS Quartal
FROM
  (
    Lieferanten
    RIGHT JOIN (
      [Kategorien-Vorgangstatus]
      RIGHT JOIN (
        [Baugruppen-avista_out]
        RIGHT JOIN (
          [Artikel-all_avista_out]
          RIGHT JOIN (
            Kunden
            RIGHT JOIN Mastertab ON Kunden.[Kunden-ID] = Mastertab.[Kunden-ID]
          ) ON [Artikel-all_avista_out].[Artikel-ID] = Mastertab.[Artikel-ID 1]
        ) ON [Baugruppen-avista_out].[Baugruppen-ID] = Mastertab.[Baugruppen-ID]
      ) ON [Kategorien-Vorgangstatus].[Kategorie-Vorgangstatus] = Mastertab.[Kategorie-Status]
    ) ON Lieferanten.[Lieferant-ID] = Mastertab.[Lieferanten-Code 1]
  )
  LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
WHERE
  (
    (
      (Mastertab.ID) <> 6020
    )
    AND (
      (Mastertab.extern) = Yes
    )
    AND (
      (Mastertab.[TRUMPF?]) = Yes
    )
  )
ORDER BY
  Mastertab.Auftragsdatum DESC;
