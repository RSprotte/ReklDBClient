SELECT
  Mastertab.ID,
  tblFirma.Firma_Name AS Firma,
  Kunden.Kontaktperson,
  Mastertab.[Meldungsnummer Kunde] AS ReklNr_KD,
  Mastertab.Auftragsdatum,
  Kunden.Telefonnummer,
  Kunden.EmailAdresse,
  Mastertab.Maschinennummer,
  Mastertab.[Menge betroffen - Kunde] AS FehlermengeKD,
  Mastertab.[Bestellnummer Kunde] AS BestNrKD,
  [Baugruppen-avista_out].[Kunden-Nr] AS BG_KD,
  [Baugruppen-avista_out].[LQ Art-Nr] AS BG_LQ,
  [Artikel-all_avista_out].[Kd Mat-Nr] AS Art_KD,
  [Artikel-all_avista_out].[LQ Art-Nr] AS Art_LQ,
  [Artikel-all_avista_out].Variante,
  Mastertab.[Welcher Fehler tritt auf] AS Fehler,
  Mastertab.[Zuständiger Service-MA] AS Service_MA,
  Mastertab.Sachbearbeiter,
  Mastertab.[Wie Warum passierte der Fehler] AS Fehlerursache,
  Mastertab.[langfristige Maßnahmen] AS Maßnahmen,
  Mastertab.[Kategorie-Problem],
  Mastertab.Referenznummer,
  Mastertab.Fehlerpositionen
FROM
  (
    (
      (
        Mastertab
        LEFT JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]
      )
      LEFT JOIN [Baugruppen-avista_out] ON Mastertab.[Baugruppen-ID] = [Baugruppen-avista_out].[Baugruppen-ID]
    )
    LEFT JOIN [Artikel-all_avista_out] ON Mastertab.[Artikel-ID 1] = [Artikel-all_avista_out].[Artikel-ID]
  )
  LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
WHERE
  (
    (
      (Mastertab.extern) = Yes
    )
  );
