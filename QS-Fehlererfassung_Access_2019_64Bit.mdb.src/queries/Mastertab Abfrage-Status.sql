SELECT
  Mastertab.[Welches Teil ist betroffen],
  Mastertab.[LINK Materialnummer] AS Ausdr1,
  Mastertab.[Welcher Fehler tritt auf],
  Mastertab.[Wie Warum passierte der Fehler],
  Mastertab.Auftragsdatum,
  Mastertab.ID,
  Mastertab.[Hersteller-Teilenummer],
  tblFirma.Firma_Name,
  Mastertab.[Welche Baugruppe ist betroffen] AS Ausdr3,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.Seriennummer,
  Mastertab.[Sofortmaßnahmen (Aktion)],
  Mastertab.[Name des Kontakts] AS Ausdr4,
  Mastertab.[Kategorie-Fehleranalyse],
  Mastertab.[Kategorie-Problem]
FROM
  (
    Mastertab
    LEFT JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]
  )
  LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
WHERE
  (
    (
      (Mastertab.Auftragsdatum) Like "*.2008"
    )
    And (
      (tblFirma.Firma_Name) = "tchw"
    )
    And (
      (
        Mastertab.[Welche Baugruppe ist betroffen]
      ) Is Not Null
    )
  );
