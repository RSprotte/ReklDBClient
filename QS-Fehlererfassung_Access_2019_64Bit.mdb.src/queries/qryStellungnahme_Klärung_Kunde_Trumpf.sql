SELECT
  Mastertab.ID AS P4T,
  Mastertab.ID,
  Mastertab.Auftragsdatum,
  Mastertab.Rückmeldung_Kunde_Ist,
  Mastertab.[Meldungsnummer Kunde],
  tblFirma.Firma_Name,
  Kunden.Kontaktperson,
  [Kategorien-Vorgangstatus].Anzeige,
  ReklMemo.Portal
FROM
  (
    (
      (
        Mastertab
        LEFT JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]
      )
      LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
    )
    LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID
  )
  LEFT JOIN [Kategorien-Vorgangstatus] ON Mastertab.[Kategorie-Status] = [Kategorien-Vorgangstatus].[Kategorie-Vorgangstatus]
WHERE
  (
    (
      (Mastertab.extern) = True
    )
    AND (
      (tblFirma.Zurodnung_ID) = 4
    )
  )
GROUP BY
  Mastertab.ID,
  Mastertab.Auftragsdatum,
  Mastertab.Rückmeldung_Kunde_Ist,
  Mastertab.[Meldungsnummer Kunde],
  tblFirma.Firma_Name,
  Kunden.Kontaktperson,
  [Kategorien-Vorgangstatus].Anzeige,
  ReklMemo.Portal
HAVING
  (
    (
      (Mastertab.Auftragsdatum) > DateAdd(
        "d",
        -730,
        Now()
      )
    )
    AND (
      (
        [Kategorien-Vorgangstatus].Anzeige
      ) = "Stellungnahme Klärung Kunde"
    )
  )
  OR (
    (
      (
        [Kategorien-Vorgangstatus].Anzeige
      ) = "Gutschrift / Ersatz Lieferant"
    )
  )
  OR (
    (
      (
        [Kategorien-Vorgangstatus].Anzeige
      ) = "Rücklieferung Kunde"
    )
  )
ORDER BY
  Mastertab.Auftragsdatum DESC;
