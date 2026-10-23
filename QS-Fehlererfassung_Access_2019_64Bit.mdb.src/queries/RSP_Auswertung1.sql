SELECT
  Mastertab.ID,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.[Kategorie-Problem],
  Mastertab.extern,
  Mastertab.[Menge betroffen - Kunde],
  Mastertab.[Menge betroffen - LQ],
  Mastertab.Fehlerpositionen,
  Mastertab.Auftragsdatum,
  Mastertab.Verursacher,
  Mastertab.[Kategorie-Fehleranalyse],
  LQEinheit.LQEinheit,
  LQBereich.LQBereich,
  tblFirma.Firma_Name,
  Mastertab.TaetigkeitID,
  Mastertab.LQEinheitID,
  Lieferanten.LieferantenverantwortungLQ
FROM
  (
    (
      (
        (
          Mastertab
          LEFT JOIN LQEinheit ON Mastertab.LQEinheitID = LQEinheit.ID
        )
        LEFT JOIN LQBereich ON Mastertab.LQBereichID = LQBereich.ID
      )
      LEFT JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]
    )
    LEFT JOIN Lieferanten ON Mastertab.[Lieferanten-Code 1] = Lieferanten.[Lieferant-ID]
  )
  LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID;
