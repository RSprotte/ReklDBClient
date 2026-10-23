SELECT
  Mastertab.ID,
  LagerplatzArtikel(
    Nz([BG], "0"),
    Nz([Var Nr], "0")
  ) AS LagerBG,
  LagerplatzArtikel(
    Nz([ArtNr], "0"),
    Nz([Variante], "0")
  ) AS LagerArt,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.[Bestellnummer Kunde],
  Kunden.Kontaktperson,
  tblFirma.Firma_Name,
  [Baugruppen-avista_out].[LQ Art-Nr] AS BG,
  [Baugruppen-avista_out].[Kunden-Nr],
  [Artikel-all_avista_out].[LQ Art-Nr] AS ArtNr,
  [Artikel-all_avista_out].[Kd Mat-Nr],
  Mastertab.[Rückliefer-Menge] AS Menge,
  Vorgangstyp.Anzeige
FROM
  (
    (
      (
        (
          (
            Kunden
            INNER JOIN Mastertab ON Kunden.[Kunden-ID] = Mastertab.[Kunden-ID]
          )
          INNER JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
        )
        LEFT JOIN [Artikel-all_avista_out] ON Mastertab.[Artikel-ID 1] = [Artikel-all_avista_out].[Artikel-ID]
      )
      LEFT JOIN [Baugruppen-avista_out] ON Mastertab.[Baugruppen-ID] = [Baugruppen-avista_out].[Baugruppen-ID]
    )
    LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID
  )
  LEFT JOIN Vorgangstyp ON ReklMemo.VorgangstypID = Vorgangstyp.VorgangstypID
WHERE
  (
    (
      (Mastertab.[WE-Datum]) Is Null
    )
    AND (
      (Mastertab.[Umbau/Reparatur]) = True
    )
    AND (
      (Mastertab.Auftragsdatum) > #1/1/2021#
    )
  )
GROUP BY
  Mastertab.ID,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.[Bestellnummer Kunde],
  Kunden.Kontaktperson,
  tblFirma.Firma_Name,
  [Baugruppen-avista_out].[LQ Art-Nr],
  [Baugruppen-avista_out].[Kunden-Nr],
  [Artikel-all_avista_out].[LQ Art-Nr],
  [Artikel-all_avista_out].[Kd Mat-Nr],
  Mastertab.[Rückliefer-Menge],
  Vorgangstyp.Anzeige,
  [Artikel-all_avista_out].Variante,
  [Baugruppen-avista_out].[Var Nr]
ORDER BY
  Mastertab.ID DESC,
  Vorgangstyp.Anzeige;
