SELECT
  Mastertab.ID,
  Mastertab.[Meldungsnummer Kunde],
  Last(tblLager.Lagerbezeichnung) AS Werk,
  Vorgangstyp.Anzeige
FROM
  (
    (
      (
        (
          Kunden
          INNER JOIN Mastertab ON Kunden.[Kunden-ID] = Mastertab.[Kunden-ID]
        )
        INNER JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
      )
      LEFT JOIN tblLager ON Mastertab.LQEinheitID = tblLager.LQEinheitID
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
      (Mastertab.Auftragsdatum) > #1/1/2023#
    )
  )
GROUP BY
  Mastertab.ID,
  Mastertab.[Meldungsnummer Kunde],
  Vorgangstyp.Anzeige
ORDER BY
  Vorgangstyp.Anzeige;
