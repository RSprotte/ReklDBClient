SELECT
  Mastertab.ID,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.Rückmeldung_Kunde_Verschiebung AS [Fällig am],
  LQEinheit.LQEinheit,
  Mastertab.Verursacher,
  tblFirma.Firma_Name AS Firma
FROM
  (
    (
      (
        (
          Mastertab
          LEFT JOIN LQEinheit ON Mastertab.LQEinheitID = LQEinheit.ID
        )
        LEFT JOIN ReklMemo ON Mastertab.ID = ReklMemo.ID
      )
      LEFT JOIN Vorgangstyp ON ReklMemo.VorgangstypID = Vorgangstyp.VorgangstypID
    )
    LEFT JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]
  )
  LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
WHERE
  (
    (
      (Mastertab.Auftragsdatum) Between #1/1/2015#
      And Date()
    )
    AND (
      (
        Mastertab.Rückmeldung_Kunde_Ist
      ) Is Null
    )
    AND (
      (Mastertab.extern) = Yes
    )
    AND (
      (Vorgangstyp.Vorgangstyp) = "Meldung"
    )
  )
GROUP BY
  Mastertab.ID,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.Rückmeldung_Kunde_Verschiebung,
  LQEinheit.LQEinheit,
  Mastertab.Verursacher,
  tblFirma.Firma_Name
HAVING
  (
    (
      (
        Mastertab.[Meldungsnummer Kunde]
      ) Is Not Null
    )
    AND (
      (
        Mastertab.Rückmeldung_Kunde_Verschiebung
      ) Between #1/1/2015#
      And Date() - 1
    )
    AND (
      (tblFirma.Firma_Name) Like "T*"
    )
  );
