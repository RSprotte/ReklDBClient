PARAMETERS Anfangsdatum DateTime,
Enddatum DateTime,
Firma Text (255);
SELECT
  Mastertab.ID,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.Rückmeldung_Kunde_Soll AS AnKdSoll,
  Mastertab.Rückmeldung_Kunde_Ist AS AnKdIst,
  Mastertab.Rückmeldung_Kunde_Verschiebung AS Verschoben,
  Mastertab.Rückmeldung_Lieferant_Soll AS VonLiefSoll,
  Mastertab.Rückmeldung_Lieferant_Ist AS VonLiefIst,
  DateDiff(
    "y",
    (
      Now()
    ),
    [Rückmeldung_Kunde_Soll]
  ) AS Rückstand,
  tblFirma.Firma_Name,
  [Kategorien-Vorgangstatus].[Kategorie-Vorgangstatus],
  Mastertab.[QS-Kunde] AS QS
FROM
  (
    [Kategorien-Vorgangstatus]
    INNER JOIN (
      Kunden
      RIGHT JOIN Mastertab ON Kunden.[Kunden-ID] = Mastertab.[Kunden-ID]
    ) ON [Kategorien-Vorgangstatus].[Kategorie-Vorgangstatus] = Mastertab.[Kategorie-Status]
  )
  LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
WHERE
  (
    (
      (
        Mastertab.[Meldungsnummer Kunde]
      ) Is Not Null
    )
    AND (
      (
        Mastertab.Rückmeldung_Kunde_Soll
      ) = Yes
    )
    AND (
      (tblFirma.Firma_Name) = [Firma]
    )
    AND (
      (Mastertab.[QS-Kunde]) = No
    )
    AND (
      (Mastertab.[Status: erledigt]) = No
    )
    AND (
      (Mastertab.extern) = Yes
    )
    AND (
      (Mastertab.Auftragsdatum) Between [Anfangsdatum]
      And [Enddatum]
    )
    AND (
      (Mastertab.Fehlergruppe) Not Like "K"
    )
  )
  OR (
    (
      (
        Mastertab.[Meldungsnummer Kunde]
      ) Is Not Null
    )
    AND (
      (
        Mastertab.Rückmeldung_Kunde_Soll
      ) = Yes
    )
    AND (
      (tblFirma.Firma_Name) = [Firma]
    )
    AND (
      (Mastertab.[QS-Kunde]) = No
    )
    AND (
      (Mastertab.[Status: erledigt]) = No
    )
    AND (
      (Mastertab.extern) = Yes
    )
    AND (
      (Mastertab.Auftragsdatum) Between [Anfangsdatum]
      And [Enddatum]
    )
    AND (
      (Mastertab.Fehlergruppe) Is Null
    )
  )
ORDER BY
  Mastertab.Rückmeldung_Kunde_Soll;
