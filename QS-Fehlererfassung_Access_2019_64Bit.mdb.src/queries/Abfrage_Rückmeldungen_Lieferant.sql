SELECT
  Mastertab.ID,
  Lieferanten.[Lieferant-ID],
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.Rückmeldung_Lieferant_Soll,
  Mastertab.Rückmeldung_Lieferant_Ist,
  Mastertab.[Status: erledigt]
FROM
  Lieferanten
  INNER JOIN Mastertab ON Lieferanten.[Lieferant-ID] = Mastertab.[Lieferanten-Code 1]
WHERE
  (
    (
      (
        Mastertab.Rückmeldung_Lieferant_Soll
      ) = Yes
    )
    AND (
      (
        Mastertab.Rückmeldung_Lieferant_Ist
      ) Is Null
    )
    AND (
      (Mastertab.Auftragsdatum) Like "*2010"
    )
  )
  OR (
    (
      (Mastertab.Auftragsdatum) Like "*2009"
    )
  )
ORDER BY
  Mastertab.Rückmeldung_Kunde_Soll;
