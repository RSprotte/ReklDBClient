SELECT
  Mastertab.Auftragsdatum,
  Mastertab.ID,
  Mastertab.[Meldungsnummer Kunde],
  Mastertab.Rückmeldung_Kunde_Verschiebung,
  Mastertab.Rückmeldung_Kunde_Ist,
  Mastertab.[Lieferanten-Code 1]
FROM
  Mastertab
ORDER BY
  Mastertab.ID DESC;
