SELECT
  tblProtokoll.Vorgang_ID,
  tblProtokoll.Datum,
  tblProtokoll.KZZ,
  tblProtokoll.Tabelle,
  tblProtokoll.Feld,
  tblProtokoll.Alter_Wert,
  tblProtokoll.Neuer_Wert
FROM
  tblProtokoll
WHERE
  (
    (
      (tblProtokoll.Tabelle) <> ""
    )
  )
ORDER BY
  tblProtokoll.Datum DESC,
  tblProtokoll.Feld,
  tblProtokoll.Datum;
