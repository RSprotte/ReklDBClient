SELECT
  Mastertab.Auftragsdatum,
  Mastertab.ID,
  Mastertab.[Meldungsnummer Kunde],
  tblFirma.Firma_Name,
  Bearbeiter.Kurzzeichen AS [KZ-Service],
  Sachbearbeiter.Kurzzeichen AS [KZ-Sachbearbeiter],
  [Kategorien-Vorgangstatus].Anzeige AS Status_Vorgang,
  Mastertab.Rückmeldung_Kunde_Verschiebung,
  Taetigkeiten.Anzeige AS [Taetigkeit-Anzeige],
  ZuordnungDefTaetigkeit.Anzeige AS [Fehlerbild-Anzeige],
  [Kategorien-Problemstatus].Anzeige AS Qualifizierung,
  Lieferanten.Lieferantenname,
  Mastertab.Rückmeldung_Kunde_Ist,
  Verursacher.Anzeige AS [Verursacher-Anzeige]
FROM
  (
    (
      (
        (
          (
            (
              (
                (
                  (
                    Mastertab
                    LEFT JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]
                  )
                  LEFT JOIN Bearbeiter ON Mastertab.[Zuständiger Service-MA] = Bearbeiter.Name
                )
                LEFT JOIN Sachbearbeiter ON Mastertab.Sachbearbeiter = Sachbearbeiter.Name
              )
              LEFT JOIN Taetigkeiten ON Mastertab.TaetigkeitID = Taetigkeiten.ID
            )
            LEFT JOIN ZuordnungDefTaetigkeit ON Mastertab.DefTaetigkeitID = ZuordnungDefTaetigkeit.ID
          )
          LEFT JOIN Lieferanten ON Mastertab.[Lieferanten-Code 1] = Lieferanten.[Lieferant-ID]
        )
        LEFT JOIN [Kategorien-Vorgangstatus] ON Mastertab.[Kategorie-Status] = [Kategorien-Vorgangstatus].[Kategorie-Vorgangstatus]
      )
      LEFT JOIN Verursacher ON Mastertab.Verursacher = Verursacher.Verursacher
    )
    LEFT JOIN [Kategorien-Problemstatus] ON Mastertab.[Kategorie-Problem] = [Kategorien-Problemstatus].[Kategorie-Problemstatus]
  )
  LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
ORDER BY
  Mastertab.ID DESC;
