SELECT
  Mastertab.ID,
  LQEinheit.LQEinheit,
  LQBereich.LQBereich,
  LQEinheit.MailTOLQEinheit AS ANE,
  LQEinheit.MailCCLQEinheit AS CCE,
  LQEinheit.MailEskalationLQEinheit AS BCCE,
  Mastertab.[Welcher Fehler tritt auf] AS Fehlerbeschreibung,
  Bearbeiter.[E-Mail] AS VON,
  tblFirma.Firma_Name,
  LQBereich.MailTOLQBereich AS ANB,
  LQBereich.MailCCLQBereich AS CCB,
  LQBereich.MailEskalationLQBereich AS BCCB,
  LieferantAnsprechpartner.Email AS ANLIEF,
  Sachbearbeiter.EmailName AS CCSACHB,
  Sachbearbeiter.TelefonBeruflich AS Telefon,
  Mastertab.Maschinennummer
FROM
  (
    (
      (
        (
          (
            (
              Mastertab
              LEFT JOIN LQEinheit ON Mastertab.LQEinheitID = LQEinheit.ID
            )
            INNER JOIN Kunden ON Mastertab.[Kunden-ID] = Kunden.[Kunden-ID]
          )
          LEFT JOIN tblFirma ON Kunden.FirmaID = tblFirma.Firma_ID
        )
        LEFT JOIN Bearbeiter ON Mastertab.[Zuständiger Service-MA] = Bearbeiter.Name
      )
      LEFT JOIN LQBereich ON Mastertab.LQBereichID = LQBereich.ID
    )
    LEFT JOIN Sachbearbeiter ON Mastertab.Sachbearbeiter = Sachbearbeiter.Name
  )
  LEFT JOIN LieferantAnsprechpartner ON Mastertab.[LiefAnsprechp-ID] = LieferantAnsprechpartner.AnsprechpartnerID;
