SELECT
  Sachbearbeiter.[SB-ID],
  Sachbearbeiter.EmailName AS MailVon,
  Sachbearbeiter.EmailName AS MailTo,
  LQEinheit.MailCCLQEinheit AS MailCC,
  LQEinheit.MailEskalationLQEinheit AS MailBCC,
  Sachbearbeiter.Kurzzeichen
FROM
  Sachbearbeiter
  INNER JOIN LQEinheit ON Sachbearbeiter.Einheit = LQEinheit.LQEinheit;
