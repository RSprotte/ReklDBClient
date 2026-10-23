SELECT
  tblTicketListe.TicketID,
  tblTicketListe.MailVon AS VON,
  tblTicketListe.MailAn AS AN,
  tblTicketListe.MailCC AS CC,
  tblTicketListe.MailBCC AS BCC,
  tblTicketListe.MailBetreff AS Betreff,
  tblTicketListe.MailText AS Body,
  tblTicketListe.TicketTypID,
  tblTicketListe.LinkZuVorgangsOrdner AS Link,
  tblTicketListe.ReklID
FROM
  tblTicketListe;
