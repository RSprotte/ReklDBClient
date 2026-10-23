SELECT
  RSP_Lokalisierung.[Text-ID],
  RSP_Lokalisierung.TypeOfObject,
  RSP_Lokalisierung.OjectID_Name,
  RSP_Lokalisierung.TextID,
  RSP_Lokalisierung.ÜbersetztCN,
  RSP_Lokalisierung.BasisText,
  RSP_Lokalisierung.LängeText,
  RSP_Lokalisierung.TextCN
FROM
  RSP_Lokalisierung
WHERE
  (
    (
      (RSP_Lokalisierung.ÜbersetztCN) = False
    )
  );
