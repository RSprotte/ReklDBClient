SELECT
  tblHost.HostID,
  tblHost.HostName,
  tblHost.KZ,
  tblHost.Datum,
  tblHost.ClientVersion,
  tblHost.Angemeldet
FROM
  tblHost
WHERE
  (
    (
      (tblHost.Datum) = Date()
    )
    AND (
      (tblHost.Angemeldet) = True
    )
  );
