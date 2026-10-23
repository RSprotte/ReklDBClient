SELECT
  tblFolgeStatusAbhängig.IDAktiveStatus,
  [Kategorien-Vorgangstatus].Anzeige,
  tblFolgeStatusAbhängig.IDFolgeStatus,
  [Kategorien-Vorgangstatus_1].Anzeige,
  tblFolgeStatusAbhängig.IDVorgangsTyp,
  Vorgangstyp.Anzeige
FROM
  (
    (
      tblFolgeStatusAbhängig
      LEFT JOIN [Kategorien-Vorgangstatus] AS [Kategorien-Vorgangstatus_1] ON tblFolgeStatusAbhängig.IDFolgeStatus = [Kategorien-Vorgangstatus_1].ID
    )
    LEFT JOIN [Kategorien-Vorgangstatus] ON tblFolgeStatusAbhängig.IDAktiveStatus = [Kategorien-Vorgangstatus].ID
  )
  LEFT JOIN Vorgangstyp ON tblFolgeStatusAbhängig.IDVorgangsTyp = Vorgangstyp.VorgangstypID;
