SELECT
  tblAccessWsID_AvistaWsID.AvistaWsID,
  tblAccessWsID_AvistaWsID.AvistaWsName,
  tblAccessWsID_AvistaWsID.AvistaGrName,
  tblAccessWsID_AvistaWsID.AvistaGrBez,
  LQWorkstation.ID,
  LQWorkstation.LQWorkstation,
  LQBereich.LQBereich
FROM
  tblAccessWsID_AvistaWsID
  INNER JOIN (
    LQWorkstation
    INNER JOIN LQBereich ON LQWorkstation.[ID-LQBereich] = LQBereich.ID
  ) ON tblAccessWsID_AvistaWsID.AvistaWsID = LQWorkstation.WSAvistaID
ORDER BY
  tblAccessWsID_AvistaWsID.AvistaWsID;
