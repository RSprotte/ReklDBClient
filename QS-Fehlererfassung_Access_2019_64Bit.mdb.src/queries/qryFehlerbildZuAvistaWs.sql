SELECT
  LQWorkstation.ID,
  LQWorkstation.WSAvistaID,
  ZuordnungWsEinzeltaetigkeit.EinzelTaetigkeitMelderID,
  get_IntegerfromAccessTrueFalse([StatusMelder]) AS Aktiv
FROM
  ZuordnungWsEinzeltaetigkeit
  INNER JOIN LQWorkstation ON ZuordnungWsEinzeltaetigkeit.WsID = LQWorkstation.ID
WHERE
  (
    (
      (
        ZuordnungWsEinzeltaetigkeit.EinzelTaetigkeitMelderID
      ) > 0
    )
  );
